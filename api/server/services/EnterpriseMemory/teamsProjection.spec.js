jest.mock(
  '@librechat/data-schemas',
  () => ({
    logger: {
      info: jest.fn(),
      warn: jest.fn(),
      error: jest.fn(),
    },
  }),
  { virtual: true },
);

jest.mock('~/models', () => ({
  createEnterpriseMemoryJob: jest.fn(),
  updateEnterpriseMemoryJob: jest.fn(),
  findTeamsArchiveConversations: jest.fn(),
  findTeamsArchiveMessages: jest.fn(),
  upsertEnterpriseMemoryEntity: jest.fn(),
  bulkUpsertEnterpriseMemoryRelationships: jest.fn(),
  bulkUpsertEnterpriseMemoryChunks: jest.fn(),
}));

const db = require('~/models');
const { logger } = require('@librechat/data-schemas');
const { projectTeamsArchiveSyncToMemory } = require('./teamsProjection');

describe('teamsProjection', () => {
  beforeEach(() => {
    jest.clearAllMocks();

    db.createEnterpriseMemoryJob.mockResolvedValue({ _id: 'job-1' });
    db.updateEnterpriseMemoryJob.mockResolvedValue({ _id: 'job-1' });
    db.upsertEnterpriseMemoryEntity.mockImplementation(async (record) => ({
      _id: `${record.entityType}-${record.canonicalKey}`,
      ...record,
    }));
    db.bulkUpsertEnterpriseMemoryRelationships.mockResolvedValue(0);
    db.bulkUpsertEnterpriseMemoryChunks.mockResolvedValue(0);
  });

  it('records projection diagnostics for missing, empty, and zero-chunk conversations', async () => {
    db.findTeamsArchiveConversations
      .mockResolvedValueOnce([])
      .mockResolvedValueOnce([
        {
          graphChatId: 'chat-empty',
          chatType: 'meeting',
          topic: 'Empty meeting',
          messageCount: 0,
          participants: [],
        },
      ])
      .mockResolvedValueOnce([
        {
          graphChatId: 'chat-textless',
          chatType: 'meeting',
          topic: 'System-heavy meeting',
          messageCount: 3,
          participants: [],
        },
      ]);

    db.findTeamsArchiveMessages
      .mockResolvedValueOnce([])
      .mockResolvedValueOnce([
        {
          graphMessageId: 'm-1',
          graphChatId: 'chat-textless',
          bodyText: '',
          summary: '',
          subject: '',
          mentions: [],
          attachments: [],
          sentDateTime: new Date('2026-05-01T00:00:00.000Z'),
        },
        {
          graphMessageId: 'm-2',
          graphChatId: 'chat-textless',
          bodyText: '',
          summary: '',
          subject: '',
          mentions: [],
          attachments: [],
          sentDateTime: new Date('2026-05-02T00:00:00.000Z'),
        },
        {
          graphMessageId: 'm-3',
          graphChatId: 'chat-textless',
          bodyText: '',
          summary: '',
          subject: '',
          mentions: [],
          attachments: [],
          sentDateTime: new Date('2026-05-03T00:00:00.000Z'),
        },
      ]);

    const result = await projectTeamsArchiveSyncToMemory({
      userId: 'user-1',
      tenantId: 'tenant-1',
      syncJobId: 'sync-1',
      graphChatIds: ['missing-chat', 'chat-empty', 'chat-textless'],
    });

    expect(result).toMatchObject({
      status: 'success',
      projectedConversationCount: 2,
      chunkCount: 0,
      projectionDiagnostics: {
        missingConversationCount: 1,
        zeroMessageConversationCount: 1,
        zeroChunkConversationCount: 2,
        truncatedConversationCount: 0,
        totalMessagesLoaded: 3,
        totalChunkableMessages: 0,
        totalSkippedEmptyTextMessages: 3,
        projectionMessageFetchLimit: 5000,
      },
    });

    expect(db.updateEnterpriseMemoryJob).toHaveBeenCalledWith(
      'job-1',
      expect.objectContaining({
        status: 'success',
        stats: expect.objectContaining({
          projectionDiagnostics: expect.objectContaining({
            missingConversationCount: 1,
            zeroChunkConversationCount: 2,
          }),
        }),
      }),
    );

    expect(logger.warn).toHaveBeenCalledWith(
      '[EnterpriseMemory] Teams projection skipped missing archived conversation',
      expect.objectContaining({
        graphChatId: 'missing-chat',
      }),
    );
    expect(logger.info).toHaveBeenCalledWith(
      '[EnterpriseMemory] Teams projection completed',
      expect.objectContaining({
        projectionDiagnostics: expect.objectContaining({
          zeroChunkConversationCount: 2,
          totalSkippedEmptyTextMessages: 3,
        }),
      }),
    );
  });
});
