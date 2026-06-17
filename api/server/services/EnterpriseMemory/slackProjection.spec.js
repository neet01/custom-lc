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
  findSlackArchiveConversations: jest.fn(),
  findSlackArchiveMessages: jest.fn(),
  upsertEnterpriseMemoryEntity: jest.fn(),
  bulkUpsertEnterpriseMemoryRelationships: jest.fn(),
  bulkUpsertEnterpriseMemoryChunks: jest.fn(),
}));

const db = require('~/models');
const { logger } = require('@librechat/data-schemas');
const { projectSlackArchiveSyncToMemory } = require('./slackProjection');

describe('slackProjection', () => {
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
    db.findSlackArchiveConversations
      .mockResolvedValueOnce([])
      .mockResolvedValueOnce([
        {
          slackConversationId: 'CEMPTY',
          conversationType: 'public_channel',
          name: 'empty-room',
          messageCount: 0,
          participants: [],
        },
      ])
      .mockResolvedValueOnce([
        {
          slackConversationId: 'CTEXTLESS',
          conversationType: 'private_channel',
          name: 'ops-room',
          messageCount: 2,
          participants: [],
        },
      ]);

    db.findSlackArchiveMessages
      .mockResolvedValueOnce([])
      .mockResolvedValueOnce([
        {
          slackMessageTs: '1714521600.000100',
          normalizedText: '',
          text: '',
          isChunkable: false,
          skipChunkReason: 'system_like',
          mentions: [],
          attachments: [],
          sentAt: new Date('2026-05-01T00:00:00.000Z'),
        },
        {
          slackMessageTs: '1714608000.000200',
          normalizedText: '',
          text: '',
          isChunkable: false,
          skipChunkReason: 'empty_text',
          mentions: [],
          attachments: [],
          sentAt: new Date('2026-05-02T00:00:00.000Z'),
        },
      ]);

    const result = await projectSlackArchiveSyncToMemory({
      userId: 'user-1',
      tenantId: 'tenant-1',
      syncJobId: 'sync-1',
      slackConversationIds: ['CMISSING', 'CEMPTY', 'CTEXTLESS'],
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
        totalMessagesLoaded: 2,
        totalChunkableMessages: 0,
        totalSkippedEmptyTextMessages: 2,
        projectionMessageFetchLimit: 5000,
        zeroChunkReasonCounts: {
          system_like: 1,
          empty_text: 1,
        },
        searchableConversationCountsByType: {
          public_channel: 0,
          private_channel: 0,
          im: 0,
          mpim: 0,
          unknown: 0,
        },
      },
    });

    expect(logger.warn).toHaveBeenCalledWith(
      '[EnterpriseMemory] Slack projection skipped missing archived conversation',
      expect.objectContaining({
        slackConversationId: 'CMISSING',
      }),
    );
    expect(logger.info).toHaveBeenCalledWith(
      '[EnterpriseMemory] Slack projection completed',
      expect.objectContaining({
        projectionDiagnostics: expect.objectContaining({
          zeroChunkConversationCount: 2,
          totalSkippedEmptyTextMessages: 2,
        }),
      }),
    );
  });

  it('creates message and conversation window chunks for searchable Slack content', async () => {
    db.findSlackArchiveConversations.mockResolvedValue([
      {
        slackConversationId: 'CWINDOW',
        teamId: 'T1',
        conversationType: 'im',
        name: '',
        topic: '',
        purpose: '',
        messageCount: 2,
        participants: [{ slackUserId: 'U1', displayName: 'Manager', email: 'manager@example.com' }],
      },
    ]);
    db.findSlackArchiveMessages.mockResolvedValue([
      {
        slackConversationId: 'CWINDOW',
        slackMessageTs: '1714521600.000100',
        slackUserId: 'U1',
        displayName: 'Manager',
        normalizedText: 'First update',
        text: 'First update',
        isChunkable: true,
        mentions: [],
        attachments: [],
        files: [],
        sentAt: new Date('2026-05-01T00:00:00.000Z'),
      },
      {
        slackConversationId: 'CWINDOW',
        slackMessageTs: '1714522200.000200',
        slackUserId: 'U2',
        displayName: 'Analyst',
        normalizedText: 'Second update with <@U1>',
        text: 'Second update with <@U1>',
        isChunkable: true,
        mentions: [{ slackUserId: 'U1', displayName: 'Manager' }],
        attachments: [],
        files: [],
        sentAt: new Date('2026-05-01T00:10:00.000Z'),
      },
    ]);

    const result = await projectSlackArchiveSyncToMemory({
      userId: 'user-1',
      tenantId: 'tenant-1',
      syncJobId: 'sync-2',
      slackConversationIds: ['CWINDOW'],
    });

    expect(result).toMatchObject({
      status: 'success',
      projectedConversationCount: 1,
      chunkCount: 3,
      projectionDiagnostics: {
        zeroChunkConversationCount: 0,
        searchableConversationCountsByType: {
          im: 1,
        },
      },
    });

    expect(db.bulkUpsertEnterpriseMemoryChunks).toHaveBeenCalledTimes(1);
    const chunkRecords = db.bulkUpsertEnterpriseMemoryChunks.mock.calls[0][0];
    expect(chunkRecords).toHaveLength(3);
    expect(chunkRecords[0]).toMatchObject({
      source: 'slack',
      sourceRecordType: 'slack_message',
      sourceParentRecordId: 'CWINDOW',
      chunkType: 'message',
      text: 'First update',
    });
    expect(chunkRecords[2]).toMatchObject({
      source: 'slack',
      sourceRecordType: 'slack_conversation',
      sourceParentRecordId: 'CWINDOW',
      chunkType: 'conversation_window',
    });
  });
});
