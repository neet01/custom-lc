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
          isChunkable: false,
          skipChunkReason: 'system_like_message',
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
          isChunkable: false,
          skipChunkReason: 'empty_normalized_text',
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
          isChunkable: false,
          skipChunkReason: 'empty_normalized_text',
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
        zeroChunkReasonCounts: {
          system_like_message: 1,
          empty_normalized_text: 2,
        },
        searchableConversationCountsByChatType: {
          oneOnOne: 0,
          group: 0,
          meeting: 0,
          unknown: 0,
        },
        participantDegradedConversationCount: 0,
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
            zeroChunkReasonCounts: expect.objectContaining({
              empty_normalized_text: 2,
            }),
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
          zeroChunkReasonCounts: expect.objectContaining({
            system_like_message: 1,
            empty_normalized_text: 2,
          }),
        }),
      }),
    );
  });

  it('creates message and conversation_window chunks and tracks searchable chat-type diagnostics', async () => {
    db.findTeamsArchiveConversations.mockResolvedValue([
      {
        graphChatId: 'chat-windowed',
        chatType: 'oneOnOne',
        topic: 'Windowed chat',
        messageCount: 3,
        participantDegraded: true,
        participants: [{ displayName: 'Manager', email: 'manager@example.com', source: 'graph' }],
      },
    ]);
    db.findTeamsArchiveMessages.mockResolvedValue([
      {
        graphMessageId: 'm-1',
        graphChatId: 'chat-windowed',
        fromDisplayName: 'Manager',
        fromEmail: 'manager@example.com',
        fromUserId: 'manager-1',
        bodyText: 'First message',
        bodyPreview: 'First message',
        isChunkable: true,
        mentions: [],
        attachments: [],
        sentDateTime: new Date('2026-05-01T00:00:00.000Z'),
      },
      {
        graphMessageId: 'm-2',
        graphChatId: 'chat-windowed',
        fromDisplayName: 'Test User',
        fromEmail: 'user@example.com',
        fromUserId: 'user-1',
        bodyText: 'Second message',
        bodyPreview: 'Second message',
        isChunkable: true,
        mentions: [],
        attachments: [],
        sentDateTime: new Date('2026-05-01T00:10:00.000Z'),
      },
      {
        graphMessageId: 'm-3',
        graphChatId: 'chat-windowed',
        fromDisplayName: 'Manager',
        fromEmail: 'manager@example.com',
        fromUserId: 'manager-1',
        bodyText: 'Third message',
        bodyPreview: 'Third message',
        isChunkable: true,
        mentions: [],
        attachments: [],
        sentDateTime: new Date('2026-05-01T00:20:00.000Z'),
      },
    ]);

    const result = await projectTeamsArchiveSyncToMemory({
      userId: 'user-1',
      tenantId: 'tenant-1',
      syncJobId: 'sync-2',
      graphChatIds: ['chat-windowed'],
    });

    expect(result).toMatchObject({
      status: 'success',
      projectedConversationCount: 1,
      chunkCount: 4,
      projectionDiagnostics: {
        searchableConversationCountsByChatType: {
          oneOnOne: 1,
          group: 0,
          meeting: 0,
          unknown: 0,
        },
        participantDegradedConversationCount: 1,
      },
    });

    const chunkRecords = db.bulkUpsertEnterpriseMemoryChunks.mock.calls[0][0];
    expect(chunkRecords.filter((record) => record.chunkType === 'message')).toHaveLength(3);
    expect(chunkRecords.filter((record) => record.chunkType === 'conversation_window')).toHaveLength(1);
    expect(chunkRecords.find((record) => record.chunkType === 'conversation_window')).toMatchObject({
      sourceRecordType: 'teams_chat',
      sourceRecordId: 'chat-windowed',
      sourceParentRecordId: 'chat-windowed',
      metadata: expect.objectContaining({
        chatType: 'oneOnOne',
        includedMessageCount: 3,
        messageIds: ['m-1', 'm-2', 'm-3'],
      }),
    });
  });

  it('excludes deferred conversations and marks a partial run without indexing them', async () => {
    db.findTeamsArchiveConversations.mockImplementation(async (filter) => {
      if (filter.graphChatId === 'chat-deferred') {
        throw new Error('deferred conversation must not be queried for projection');
      }
      return [
        {
          graphChatId: 'chat-complete',
          chatType: 'group',
          topic: 'Complete chat',
          messageCount: 1,
          participants: [],
        },
      ];
    });
    db.findTeamsArchiveMessages.mockResolvedValue([
      {
        graphMessageId: 'm-1',
        graphChatId: 'chat-complete',
        fromDisplayName: 'Alice',
        fromEmail: 'alice@example.com',
        bodyText: 'Only message',
        bodyPreview: 'Only message',
        isChunkable: true,
        mentions: [],
        attachments: [],
        sentDateTime: new Date('2026-05-01T00:00:00.000Z'),
      },
    ]);

    const result = await projectTeamsArchiveSyncToMemory({
      userId: 'user-1',
      tenantId: 'tenant-1',
      syncJobId: 'sync-partial',
      graphChatIds: ['chat-complete', 'chat-deferred'],
      runStatus: 'partial',
      deferredGraphChatIds: ['chat-deferred'],
    });

    expect(result).toMatchObject({
      status: 'success',
      projectedConversationCount: 1,
      sourceRunStatus: 'partial',
      sourceRunPartial: true,
      deferredConversationCount: 1,
    });

    const createStats = db.createEnterpriseMemoryJob.mock.calls[0][0].stats;
    expect(createStats).toMatchObject({
      requestedConversationCount: 1,
      sourceRunPartial: true,
      excludedDeferredConversationCount: 1,
    });
    expect(logger.warn).toHaveBeenCalledWith(
      '[EnterpriseMemory] Projecting a partial Teams sync run; deferred conversations excluded',
      expect.objectContaining({ deferredConversationCount: 1, excludedDeferredConversationCount: 1 }),
    );
  });
});
