jest.mock(
  '@librechat/api',
  () => ({
    isEnabled: (value) => value === true || value === 'true' || value === '1',
  }),
  { virtual: true },
);

jest.mock(
  '@librechat/data-schemas',
  () => ({
    logger: {
      warn: jest.fn(),
      info: jest.fn(),
      error: jest.fn(),
    },
    runAsSystem: async (fn) => fn(),
  }),
  { virtual: true },
);

jest.mock('~/server/services/GraphTokenService', () => ({
  getGraphApiToken: jest.fn(),
}));

jest.mock('~/server/services/EnterpriseMemory/teamsProjection', () => ({
  projectTeamsArchiveSyncToMemory: jest.fn(),
}));

jest.mock('~/server/services/EnterpriseMemory/retrieval', () => ({
  searchTeamsMemoryChunks: jest.fn(),
}));

jest.mock('~/models', () => ({
  countTeamsArchiveConversations: jest.fn(),
  countTeamsArchiveMessages: jest.fn(),
  findLatestTeamsArchiveSyncJob: jest.fn(),
  findLatestEnterpriseMemoryJob: jest.fn(),
  getTeamsArchiveBackfillState: jest.fn(),
  countEnterpriseMemoryChunks: jest.fn(),
  countEnterpriseMemoryEntities: jest.fn(),
  countDistinctEnterpriseMemoryChunkField: jest.fn(),
  countActiveTeamsArchiveSyncLeases: jest.fn(),
  findTeamsArchiveMessages: jest.fn(),
  findTeamsArchiveConversations: jest.fn(),
  updateTeamsArchiveSyncJob: jest.fn(),
}));

const db = require('~/models');
const { searchTeamsMemoryChunks } = require('~/server/services/EnterpriseMemory/retrieval');
const TeamsArchiveService = require('./TeamsArchiveService');

describe('TeamsArchiveService', () => {
  const originalEnv = process.env;
  const user = {
    id: 'user-1',
    provider: 'openid',
    openidId: 'entra-user-1',
    email: 'user@example.com',
    name: 'Test User',
    username: 'test.user',
    federatedTokens: {
      access_token: 'openid-token',
    },
  };

  beforeEach(() => {
    jest.clearAllMocks();
    process.env = {
      ...originalEnv,
      TEAMS_ARCHIVE_ENABLED: 'true',
      OPENID_REUSE_TOKENS: 'true',
    };

    db.findLatestTeamsArchiveSyncJob.mockResolvedValue(null);
    db.findLatestEnterpriseMemoryJob.mockResolvedValue(null);
    db.getTeamsArchiveBackfillState.mockResolvedValue(null);
    db.countEnterpriseMemoryChunks.mockResolvedValue(0);
    db.countEnterpriseMemoryEntities.mockResolvedValue(0);
    db.countDistinctEnterpriseMemoryChunkField.mockResolvedValue(0);
    db.countActiveTeamsArchiveSyncLeases.mockResolvedValue(0);
    db.findTeamsArchiveMessages.mockResolvedValue([]);
    db.findTeamsArchiveConversations.mockResolvedValue([]);
    searchTeamsMemoryChunks.mockResolvedValue({ results: [] });
  });

  afterEach(() => {
    process.env = originalEnv;
  });

  it('reports projected conversation coverage separately from searchable chunk coverage', async () => {
    const latestSync = {
      _id: 'sync-1',
      status: 'completed',
      mode: 'chats',
      phase: 'complete',
      checkpoint: { page: 4 },
      stats: { discovered: 1022 },
      requestedChatLimit: 5000,
      requestedMessagesPerChat: 1000,
      discoveredChatCount: 1022,
      processedChatCount: 1008,
      skippedChatCount: 0,
      projectionJobId: 'projection-1',
      conversationCount: 1022,
      messageCount: 66517,
      startedAt: new Date('2026-05-17T06:00:00.000Z'),
      completedAt: new Date('2026-05-17T06:30:00.000Z'),
    };

    db.countTeamsArchiveConversations.mockResolvedValue(1022);
    db.countTeamsArchiveMessages.mockResolvedValue(66517);
    db.findLatestTeamsArchiveSyncJob.mockImplementation((filter) =>
      Promise.resolve(filter?.status === 'running' ? null : latestSync),
    );
    db.findLatestEnterpriseMemoryJob.mockResolvedValue({
      _id: 'projection-1',
      status: 'success',
      startedAt: new Date('2026-05-17T06:31:00.000Z'),
      completedAt: new Date('2026-05-17T06:32:00.000Z'),
      stats: {
        projectedConversationCount: 780,
        chunkCount: 1122,
      },
    });
    db.getTeamsArchiveBackfillState.mockResolvedValue({
      status: 'syncing',
      discoveryComplete: true,
      nextChatPageLink: null,
      discoveredChatCount: 1022,
      completedChatCount: 1008,
      pendingChatCount: 12,
      runningChatCount: 0,
      failedChatCount: 2,
      totalMessageCount: 66517,
      lastSyncJobId: 'sync-1',
      lastProjectionJobId: 'projection-1',
      lastDiscoveredAt: new Date('2026-05-17T06:05:00.000Z'),
      lastCompletedAt: new Date('2026-05-17T06:30:00.000Z'),
      lastHeartbeatAt: new Date('2026-05-17T06:30:00.000Z'),
      errorMessage: null,
    });
    db.countEnterpriseMemoryChunks.mockResolvedValue(1122);
    db.countEnterpriseMemoryEntities.mockResolvedValue(780);
    db.countDistinctEnterpriseMemoryChunkField.mockResolvedValue(776);
    db.countActiveTeamsArchiveSyncLeases.mockResolvedValue(1);

    const result = await TeamsArchiveService.getStatus(user);

    expect(result.projectionCoverage).toEqual({
      indexedConversationCount: 780,
      totalConversationCount: 1022,
      indexedChunkCount: 1122,
      searchableConversationCount: 776,
      pendingConversationCount: 242,
      fullyIndexed: false,
      coveragePercent: 76.3,
    });
    expect(result.backfillState).toMatchObject({
      status: 'paused',
      completedChatCount: 1008,
      pendingChatCount: 12,
      failedChatCount: 2,
    });
    expect(result.latestSync).toMatchObject({
      id: 'sync-1',
      status: 'completed',
      requestedChatLimit: 5000,
      requestedMessagesPerChat: 1000,
    });
    expect(result.latestProjection).toMatchObject({
      id: 'projection-1',
      status: 'success',
      stats: {
        projectedConversationCount: 780,
        chunkCount: 1122,
      },
    });
  });

  it('falls back to archived recent messages when enterprise memory returns no results', async () => {
    searchTeamsMemoryChunks.mockResolvedValue({
      retrievalMode: 'enterprise_memory',
      results: [],
    });
    db.findTeamsArchiveMessages.mockResolvedValue([
      {
        _id: 'msg-1',
        graphMessageId: 'graph-msg-1',
        graphChatId: 'chat-1',
        fromUserId: 'entra-user-1',
        fromDisplayName: 'Test User',
        fromEmail: 'user@example.com',
        bodyPreview: 'Terraform module update',
        bodyText: 'Terraform module update for network policies',
        sentDateTime: new Date('2026-05-01T12:00:00.000Z'),
      },
    ]);
    db.findTeamsArchiveConversations.mockResolvedValue([
      {
        graphChatId: 'chat-1',
        topic: 'Infrastructure',
        chatType: 'group',
        participants: [{ displayName: 'Test User', email: 'user@example.com' }],
      },
    ]);

    const result = await TeamsArchiveService.recentMessages(user, {
      query: 'terraform',
      daysBack: 30,
      limit: 4,
    });

    expect(searchTeamsMemoryChunks).toHaveBeenCalledWith(
      user,
      expect.objectContaining({
        query: 'terraform',
        daysBack: 30,
        senderScope: 'me',
        sortBy: 'recent',
      }),
    );
    expect(db.findTeamsArchiveMessages).toHaveBeenCalledWith(
      expect.objectContaining({
        user: 'user-1',
        $or: expect.any(Array),
        sentDateTime: expect.any(Object),
        $and: expect.any(Array),
      }),
      expect.objectContaining({
        limit: 4,
      }),
    );
    expect(result).toMatchObject({
      retrievalMode: 'recent_message_previews',
      resultCount: 1,
      query: 'terraform',
    });
  });

  it('falls back to archived advanced search when enterprise memory returns no results', async () => {
    searchTeamsMemoryChunks.mockResolvedValue({
      retrievalMode: 'enterprise_memory',
      results: [],
    });
    db.findTeamsArchiveMessages.mockResolvedValue([
      {
        _id: 'msg-2',
        graphMessageId: 'graph-msg-2',
        graphChatId: 'chat-2',
        fromUserId: 'entra-user-1',
        fromDisplayName: 'Test User',
        fromEmail: 'user@example.com',
        bodyPreview: 'Turkey design review status',
        bodyText: 'Turkey design review status and next milestone details',
        sentDateTime: new Date('2026-04-18T15:00:00.000Z'),
      },
    ]);
    db.findTeamsArchiveConversations.mockResolvedValue([
      {
        graphChatId: 'chat-2',
        topic: 'Turkey design review',
        chatType: 'group',
        participants: [{ displayName: 'Test User', email: 'user@example.com' }],
      },
    ]);

    const result = await TeamsArchiveService.advancedSearchMessages(user, {
      topic: 'Turkey design review',
      senderScope: 'me',
      limit: 4,
    });

    expect(searchTeamsMemoryChunks).toHaveBeenCalledWith(
      user,
      expect.objectContaining({
        topic: 'Turkey design review',
        senderScope: 'me',
        limit: 4,
      }),
    );
    expect(db.findTeamsArchiveMessages).toHaveBeenCalledWith(
      expect.objectContaining({
        user: 'user-1',
        $or: expect.any(Array),
        $and: expect.any(Array),
      }),
      expect.objectContaining({
        limit: 4,
      }),
    );
    expect(result).toMatchObject({
      retrievalMode: 'advanced_message_previews',
      resultCount: 1,
      topic: 'Turkey design review',
      senderScope: 'me',
    });
  });

  it('returns the full archived message body when a preview was truncated', async () => {
    db.findTeamsArchiveMessages.mockImplementation((filter) => {
      if (filter?.graphMessageId === 'graph-msg-3') {
        return Promise.resolve([
          {
            _id: 'msg-3',
            graphMessageId: 'graph-msg-3',
            graphChatId: 'chat-3',
            fromDisplayName: 'Test User',
            fromEmail: 'user@example.com',
            bodyPreview: 'kubectl get pods ...',
            bodyText: 'kubectl get pods -A\nkubectl describe pod api-0 -n cortex\nkubectl logs api-0',
            bodyContentType: 'text',
            sentDateTime: new Date('2026-03-01T10:00:00.000Z'),
            attachments: [],
            mentions: [],
            webUrl: 'https://teams.example/messages/graph-msg-3',
          },
        ]);
      }

      if (filter?._id === 'graph-msg-3') {
        return Promise.resolve([]);
      }

      return Promise.resolve([]);
    });
    db.findTeamsArchiveConversations.mockResolvedValue([
      {
        graphChatId: 'chat-3',
        topic: 'Kubernetes troubleshooting',
        chatType: 'oneOnOne',
        participants: [{ displayName: 'Manager', email: 'manager@example.com' }],
      },
    ]);

    const result = await TeamsArchiveService.getMessageBody(user, {
      messageId: 'graph-msg-3',
    });

    expect(result).toMatchObject({
      retrievalMode: 'message_body',
      resolved: true,
      message: {
        graphMessageId: 'graph-msg-3',
        graphChatId: 'chat-3',
        topic: 'Kubernetes troubleshooting',
        chatType: 'oneOnOne',
        bodyPreview: 'kubectl get pods ...',
        bodyText:
          'kubectl get pods -A\nkubectl describe pod api-0 -n cortex\nkubectl logs api-0',
        previewWasTruncated: true,
      },
    });
    expect(result.message.bodyTextLength).toBeGreaterThan(result.message.previewLength);
  });
});
