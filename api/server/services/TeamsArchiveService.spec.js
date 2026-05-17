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

  it('uses sender fallback for advanced search when participant metadata is missing', async () => {
    searchTeamsMemoryChunks.mockResolvedValue({
      retrievalMode: 'enterprise_memory',
      results: [],
    });
    db.findTeamsArchiveConversations
      .mockResolvedValueOnce([])
      .mockResolvedValueOnce([
        {
          _id: 'conv-adv-fallback',
          graphChatId: 'chat-adv-fallback',
          chatType: 'oneOnOne',
          topic: 'Rachel infrastructure',
          participants: [],
          firstMessageAt: new Date('2026-02-01T00:00:00.000Z'),
          lastMessageAt: new Date('2026-03-01T00:00:00.000Z'),
          messageCount: 8,
          syncStatus: 'complete',
        },
      ])
      .mockResolvedValueOnce([
        {
          graphChatId: 'chat-adv-fallback',
          chatType: 'oneOnOne',
          topic: 'Rachel infrastructure',
          participants: [],
        },
      ]);
    db.findTeamsArchiveMessages
      .mockResolvedValueOnce([
        {
          _id: 'sender-hit',
          graphMessageId: 'sender-hit-g',
          graphChatId: 'chat-adv-fallback',
          fromDisplayName: 'Rachel Steele',
          fromEmail: 'rachel@example.com',
          bodyPreview: 'Discussing infrastructure updates',
          bodyText: 'Discussing infrastructure updates',
          sentDateTime: new Date('2026-03-01T00:00:00.000Z'),
        },
      ])
      .mockResolvedValueOnce([
        {
          _id: 'adv-hit',
          graphMessageId: 'adv-hit-g',
          graphChatId: 'chat-adv-fallback',
          fromDisplayName: 'Test User',
          fromEmail: 'user@example.com',
          bodyPreview: 'Infrastructure review with Rachel',
          bodyText: 'Infrastructure review with Rachel',
          sentDateTime: new Date('2026-03-02T00:00:00.000Z'),
        },
      ]);

    const result = await TeamsArchiveService.advancedSearchMessages(user, {
      topic: 'infrastructure',
      participants: ['Rachel Steele'],
      chatType: 'oneOnOne',
      limit: 4,
    });

    expect(result).toMatchObject({
      retrievalMode: 'advanced_message_previews',
      resultCount: 1,
      chatType: 'oneOnOne',
      topic: 'infrastructure',
      resolvedConversation: {
        graphChatId: 'chat-adv-fallback',
      },
    });
    expect(result.results[0]).toMatchObject({
      graphChatId: 'chat-adv-fallback',
      graphMessageId: 'adv-hit-g',
    });
  });

  it('lists scoped conversations with participant, topic, chat type, and daysBack filters', async () => {
    db.findTeamsArchiveConversations.mockResolvedValue([
      {
        _id: 'conv-1',
        graphChatId: 'chat-10',
        chatType: 'oneOnOne',
        topic: 'Quarterly review',
        participants: [{ displayName: 'Manager', email: 'manager@example.com' }],
        lastMessageAt: new Date('2026-05-01T12:00:00.000Z'),
        updatedAt: new Date('2026-05-01T12:00:00.000Z'),
        messageCount: 24,
      },
    ]);

    const result = await TeamsArchiveService.listConversations(user, {
      participants: ['Manager'],
      topic: 'review',
      chatType: 'oneOnOne',
      daysBack: 90,
      limit: 10,
    });

    expect(db.findTeamsArchiveConversations).toHaveBeenCalledWith(
      expect.objectContaining({
        user: 'user-1',
        chatType: 'oneOnOne',
        lastMessageAt: expect.any(Object),
        $and: expect.any(Array),
      }),
      expect.objectContaining({
        limit: 5,
        offset: 0,
      }),
    );
    expect(result).toMatchObject({
      retrievalMode: 'conversation_list',
      chatType: 'oneOnOne',
      topic: 'review',
      daysBack: 90,
      participants: ['Manager'],
    });
    expect(result.conversations).toHaveLength(1);
    expect(result.conversations[0]).toMatchObject({
      graphChatId: 'chat-10',
      chatType: 'oneOnOne',
      topic: 'Quarterly review',
    });
  });

  it('returns disambiguation candidates for conversation dossiers when multiple chats match', async () => {
    db.findTeamsArchiveConversations.mockResolvedValue([
      {
        _id: 'conv-a',
        graphChatId: 'chat-a',
        chatType: 'oneOnOne',
        topic: 'Rachel weekly sync',
        participants: [{ displayName: 'Rachel', email: 'rachel@example.com' }],
        firstMessageAt: new Date('2026-01-01T00:00:00.000Z'),
        lastMessageAt: new Date('2026-05-01T00:00:00.000Z'),
        messageCount: 40,
        syncStatus: 'complete',
      },
      {
        _id: 'conv-b',
        graphChatId: 'chat-b',
        chatType: 'oneOnOne',
        topic: 'Rachel project thread',
        participants: [{ displayName: 'Rachel', email: 'rachel@example.com' }],
        firstMessageAt: new Date('2025-11-01T00:00:00.000Z'),
        lastMessageAt: new Date('2026-04-15T00:00:00.000Z'),
        messageCount: 22,
        syncStatus: 'complete',
      },
    ]);

    const result = await TeamsArchiveService.getConversationDossier(user, {
      participants: ['Rachel'],
      chatType: 'oneOnOne',
    });

    expect(result).toMatchObject({
      retrievalMode: 'conversation_dossier',
      resolved: false,
      candidateCount: 2,
      chatType: 'oneOnOne',
      participants: ['Rachel'],
    });
    expect(result.candidates).toHaveLength(2);
    expect(result.candidates[0]).toMatchObject({
      graphChatId: 'chat-a',
      topic: 'Rachel weekly sync',
    });
  });

  it('uses sender fallback when participant metadata is missing from conversations', async () => {
    db.findTeamsArchiveConversations
      .mockResolvedValueOnce([])
      .mockResolvedValueOnce([
        {
          _id: 'conv-fallback',
          graphChatId: 'chat-fallback',
          chatType: 'oneOnOne',
          topic: 'Legacy 1:1',
          participants: [],
          firstMessageAt: new Date('2026-01-10T00:00:00.000Z'),
          lastMessageAt: new Date('2026-04-10T00:00:00.000Z'),
          messageCount: 3,
          syncStatus: 'complete',
        },
      ]);
    db.findTeamsArchiveMessages
      .mockResolvedValueOnce([
        {
          _id: 'm-fallback-1',
          graphMessageId: 'gm-fallback-1',
          graphChatId: 'chat-fallback',
          fromDisplayName: 'Rachel Steele',
          fromEmail: 'rachel@example.com',
          bodyPreview: 'Following up on staffing',
          bodyText: 'Following up on staffing',
          sentDateTime: new Date('2026-04-10T00:00:00.000Z'),
        },
      ])
      .mockResolvedValueOnce([
        {
          _id: 'm-fallback-1',
          graphMessageId: 'gm-fallback-1',
          graphChatId: 'chat-fallback',
          fromDisplayName: 'Rachel Steele',
          fromEmail: 'rachel@example.com',
          bodyPreview: 'Following up on staffing',
          bodyText: 'Following up on staffing',
          sentDateTime: new Date('2026-04-10T00:00:00.000Z'),
        },
        {
          _id: 'm-fallback-2',
          graphMessageId: 'gm-fallback-2',
          graphChatId: 'chat-fallback',
          fromDisplayName: 'Test User',
          fromEmail: 'user@example.com',
          bodyPreview: 'Here is the update',
          bodyText: 'Here is the update',
          sentDateTime: new Date('2026-04-11T00:00:00.000Z'),
        },
      ]);
    db.countTeamsArchiveMessages.mockResolvedValue(2);

    const result = await TeamsArchiveService.getConversationDossier(user, {
      participants: ['Rachel Steele'],
      chatType: 'oneOnOne',
    });

    expect(db.findTeamsArchiveMessages).toHaveBeenNthCalledWith(
      1,
      expect.objectContaining({
        user: 'user-1',
        $or: expect.any(Array),
      }),
      expect.objectContaining({
        limit: 2000,
      }),
    );
    expect(result).toMatchObject({
      retrievalMode: 'conversation_dossier',
      resolved: true,
      archiveBacked: true,
      chat: {
        graphChatId: 'chat-fallback',
      },
      completeness: {
        loadedAllMessages: true,
        totalMessagesInScope: 2,
      },
    });
  });

  it('returns completeness metadata for resolved conversation dossiers', async () => {
    db.findTeamsArchiveConversations.mockResolvedValue([
      {
        _id: 'conv-dossier',
        graphChatId: 'chat-dossier',
        chatType: 'group',
        topic: 'Design review',
        participants: [{ displayName: 'Lead', email: 'lead@example.com' }],
        firstMessageAt: new Date('2026-02-01T00:00:00.000Z'),
        lastMessageAt: new Date('2026-02-03T00:00:00.000Z'),
        messageCount: 3,
        syncStatus: 'complete',
      },
    ]);
    db.countTeamsArchiveMessages.mockResolvedValue(3);
    db.findTeamsArchiveMessages.mockResolvedValue([
      {
        _id: 'msg-a',
        graphMessageId: 'g-a',
        graphChatId: 'chat-dossier',
        fromDisplayName: 'Lead',
        fromEmail: 'lead@example.com',
        bodyPreview: 'Kickoff agenda',
        bodyText: 'Kickoff agenda',
        sentDateTime: new Date('2026-02-01T10:00:00.000Z'),
      },
      {
        _id: 'msg-b',
        graphMessageId: 'g-b',
        graphChatId: 'chat-dossier',
        fromDisplayName: 'Test User',
        fromEmail: 'user@example.com',
        bodyPreview: 'Design review topic',
        bodyText: 'Design review topic',
        sentDateTime: new Date('2026-02-02T10:00:00.000Z'),
      },
      {
        _id: 'msg-c',
        graphMessageId: 'g-c',
        graphChatId: 'chat-dossier',
        fromDisplayName: 'Lead',
        fromEmail: 'lead@example.com',
        bodyPreview: 'Action items',
        bodyText: 'Action items',
        sentDateTime: new Date('2026-02-03T10:00:00.000Z'),
      },
    ]);

    const result = await TeamsArchiveService.getConversationDossier(user, {
      chatId: 'chat-dossier',
      topic: 'design',
    });

    expect(result).toMatchObject({
      retrievalMode: 'conversation_dossier',
      resolved: true,
      archiveBacked: true,
      completeness: {
        loadedAllMessages: true,
        loadedMessages: 3,
        totalMessagesInScope: 3,
        truncated: false,
      },
      matchedMessages: 1,
    });
    expect(result.highlights.length).toBeGreaterThan(0);
  });

  it('returns a bounded message window around the anchor message', async () => {
    db.findTeamsArchiveConversations
      .mockResolvedValueOnce([{ graphChatId: 'chat-window' }])
      .mockResolvedValueOnce([
        {
          graphChatId: 'chat-window',
          topic: 'Window test',
          chatType: 'group',
          participants: [{ displayName: 'Lead', email: 'lead@example.com' }],
        },
      ]);
    db.findTeamsArchiveMessages
      .mockResolvedValueOnce([
        {
          _id: 'anchor',
          graphMessageId: 'anchor-g',
          graphChatId: 'chat-window',
          fromDisplayName: 'Lead',
          fromEmail: 'lead@example.com',
          bodyPreview: 'Anchor text',
          bodyText: 'Anchor text',
          sentDateTime: new Date('2026-03-02T12:00:00.000Z'),
        },
      ])
      .mockResolvedValueOnce([
        {
          _id: 'before-1',
          graphMessageId: 'before-g',
          graphChatId: 'chat-window',
          fromDisplayName: 'Lead',
          fromEmail: 'lead@example.com',
          bodyPreview: 'Before text',
          bodyText: 'Before text',
          sentDateTime: new Date('2026-03-02T11:00:00.000Z'),
        },
        {
          _id: 'anchor',
          graphMessageId: 'anchor-g',
          graphChatId: 'chat-window',
          fromDisplayName: 'Lead',
          fromEmail: 'lead@example.com',
          bodyPreview: 'Anchor text',
          bodyText: 'Anchor text',
          sentDateTime: new Date('2026-03-02T12:00:00.000Z'),
        },
      ])
      .mockResolvedValueOnce([
        {
          _id: 'anchor',
          graphMessageId: 'anchor-g',
          graphChatId: 'chat-window',
          fromDisplayName: 'Lead',
          fromEmail: 'lead@example.com',
          bodyPreview: 'Anchor text',
          bodyText: 'Anchor text',
          sentDateTime: new Date('2026-03-02T12:00:00.000Z'),
        },
        {
          _id: 'after-1',
          graphMessageId: 'after-g',
          graphChatId: 'chat-window',
          fromDisplayName: 'Test User',
          fromEmail: 'user@example.com',
          bodyPreview: 'After text',
          bodyText: 'After text',
          sentDateTime: new Date('2026-03-02T13:00:00.000Z'),
        },
      ]);

    const result = await TeamsArchiveService.getMessagesWindow(user, {
      chatId: 'chat-window',
      aroundMessageId: 'anchor-g',
      before: 1,
      after: 1,
    });

    expect(result).toMatchObject({
      retrievalMode: 'message_window',
      chatId: 'chat-window',
      graphChatId: 'chat-window',
      anchorGraphMessageId: 'anchor-g',
    });
    expect(result.messages).toHaveLength(3);
    expect(result.messages.map((message) => message.graphMessageId)).toEqual([
      'before-g',
      'anchor-g',
      'after-g',
    ]);
  });

  it('summarizes a conversation with matched message counts and highlights', async () => {
    db.findTeamsArchiveConversations
      .mockResolvedValueOnce([{ graphChatId: 'chat-summary' }])
      .mockResolvedValueOnce([
        {
          graphChatId: 'chat-summary',
          topic: 'Kubernetes troubleshooting',
          chatType: 'oneOnOne',
          participants: [{ displayName: 'Manager', email: 'manager@example.com' }],
        },
      ]);
    db.findTeamsArchiveMessages.mockResolvedValue([
      {
        _id: 'sum-1',
        graphMessageId: 'sum-g-1',
        graphChatId: 'chat-summary',
        fromDisplayName: 'Manager',
        fromEmail: 'manager@example.com',
        bodyPreview: 'Need cluster help',
        bodyText: 'Need cluster help',
        sentDateTime: new Date('2026-01-01T10:00:00.000Z'),
      },
      {
        _id: 'sum-2',
        graphMessageId: 'sum-g-2',
        graphChatId: 'chat-summary',
        fromDisplayName: 'Test User',
        fromEmail: 'user@example.com',
        bodyPreview: 'kubectl get pods -A',
        bodyText: 'kubectl get pods -A',
        sentDateTime: new Date('2026-01-02T10:00:00.000Z'),
      },
    ]);

    const result = await TeamsArchiveService.summarizeConversation(user, {
      chatId: 'chat-summary',
      topic: 'kubectl',
      limit: 2,
    });

    expect(result).toMatchObject({
      retrievalMode: 'conversation_summary',
      chatId: 'chat-summary',
      query: 'kubectl',
      totalMessages: 2,
      matchedMessages: 1,
    });
    expect(result.highlights).toHaveLength(1);
    expect(result.highlights[0]).toMatchObject({
      graphMessageId: 'sum-g-2',
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
