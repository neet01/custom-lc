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
  upsertTeamsArchiveBackfillState: jest.fn(),
  upsertTeamsArchiveConversation: jest.fn(),
  bulkUpsertTeamsArchiveMessages: jest.fn(),
  createTeamsArchiveSyncJob: jest.fn(),
  updateTeamsArchiveSyncJob: jest.fn(),
  updateTeamsArchiveConversation: jest.fn(),
  acquireTeamsArchiveSyncLease: jest.fn(),
  refreshTeamsArchiveSyncLease: jest.fn(),
  releaseTeamsArchiveSyncLease: jest.fn(),
  countEnterpriseMemoryChunks: jest.fn(),
  countEnterpriseMemoryEntities: jest.fn(),
  countDistinctEnterpriseMemoryChunkField: jest.fn(),
  countActiveTeamsArchiveSyncLeases: jest.fn(),
  findTeamsArchiveMessages: jest.fn(),
  findTeamsArchiveConversations: jest.fn(),
  findTeamsArchiveSyncJobById: jest.fn(),
}));

const db = require('~/models');
const { getGraphApiToken } = require('~/server/services/GraphTokenService');
const { projectTeamsArchiveSyncToMemory } = require('~/server/services/EnterpriseMemory/teamsProjection');
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
    db.upsertTeamsArchiveBackfillState.mockImplementation(async (record) => record);
    db.upsertTeamsArchiveConversation.mockImplementation(async (record) => ({
      _id: record.graphChatId || 'conv-1',
      ...record,
    }));
    db.bulkUpsertTeamsArchiveMessages.mockResolvedValue(0);
    db.createTeamsArchiveSyncJob.mockResolvedValue({ _id: 'sync-1', id: 'sync-1', status: 'running' });
    db.updateTeamsArchiveSyncJob.mockImplementation(async (_id, updates) => ({
      _id: 'sync-1',
      id: 'sync-1',
      ...updates,
    }));
    db.updateTeamsArchiveConversation.mockImplementation(async (id, updates) => ({
      _id: id,
      id,
      ...updates,
    }));
    db.acquireTeamsArchiveSyncLease.mockResolvedValue({ leaseKey: 'lease-1' });
    db.refreshTeamsArchiveSyncLease.mockResolvedValue({ leaseKey: 'lease-1' });
    db.releaseTeamsArchiveSyncLease.mockResolvedValue(true);
    db.countEnterpriseMemoryChunks.mockResolvedValue(0);
    db.countEnterpriseMemoryEntities.mockResolvedValue(0);
    db.countDistinctEnterpriseMemoryChunkField.mockResolvedValue(0);
    db.countActiveTeamsArchiveSyncLeases.mockResolvedValue(0);
    db.findTeamsArchiveMessages.mockResolvedValue([]);
    db.findTeamsArchiveConversations.mockResolvedValue([]);
    db.findTeamsArchiveSyncJobById.mockResolvedValue(null);
    projectTeamsArchiveSyncToMemory.mockResolvedValue({ status: 'success' });
    searchTeamsMemoryChunks.mockResolvedValue({ results: [] });
    db.countTeamsArchiveConversations.mockResolvedValue(0);
    db.countTeamsArchiveMessages.mockResolvedValue(0);
    getGraphApiToken.mockResolvedValue({ access_token: 'graph-token-1' });
    global.fetch = jest.fn();
  });

  afterEach(() => {
    process.env = originalEnv;
    delete global.fetch;
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

    db.countTeamsArchiveConversations.mockImplementation((filter) => {
      if (filter?.participantDegraded) {
        return Promise.resolve(18);
      }
      return Promise.resolve(1022);
    });
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
        projectionDiagnostics: {
          missingConversationCount: 0,
          zeroMessageConversationCount: 150,
          zeroChunkConversationCount: 200,
          truncatedConversationCount: 4,
          totalMessagesLoaded: 58000,
          totalChunkableMessages: 54120,
          totalSkippedEmptyTextMessages: 3880,
          projectionMessageFetchLimit: 5000,
          zeroChunkReasonCounts: {
            empty_normalized_text: 140,
            system_like_message: 60,
          },
          searchableConversationCountsByChatType: {
            oneOnOne: 320,
            group: 280,
            meeting: 176,
          },
          participantDegradedConversationCount: 18,
        },
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
    db.countDistinctEnterpriseMemoryChunkField
      .mockResolvedValueOnce(776)
      .mockResolvedValueOnce(320)
      .mockResolvedValueOnce(280)
      .mockResolvedValueOnce(176);
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
      projectionDiagnostics: {
        missingConversationCount: 0,
        zeroMessageConversationCount: 150,
        zeroChunkConversationCount: 200,
        truncatedConversationCount: 4,
        totalMessagesLoaded: 58000,
        totalChunkableMessages: 54120,
        totalSkippedEmptyTextMessages: 3880,
        projectionMessageFetchLimit: 5000,
        zeroChunkReasonCounts: {
          empty_normalized_text: 140,
          system_like_message: 60,
        },
        searchableConversationCountsByChatType: {
          oneOnOne: 320,
          group: 280,
          meeting: 176,
        },
        participantDegradedConversationCount: 18,
        chunkableMessageRate: 93.3,
        skippedEmptyTextRate: 6.7,
      },
    });
    expect(result.searchabilityDiagnostics).toEqual({
      discoveredConversationCount: 1022,
      archivedConversationCount: 1022,
      projectedConversationCount: 780,
      searchableConversationCount: 776,
      zeroChunkConversationCount: 200,
      zeroChunkReasonCounts: {
        empty_normalized_text: 140,
        system_like_message: 60,
      },
      searchableConversationCountsByChatType: {
        oneOnOne: 320,
        group: 280,
        meeting: 176,
      },
      participantDegradedConversationCount: 18,
    });
  });

  it('paginates chat members across multiple Graph pages during discovery', async () => {
    global.fetch.mockImplementation(async (url) => {
      const href = String(url);

      if (href.includes('/me/chats')) {
        return {
          ok: true,
          status: 200,
          json: async () => ({
            value: [
              {
                id: 'chat-1',
                chatType: 'group',
                topic: 'Member pagination',
                lastUpdatedDateTime: '2026-05-01T10:00:00.000Z',
              },
            ],
          }),
        };
      }

      if (
        href.includes('/chats/chat-1/members') &&
        (href.includes('$top=50') || href.includes('%24top=50'))
      ) {
        return {
          ok: true,
          status: 200,
          json: async () => ({
            value: [{ displayName: 'Alex', email: 'alex@example.com', userId: 'user-alex' }],
            '@odata.nextLink': 'https://graph.microsoft.us/v1.0/chats/chat-1/members?page=2',
          }),
        };
      }

      if (href.includes('/chats/chat-1/members?page=2')) {
        return {
          ok: true,
          status: 200,
          json: async () => ({
            value: [{ displayName: 'Blair', email: 'blair@example.com', userId: 'user-blair' }],
          }),
        };
      }

      if (href.includes('/chats/chat-1/messages')) {
        return {
          ok: true,
          status: 200,
          json: async () => ({
            value: [],
          }),
        };
      }

      throw new Error(`Unhandled fetch URL: ${href}`);
    });

    db.findTeamsArchiveConversations.mockImplementation((filter) => {
      if (filter?.graphChatId?.$in) {
        return Promise.resolve([]);
      }
      if (filter?.syncStatus) {
        return Promise.resolve([
          {
            _id: 'conv-1',
            id: 'conv-1',
            graphChatId: 'chat-1',
            chatType: 'group',
            participants: [],
          },
        ]);
      }
      return Promise.resolve([]);
    });
    db.countTeamsArchiveMessages.mockResolvedValue(0);

    await TeamsArchiveService.syncUserArchive(user, {
      chatLimit: 1,
      messagesPerChat: 5,
    });

    expect(db.upsertTeamsArchiveConversation).toHaveBeenCalledWith(
      expect.objectContaining({
        graphChatId: 'chat-1',
        participants: expect.arrayContaining([
          expect.objectContaining({
            displayName: 'Alex',
            source: 'graph',
            confidence: 'high',
          }),
          expect.objectContaining({
            displayName: 'Blair',
            source: 'graph',
            confidence: 'high',
          }),
        ]),
        participantMetadataSource: 'graph',
      }),
    );
  });

  it('falls back to sender-derived participants when Graph member enrichment fails', async () => {
    global.fetch.mockImplementation(async (url) => {
      const href = String(url);

      if (href.includes('/me/chats')) {
        return {
          ok: true,
          status: 200,
          json: async () => ({
            value: [
              {
                id: 'chat-fallback',
                chatType: 'oneOnOne',
                topic: 'Fallback participants',
                lastUpdatedDateTime: '2026-05-01T10:00:00.000Z',
              },
            ],
          }),
        };
      }

      if (href.includes('/chats/chat-fallback/members')) {
        return {
          ok: false,
          status: 403,
          statusText: 'Forbidden',
          json: async () => ({ error: { message: 'InsufficientPrivileges' } }),
        };
      }

      if (href.includes('/chats/chat-fallback/messages')) {
        return {
          ok: true,
          status: 200,
          json: async () => ({
            value: [
              {
                id: 'graph-msg-fallback',
                messageType: 'message',
                from: {
                  user: {
                    id: 'user-rachel',
                    displayName: 'Rachel Steele',
                    email: 'rachel@example.com',
                  },
                },
                body: {
                  contentType: 'html',
                  content: '<p>Following up on staffing</p>',
                },
                mentions: [],
                createdDateTime: '2026-05-01T11:00:00.000Z',
              },
            ],
          }),
        };
      }

      throw new Error(`Unhandled fetch URL: ${href}`);
    });

    db.findTeamsArchiveConversations.mockImplementation((filter) => {
      if (filter?.graphChatId?.$in) {
        return Promise.resolve([]);
      }
      if (filter?.syncStatus) {
        return Promise.resolve([
          {
            _id: 'conv-fallback',
            id: 'conv-fallback',
            graphChatId: 'chat-fallback',
            chatType: 'oneOnOne',
            participants: [],
            participantDegraded: true,
          },
        ]);
      }
      return Promise.resolve([]);
    });
    db.countTeamsArchiveMessages.mockResolvedValue(1);

    await TeamsArchiveService.syncUserArchive(user, {
      chatLimit: 1,
      messagesPerChat: 5,
    });

    expect(db.updateTeamsArchiveConversation).toHaveBeenLastCalledWith(
      'conv-fallback',
      expect.objectContaining({
        participants: expect.arrayContaining([
          expect.objectContaining({
            displayName: 'Rachel Steele',
            email: 'rachel@example.com',
            userId: 'user-rachel',
            source: 'inferred_from_messages',
          }),
        ]),
        participantDegraded: true,
        participantMetadataSource: 'inferred_from_messages',
      }),
    );
  });

  it('stores richer normalized message fields and preserves HTML structure for searchability', async () => {
    global.fetch.mockImplementation(async (url) => {
      const href = String(url);

      if (href.includes('/me/chats')) {
        return {
          ok: true,
          status: 200,
          json: async () => ({
            value: [
              {
                id: 'chat-html',
                chatType: 'group',
                topic: 'HTML normalization',
                lastUpdatedDateTime: '2026-05-01T10:00:00.000Z',
              },
            ],
          }),
        };
      }

      if (href.includes('/chats/chat-html/members')) {
        return {
          ok: true,
          status: 200,
          json: async () => ({
            value: [],
          }),
        };
      }

      if (href.includes('/chats/chat-html/messages')) {
        return {
          ok: true,
          status: 200,
          json: async () => ({
            value: [
              {
                id: 'graph-msg-html',
                messageType: 'message',
                from: {
                  user: {
                    id: 'user-lead',
                    displayName: 'Lead',
                    email: 'lead@example.com',
                  },
                },
                body: {
                  contentType: 'html',
                  content:
                    '<p>Hello&nbsp;team<br>Agenda:</p><ul><li>First item</li><li>Second item</li></ul><table><tr><td>Owner</td><td>Alex</td></tr></table><p><a href="https://example.com/doc">Spec</a></p>',
                },
                mentions: [],
                createdDateTime: '2026-05-01T11:00:00.000Z',
              },
            ],
          }),
        };
      }

      throw new Error(`Unhandled fetch URL: ${href}`);
    });

    db.findTeamsArchiveConversations.mockImplementation((filter) => {
      if (filter?.graphChatId?.$in) {
        return Promise.resolve([]);
      }
      if (filter?.syncStatus) {
        return Promise.resolve([
          {
            _id: 'conv-html',
            id: 'conv-html',
            graphChatId: 'chat-html',
            chatType: 'group',
            participants: [],
          },
        ]);
      }
      return Promise.resolve([]);
    });
    db.countTeamsArchiveMessages.mockResolvedValue(1);

    await TeamsArchiveService.syncUserArchive(user, {
      chatLimit: 1,
      messagesPerChat: 5,
    });

    expect(db.bulkUpsertTeamsArchiveMessages).toHaveBeenCalledWith([
      expect.objectContaining({
        graphMessageId: 'graph-msg-html',
        bodyText: expect.stringContaining('Hello team'),
        normalizedTextLength: expect.any(Number),
        isSystemLikeMessage: false,
        isChunkable: true,
        skipChunkReason: '',
      }),
    ]);
    const storedRecord = db.bulkUpsertTeamsArchiveMessages.mock.calls[0][0][0];
    expect(storedRecord.bodyText).toContain('Agenda:');
    expect(storedRecord.bodyText).toContain('- First item');
    expect(storedRecord.bodyText).toContain('Owner | Alex');
    expect(storedRecord.bodyText).toContain('Spec (https://example.com/doc)');
    expect(storedRecord.normalizedTextLength).toBeGreaterThan(20);
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
        $and: expect.any(Array),
      }),
      expect.objectContaining({
        limit: 12,
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

  it('searches archive message bodies for broad topic queries even when conversation topic does not match', async () => {
    searchTeamsMemoryChunks.mockResolvedValue({
      retrievalMode: 'enterprise_memory',
      results: [],
    });
    db.findTeamsArchiveMessages.mockResolvedValue([
      {
        _id: 'msg-topic-body',
        graphMessageId: 'graph-topic-body',
        graphChatId: 'chat-topic-body',
        fromUserId: 'entra-user-1',
        fromDisplayName: 'Test User',
        fromEmail: 'user@example.com',
        bodyPreview: 'Condor thermal issue',
        bodyText: 'Condor thermal issue and heat shield rework details',
        sentDateTime: new Date('2026-04-20T15:00:00.000Z'),
      },
    ]);
    db.findTeamsArchiveConversations.mockResolvedValue([
      {
        graphChatId: 'chat-topic-body',
        topic: 'Random social thread',
        chatType: 'group',
        participants: [{ displayName: 'Test User', email: 'user@example.com' }],
      },
    ]);

    const result = await TeamsArchiveService.advancedSearchMessages(user, {
      topic: 'condor thermal',
      limit: 4,
    });

    expect(db.findTeamsArchiveMessages.mock.calls[0][0]).toEqual(
      expect.objectContaining({
        user: 'user-1',
        $and: expect.any(Array),
      }),
    );
    expect(db.findTeamsArchiveMessages.mock.calls[0][0].graphChatId).toBeUndefined();
    expect(result).toMatchObject({
      retrievalMode: 'advanced_message_previews',
      resultCount: 1,
      topic: 'condor thermal',
      trace: expect.objectContaining({
        archiveUnionRan: true,
        archiveResultCount: 1,
      }),
    });
    expect(result.results[0]).toMatchObject({
      graphChatId: 'chat-topic-body',
      graphMessageId: 'graph-topic-body',
      topic: 'Random social thread',
    });
  });

  it('does not let weak non-empty memory results suppress archive fallback', async () => {
    searchTeamsMemoryChunks.mockResolvedValue({
      retrievalMode: 'enterprise_memory',
      results: [
        {
          id: 'chunk-weak-1',
          graphMessageId: 'memory-graph-1',
          graphChatId: 'chat-memory',
          topic: 'Turkey review',
          chatType: 'group',
          participants: [{ displayName: 'Lead', email: 'lead@example.com' }],
          fromDisplayName: 'Lead',
          fromEmail: 'lead@example.com',
          summary: 'Turkey',
          excerpt: 'Turkey',
          sentDateTime: new Date('2026-04-17T15:00:00.000Z'),
        },
      ],
    });
    db.findTeamsArchiveMessages.mockResolvedValue([
      {
        _id: 'msg-archive-1',
        graphMessageId: 'archive-graph-1',
        graphChatId: 'chat-archive',
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
        graphChatId: 'chat-archive',
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

    expect(result).toMatchObject({
      retrievalMode: 'advanced_message_previews',
      resultCount: 2,
      trace: expect.objectContaining({
        memoryResultCount: 1,
        archiveUnionRan: true,
        archiveResultCount: 1,
      }),
    });
    expect(result.results.map((entry) => entry.graphMessageId)).toEqual([
      'archive-graph-1',
      'memory-graph-1',
    ]);
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

  it('triggers dossier and full-body escalation for completeness-sensitive one-on-one queries', async () => {
    const conversation = {
      _id: 'conv-exact',
      graphChatId: 'chat-exact',
      chatType: 'oneOnOne',
      topic: 'Rachel project thread',
      participants: [{ displayName: 'Rachel', email: 'rachel@example.com' }],
      firstMessageAt: new Date('2026-03-01T00:00:00.000Z'),
      lastMessageAt: new Date('2026-03-03T00:00:00.000Z'),
      messageCount: 2,
      syncStatus: 'complete',
    };
    const truncatedMessage = {
      _id: 'msg-exact-1',
      graphMessageId: 'graph-exact-1',
      graphChatId: 'chat-exact',
      fromDisplayName: 'Rachel',
      fromEmail: 'rachel@example.com',
      bodyPreview: 'Action items ...',
      bodyText: 'Action items: update thermal model, confirm exact wording, send decision memo.',
      sentDateTime: new Date('2026-03-02T10:00:00.000Z'),
      attachments: [],
      mentions: [],
      webUrl: 'https://teams.example/messages/graph-exact-1',
      bodyContentType: 'text',
    };
    const followupMessage = {
      _id: 'msg-exact-2',
      graphMessageId: 'graph-exact-2',
      graphChatId: 'chat-exact',
      fromDisplayName: 'Test User',
      fromEmail: 'user@example.com',
      bodyPreview: 'Will send memo',
      bodyText: 'Will send memo after the thermal review closes.',
      sentDateTime: new Date('2026-03-03T10:00:00.000Z'),
      attachments: [],
      mentions: [],
      webUrl: 'https://teams.example/messages/graph-exact-2',
      bodyContentType: 'text',
    };

    searchTeamsMemoryChunks.mockResolvedValue({
      retrievalMode: 'enterprise_memory',
      results: [],
    });
    db.countTeamsArchiveMessages.mockResolvedValue(2);
    db.findTeamsArchiveConversations.mockImplementation((filter) => {
      if (filter?.graphChatId === 'chat-exact') {
        return Promise.resolve([conversation]);
      }
      if (filter?.chatType === 'oneOnOne') {
        return Promise.resolve([conversation]);
      }
      if (filter?.graphChatId?.$in) {
        return Promise.resolve([conversation]);
      }
      return Promise.resolve([]);
    });
    db.findTeamsArchiveMessages.mockImplementation((filter) => {
      if (filter?.graphMessageId === 'graph-exact-1') {
        return Promise.resolve([truncatedMessage]);
      }
      if (filter?.graphMessageId === 'graph-exact-2') {
        return Promise.resolve([followupMessage]);
      }
      if (filter?._id === 'graph-exact-1' || filter?._id === 'graph-exact-2') {
        return Promise.resolve([]);
      }
      if (filter?.graphChatId?.$in?.includes?.('chat-exact') && filter?.$and) {
        return Promise.resolve([truncatedMessage]);
      }
      if (filter?.graphChatId === 'chat-exact') {
        return Promise.resolve([truncatedMessage, followupMessage]);
      }
      return Promise.resolve([]);
    });

    const result = await TeamsArchiveService.advancedSearchMessages(user, {
      topic: 'exact wording and action items',
      participants: ['Rachel'],
      chatType: 'oneOnOne',
      limit: 4,
    });

    expect(result).toMatchObject({
      retrievalMode: 'advanced_message_previews',
      resultCount: 1,
      resolvedConversation: {
        graphChatId: 'chat-exact',
      },
      conversationDossier: expect.objectContaining({
        retrievalMode: 'conversation_dossier',
        resolved: true,
      }),
      trace: expect.objectContaining({
        conversationDossierRan: true,
        fullBodyEscalationRan: true,
      }),
    });
    expect(result.fullBodies).toHaveLength(1);
    expect(result.fullBodies[0]).toMatchObject({
      retrievalMode: 'message_body',
      resolved: true,
      message: {
        graphMessageId: 'graph-exact-1',
        previewWasTruncated: true,
      },
    });
  });

  it('resolves Graph chat ids without attempting unsafe ObjectId conversation lookup', async () => {
    const graphChatId = '19:meeting_YTY0ZWU3NTEtNWJjNi00NmNkLTgxODAtZjdjMmQxZTEzZDQz@thread.v2';
    const conversation = {
      _id: 'conversation-object-id',
      graphChatId,
      chatType: 'meeting',
      topic: 'IT Eng Standup',
      participants: [{ displayName: 'Test User', email: 'user@example.com' }],
    };

    db.findTeamsArchiveConversations.mockResolvedValue([conversation]);
    db.countTeamsArchiveMessages.mockResolvedValue(1);
    db.findTeamsArchiveMessages.mockResolvedValue([
      {
        _id: 'meeting-message-1',
        graphMessageId: 'meeting-message-graph-1',
        graphChatId,
        fromUserId: 'entra-user-1',
        fromDisplayName: 'Test User',
        fromEmail: 'user@example.com',
        bodyPreview: 'Large meeting update',
        bodyText: 'Large meeting update',
        sentDateTime: new Date('2026-05-20T12:00:00.000Z'),
      },
    ]);

    const result = await TeamsArchiveService.getConversationDossier(user, {
      chatId: graphChatId,
      topic: 'update',
      limit: 4,
    });

    expect(result).toMatchObject({
      retrievalMode: 'conversation_dossier',
      resolved: true,
      chat: {
        graphChatId,
        chatType: 'meeting',
      },
      matchedMessages: 1,
    });
    expect(db.findTeamsArchiveConversations).not.toHaveBeenCalledWith(
      expect.objectContaining({
        _id: graphChatId,
      }),
      expect.anything(),
    );
  });

  it('scopes advanced archive search by Graph chat id when chatId is provided', async () => {
    const graphChatId = '19:meeting_YTY0ZWU3NTEtNWJjNi00NmNkLTgxODAtZjdjMmQxZTEzZDQz@thread.v2';
    searchTeamsMemoryChunks.mockResolvedValue({
      retrievalMode: 'enterprise_memory',
      results: [],
    });
    db.findTeamsArchiveConversations.mockResolvedValue([
      {
        _id: 'conversation-object-id',
        graphChatId,
        chatType: 'meeting',
        topic: 'IT Eng Standup',
        participants: [{ displayName: 'Test User', email: 'user@example.com' }],
      },
    ]);
    db.findTeamsArchiveMessages.mockResolvedValue([
      {
        _id: 'meeting-message-1',
        graphMessageId: 'meeting-message-graph-1',
        graphChatId,
        fromUserId: 'entra-user-1',
        fromDisplayName: 'Test User',
        fromEmail: 'user@example.com',
        bodyPreview: 'Large meeting update',
        bodyText: 'Large meeting update about the telemetry replay work',
        sentDateTime: new Date('2026-05-20T12:00:00.000Z'),
      },
    ]);

    const result = await TeamsArchiveService.advancedSearchMessages(user, {
      chatId: graphChatId,
      topic: 'telemetry replay',
      senderScope: 'me',
      chatType: 'meeting',
      limit: 4,
    });

    expect(db.findTeamsArchiveMessages).toHaveBeenCalledWith(
      expect.objectContaining({
        user: 'user-1',
        graphChatId,
        $or: expect.any(Array),
        $and: expect.any(Array),
      }),
      expect.objectContaining({
        limit: 12,
      }),
    );
    expect(db.findTeamsArchiveConversations).not.toHaveBeenCalledWith(
      expect.objectContaining({
        user: 'user-1',
        chatType: 'meeting',
      }),
      expect.anything(),
    );
    expect(result).toMatchObject({
      retrievalMode: 'advanced_message_previews',
      chatId: graphChatId,
      graphChatId,
      resultCount: 1,
      trace: expect.objectContaining({
        archiveResultCount: 1,
      }),
    });
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
