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
      if (filter?.$or || filter?.meaningfulMessageCount) {
        return Promise.resolve(0);
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
      staleOrIncompleteConversations: 0,
      zeroMeaningfulMessageConversations: 0,
      systemOnlyRecentConversations: 0,
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
        lastMeaningfulMessageAt: expect.any(Object),
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

  it('ranks recent meeting chats by meaningful activity and flags system-only recency', async () => {
    const meaningfulRecent = {
      _id: 'conv-meaningful',
      graphChatId: 'meeting-meaningful',
      chatType: 'meeting',
      topic: 'Enterprise IT Stand Up',
      lastMessageAt: new Date('2026-06-02T10:00:00.000Z'),
      lastMeaningfulMessageAt: new Date('2026-06-02T09:55:00.000Z'),
      lastSystemMessageAt: new Date('2026-06-02T10:00:00.000Z'),
      messageCount: 20,
      meaningfulMessageCount: 12,
      systemMessageCount: 2,
      emptyMessageCount: 1,
    };
    const systemOnlyRecent = {
      _id: 'conv-system',
      graphChatId: 'meeting-system',
      chatType: 'meeting',
      topic: 'Enterprise IT Stand Up',
      lastMessageAt: new Date('2026-06-03T10:00:00.000Z'),
      lastMeaningfulMessageAt: null,
      lastSystemMessageAt: new Date('2026-06-03T10:00:00.000Z'),
      messageCount: 4,
      meaningfulMessageCount: 0,
      systemMessageCount: 4,
      emptyMessageCount: 0,
    };
    db.findTeamsArchiveConversations.mockResolvedValue([meaningfulRecent, systemOnlyRecent]);
    db.findTeamsArchiveMessages.mockResolvedValue([
      {
        _id: 'msg-meaningful',
        graphMessageId: 'g-meaningful',
        graphChatId: 'meeting-meaningful',
        fromDisplayName: 'Test User',
        bodyPreview: 'Latest human update',
        sentDateTime: new Date('2026-06-02T09:55:00.000Z'),
        isChunkable: true,
        normalizedTextLength: 19,
        isSystemLikeMessage: false,
      },
    ]);

    const result = await TeamsArchiveService.recentMeetingChats(user, {
      topic: 'Enterprise IT Stand Up',
      limit: 5,
    });

    expect(db.findTeamsArchiveConversations).toHaveBeenCalledWith(
      expect.objectContaining({
        user: 'user-1',
        chatType: 'meeting',
        topic: expect.any(RegExp),
      }),
      expect.objectContaining({
        sort: { lastMeaningfulMessageAt: -1, lastMessageAt: -1, updatedAt: -1 },
      }),
    );
    expect(result).toMatchObject({
      retrievalMode: 'recent_meeting_chats',
      identityWarning: {
        candidateCount: 2,
      },
    });
    expect(result.conversations.map((conversation) => conversation.graphChatId)).toEqual([
      'meeting-meaningful',
      'meeting-system',
    ]);
    expect(result.conversations[1].warnings).toMatchObject({
      systemOnlyRecentActivity: true,
      noMeaningfulMessages: true,
    });
  });

  it('returns newest human-readable messages first for one resolved conversation', async () => {
    const conversation = {
      _id: 'conv-recent',
      graphChatId: 'meeting-recent',
      chatType: 'meeting',
      topic: 'Enterprise IT Stand Up',
      lastMeaningfulMessageAt: new Date('2026-06-02T10:00:00.000Z'),
    };
    db.findTeamsArchiveConversations
      .mockResolvedValueOnce([conversation])
      .mockResolvedValueOnce([conversation]);
    db.findTeamsArchiveMessages.mockResolvedValue([
      {
        _id: 'recent-2',
        graphMessageId: 'g-recent-2',
        graphChatId: 'meeting-recent',
        fromDisplayName: 'Test User',
        bodyPreview: 'Newest readable update',
        sentDateTime: new Date('2026-06-02T10:00:00.000Z'),
        isChunkable: true,
        normalizedTextLength: 22,
        isSystemLikeMessage: false,
      },
      {
        _id: 'recent-1',
        graphMessageId: 'g-recent-1',
        graphChatId: 'meeting-recent',
        fromDisplayName: 'Lead',
        bodyPreview: 'Older readable update',
        sentDateTime: new Date('2026-06-02T09:00:00.000Z'),
        isChunkable: true,
        normalizedTextLength: 21,
        isSystemLikeMessage: false,
      },
    ]);
    db.countTeamsArchiveMessages.mockResolvedValue(0);

    const result = await TeamsArchiveService.conversationRecentMessages(user, {
      chatId: 'meeting-recent',
      limit: 2,
    });

    expect(db.findTeamsArchiveMessages).toHaveBeenCalledWith(
      expect.objectContaining({
        user: 'user-1',
        graphChatId: 'meeting-recent',
        $and: expect.any(Array),
      }),
      expect.objectContaining({
        sort: { sentDateTime: -1, createdAt: -1 },
      }),
    );
    expect(result).toMatchObject({
      retrievalMode: 'conversation_recent_messages',
      resolved: true,
      selectedConversation: {
        graphChatId: 'meeting-recent',
        selectionReason: 'chatId',
      },
    });
    expect(result.messages.map((message) => message.graphMessageId)).toEqual([
      'g-recent-2',
      'g-recent-1',
    ]);
  });

  it('includes canonical trace metadata for explicit graphChatId lookups', async () => {
    const conversation = {
      _id: 'conv-trace',
      graphChatId: 'meeting-trace',
      chatType: 'meeting',
      topic: 'Traceable meeting',
      lastMeaningfulMessageAt: new Date('2026-06-02T10:00:00.000Z'),
    };
    db.findTeamsArchiveConversations
      .mockResolvedValueOnce([conversation])
      .mockResolvedValueOnce([conversation]);
    db.findTeamsArchiveMessages.mockResolvedValue([
      {
        _id: 'trace-msg',
        graphMessageId: 'g-trace',
        graphChatId: 'meeting-trace',
        bodyPreview: 'Traceable update',
        sentDateTime: new Date('2026-06-02T10:00:00.000Z'),
        isChunkable: true,
        normalizedTextLength: 16,
        isSystemLikeMessage: false,
      },
    ]);
    db.countTeamsArchiveMessages.mockResolvedValue(0);

    const result = await TeamsArchiveService.conversationRecentMessages(user, {
      chatId: 'meeting-trace',
      priorGraphChatId: 'meeting-trace',
      limit: 1,
    });

    expect(result.trace).toMatchObject({
      inputChatId: 'meeting-trace',
      resolvedGraphChatId: 'meeting-trace',
      resolvedArchiveConversationId: 'conv-trace',
      selectedBy: 'graphChatId',
      priorGraphChatId: 'meeting-trace',
      identityChanged: false,
      candidateCount: 1,
      candidateGraphChatIds: ['meeting-trace'],
    });
  });

  it('uses prior graphChatId for follow-up dossier requests instead of rediscovering by title', async () => {
    const priorConversation = {
      _id: 'conv-prior',
      graphChatId: 'meeting-prior',
      chatType: 'meeting',
      topic: 'Enterprise IT Stand Up',
      lastMeaningfulMessageAt: new Date('2026-06-02T10:00:00.000Z'),
      messageCount: 1,
      meaningfulMessageCount: 1,
    };
    db.findTeamsArchiveConversations
      .mockResolvedValueOnce([priorConversation])
      .mockResolvedValueOnce([priorConversation]);
    db.countTeamsArchiveMessages.mockResolvedValue(1);
    db.findTeamsArchiveMessages.mockResolvedValue([
      {
        _id: 'prior-msg',
        graphMessageId: 'g-prior',
        graphChatId: 'meeting-prior',
        fromDisplayName: 'Test User',
        bodyPreview: 'New message in the selected standup',
        bodyText: 'New message in the selected standup',
        sentDateTime: new Date('2026-06-02T10:00:00.000Z'),
        isChunkable: true,
        normalizedTextLength: 35,
        isSystemLikeMessage: false,
      },
    ]);

    const result = await TeamsArchiveService.getConversationDossier(user, {
      query: 'what messages are new?',
      priorGraphChatId: 'meeting-prior',
      priorTopic: 'Enterprise IT Stand Up',
      limit: 2,
    });

    expect(result).toMatchObject({
      retrievalMode: 'conversation_dossier',
      resolved: true,
      selectedConversation: {
        graphChatId: 'meeting-prior',
        selectionReason: 'prior_context',
        identityConfidence: 'high',
      },
    });
    expect(db.findTeamsArchiveConversations).not.toHaveBeenCalledWith(
      expect.objectContaining({
        topic: expect.any(RegExp),
      }),
      expect.anything(),
    );
  });

  it('uses prior graphChatId for sender follow-ups instead of switching to newer same-title meetings', async () => {
    const priorConversation = {
      _id: 'conv-prior-sender',
      graphChatId: 'meeting-prior-sender',
      chatType: 'meeting',
      topic: 'IT Eng Standup',
      lastMeaningfulMessageAt: new Date('2026-05-30T16:25:17.385Z'),
    };
    db.findTeamsArchiveConversations
      .mockResolvedValueOnce([priorConversation])
      .mockResolvedValueOnce([priorConversation]);
    db.findTeamsArchiveMessages.mockResolvedValue([
      {
        _id: 'prior-sender-msg',
        graphMessageId: 'g-prior-sender',
        graphChatId: 'meeting-prior-sender',
        fromUserId: 'sender-1',
        fromDisplayName: 'Test User',
        bodyPreview: 'Prior selected standup update',
        sentDateTime: new Date('2026-05-30T16:25:17.385Z'),
        isChunkable: true,
        normalizedTextLength: 29,
        isSystemLikeMessage: false,
      },
    ]);
    db.countTeamsArchiveMessages.mockResolvedValue(1);

    const result = await TeamsArchiveService.conversationSenderMessages(user, {
      query: 'latest messages from it eng standup',
      priorGraphChatId: 'meeting-prior-sender',
      priorTopic: 'IT Eng Standup',
      senderUserId: 'sender-1',
      limit: 2,
    });

    expect(result).toMatchObject({
      retrievalMode: 'conversation_sender_messages',
      resolved: true,
      graphChatId: 'meeting-prior-sender',
      trace: {
        selectedBy: 'priorGraphChatId',
        priorGraphChatId: 'meeting-prior-sender',
        resolvedGraphChatId: 'meeting-prior-sender',
        identityChanged: false,
      },
      messages: [
        expect.objectContaining({
          graphMessageId: 'g-prior-sender',
        }),
      ],
    });
    expect(db.findTeamsArchiveConversations).not.toHaveBeenCalledWith(
      expect.objectContaining({
        topic: expect.any(RegExp),
      }),
      expect.anything(),
    );
  });

  it('warns instead of silently selecting title-only matches across recurring meetings', async () => {
    db.findTeamsArchiveConversations.mockResolvedValue([
      {
        _id: 'conv-series-a',
        graphChatId: 'meeting-series-a',
        chatType: 'meeting',
        topic: 'Enterprise IT Stand Up',
        lastMeaningfulMessageAt: new Date('2026-06-02T10:00:00.000Z'),
      },
      {
        _id: 'conv-series-b',
        graphChatId: 'meeting-series-b',
        chatType: 'meeting',
        topic: 'Enterprise IT Stand Up',
        lastMeaningfulMessageAt: new Date('2026-05-29T10:00:00.000Z'),
      },
    ]);

    const result = await TeamsArchiveService.getConversationDossier(user, {
      topic: 'Enterprise IT Stand Up',
      chatType: 'meeting',
    });

    expect(result).toMatchObject({
      retrievalMode: 'conversation_dossier',
      resolved: false,
      candidateCount: 2,
      identityWarning: {
        reason: 'multiple_conversations_match_title_or_filters',
        candidateCount: 2,
      },
    });
  });

  it('warns instead of silently selecting sender messages for ambiguous same-title topic lookups', async () => {
    db.findTeamsArchiveConversations.mockResolvedValue([
      {
        _id: 'conv-it-a',
        graphChatId: 'meeting-it-a',
        chatType: 'meeting',
        topic: 'IT Eng Standup',
        lastMeaningfulMessageAt: new Date('2026-06-02T10:00:00.000Z'),
      },
      {
        _id: 'conv-it-b',
        graphChatId: 'meeting-it-b',
        chatType: 'meeting',
        topic: 'IT Eng Standup',
        lastMeaningfulMessageAt: new Date('2026-05-30T10:00:00.000Z'),
      },
    ]);

    const result = await TeamsArchiveService.conversationSenderMessages(user, {
      topic: 'IT Eng Standup',
      chatType: 'meeting',
      senderUserId: 'sender-1',
    });

    expect(result).toMatchObject({
      retrievalMode: 'conversation_sender_messages',
      resolved: false,
      trace: {
        selectedBy: 'ambiguousTopicLookup',
        identityWarning: expect.any(String),
        candidateCount: 2,
        candidateGraphChatIds: ['meeting-it-a', 'meeting-it-b'],
      },
      identityWarning: {
        reason: 'ambiguous_topic_lookup',
        candidateCount: 2,
      },
      messages: [],
    });
    expect(db.findTeamsArchiveMessages).not.toHaveBeenCalled();
  });

  it('resolves Mongo conversation ids for conversation_recent_messages', async () => {
    const conversation = {
      _id: '507f1f77bcf86cd799439011',
      graphChatId: 'meeting-from-mongo-id',
      chatType: 'meeting',
      topic: 'ObjectId resolution',
    };
    db.findTeamsArchiveConversations
      .mockResolvedValueOnce([])
      .mockResolvedValueOnce([conversation])
      .mockResolvedValueOnce([conversation]);
    db.findTeamsArchiveMessages.mockResolvedValue([
      {
        _id: 'mongo-msg',
        graphMessageId: 'g-mongo',
        graphChatId: 'meeting-from-mongo-id',
        bodyPreview: 'Resolved through Mongo id',
        sentDateTime: new Date('2026-06-01T10:00:00.000Z'),
        isChunkable: true,
        normalizedTextLength: 25,
        isSystemLikeMessage: false,
      },
    ]);
    db.countTeamsArchiveMessages.mockResolvedValue(0);

    const result = await TeamsArchiveService.conversationRecentMessages(user, {
      chatId: '507f1f77bcf86cd799439011',
      limit: 1,
    });

    expect(result).toMatchObject({
      retrievalMode: 'conversation_recent_messages',
      resolved: true,
      graphChatId: 'meeting-from-mongo-id',
      selectedConversation: {
        archiveConversationId: '507f1f77bcf86cd799439011',
        graphChatId: 'meeting-from-mongo-id',
      },
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
      selectedConversation: {
        graphChatId: 'chat-summary',
      },
      trust: {
        confidence: 'low',
        evidence: expect.any(Array),
        inferences: [],
        unknowns: expect.any(Array),
      },
    });
    expect(result.highlights).toHaveLength(1);
    expect(result.highlights[0]).toMatchObject({
      graphMessageId: 'sum-g-2',
    });
    expect(result.trust.evidence[0].sourceMessageIds).toEqual(['sum-g-2']);
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

  it('dry-runs conversation recency backfill without updating records', async () => {
    const conversation = {
      _id: 'conv-backfill',
      graphChatId: 'chat-backfill',
      topic: 'Backfill test',
      lastMessageAt: null,
      messageCount: 0,
    };
    db.findTeamsArchiveConversations.mockResolvedValue([conversation]);
    db.countTeamsArchiveMessages.mockResolvedValue(3);
    db.findTeamsArchiveMessages.mockResolvedValue([
      {
        graphChatId: 'chat-backfill',
        sentDateTime: new Date('2026-06-02T10:00:00.000Z'),
        isSystemLikeMessage: true,
        isChunkable: false,
        normalizedTextLength: 0,
        skipChunkReason: 'system_like_message',
      },
      {
        graphChatId: 'chat-backfill',
        sentDateTime: new Date('2026-06-02T09:00:00.000Z'),
        isSystemLikeMessage: false,
        isChunkable: true,
        normalizedTextLength: 24,
      },
      {
        graphChatId: 'chat-backfill',
        sentDateTime: new Date('2026-06-02T08:00:00.000Z'),
        isSystemLikeMessage: false,
        isChunkable: false,
        normalizedTextLength: 0,
        skipChunkReason: 'empty_normalized_text',
      },
    ]);

    const result = await TeamsArchiveService.backfillConversationRecency(user, {
      apply: false,
    });

    expect(db.updateTeamsArchiveConversation).not.toHaveBeenCalled();
    expect(result).toMatchObject({
      retrievalMode: 'conversation_recency_backfill',
      dryRun: true,
      processedConversationCount: 1,
      changedConversationCount: 1,
      updatedConversationCount: 0,
    });
    expect(result.conversations[0].newRecency).toMatchObject({
      meaningfulMessageCount: 1,
      humanMessageCount: 1,
      systemMessageCount: 1,
      emptyMessageCount: 2,
      messageCount: 3,
    });
  });

  it('applies conversation recency backfill for one graph chat id', async () => {
    const conversation = {
      _id: 'conv-backfill-apply',
      graphChatId: 'chat-backfill-apply',
      topic: 'Backfill apply test',
      lastMessageAt: null,
      messageCount: 0,
    };
    db.findTeamsArchiveConversations
      .mockResolvedValueOnce([conversation])
      .mockResolvedValueOnce([conversation]);
    db.countTeamsArchiveMessages.mockResolvedValue(1);
    db.findTeamsArchiveMessages.mockResolvedValue([
      {
        graphChatId: 'chat-backfill-apply',
        sentDateTime: new Date('2026-06-02T09:00:00.000Z'),
        isSystemLikeMessage: false,
        isChunkable: true,
        normalizedTextLength: 24,
      },
    ]);

    const result = await TeamsArchiveService.backfillConversationRecency(user, {
      chatId: 'chat-backfill-apply',
      apply: true,
    });

    expect(db.updateTeamsArchiveConversation).toHaveBeenCalledWith(
      'conv-backfill-apply',
      expect.objectContaining({
        lastMeaningfulMessageAt: new Date('2026-06-02T09:00:00.000Z'),
        meaningfulMessageCount: 1,
        messageCount: 1,
      }),
    );
    expect(result).toMatchObject({
      processedConversationCount: 1,
      updatedConversationCount: 1,
    });
  });

  it('returns sender-scoped messages from the signed-in user with match trace', async () => {
    const conversation = {
      _id: 'conv-sender',
      graphChatId: 'chat-sender',
      chatType: 'meeting',
      topic: 'Sender test',
      lastMeaningfulMessageAt: new Date('2026-06-02T10:00:00.000Z'),
    };
    db.findTeamsArchiveConversations
      .mockResolvedValueOnce([conversation])
      .mockResolvedValueOnce([conversation]);
    db.findTeamsArchiveMessages.mockResolvedValue([
      {
        _id: 'sender-msg',
        graphMessageId: 'g-sender',
        graphChatId: 'chat-sender',
        fromUserId: 'entra-user-1',
        fromEmail: 'user@example.com',
        fromDisplayName: 'Test User',
        bodyPreview: 'My latest update',
        sentDateTime: new Date('2026-06-02T10:00:00.000Z'),
        isChunkable: true,
        normalizedTextLength: 16,
        isSystemLikeMessage: false,
      },
    ]);
    db.countTeamsArchiveMessages.mockResolvedValue(1);

    const result = await TeamsArchiveService.conversationSenderMessages(user, {
      chatId: 'chat-sender',
      senderScope: 'me',
      limit: 4,
    });

    expect(db.findTeamsArchiveMessages).toHaveBeenCalledWith(
      expect.objectContaining({
        user: 'user-1',
        graphChatId: 'chat-sender',
        $or: expect.any(Array),
        $and: expect.any(Array),
      }),
      expect.objectContaining({
        sort: { sentDateTime: -1, createdAt: -1 },
      }),
    );
    expect(result).toMatchObject({
      retrievalMode: 'conversation_sender_messages',
      resolved: true,
      senderResolution: {
        senderScope: 'me',
        confidence: 'high',
      },
      retrievalTrace: {
        senderFilterApplied: true,
      },
      messages: [
        expect.objectContaining({
          graphMessageId: 'g-sender',
          senderMatch: {
            matchedBy: 'fromUserId',
            confidence: 'high',
          },
        }),
      ],
    });
  });

  it('treats explicit senderUserId as person scope even when senderScope is omitted', async () => {
    const conversation = {
      _id: 'conv-explicit-sender',
      graphChatId: 'chat-explicit-sender',
      chatType: 'meeting',
      topic: 'Explicit sender test',
    };
    db.findTeamsArchiveConversations
      .mockResolvedValueOnce([conversation])
      .mockResolvedValueOnce([conversation]);
    db.findTeamsArchiveMessages.mockResolvedValue([
      {
        _id: 'explicit-sender-msg',
        graphMessageId: 'g-explicit-sender',
        graphChatId: 'chat-explicit-sender',
        fromUserId: '0428da6a-d030-4547-bbbb-3ae6514fdf2b',
        fromEmail: 'aadUser',
        fromDisplayName: 'Praneet Kotah',
        bodyPreview: 'Legacy identity update',
        sentDateTime: new Date('2026-06-02T10:00:00.000Z'),
      },
    ]);
    db.countTeamsArchiveMessages.mockResolvedValue(1);

    const result = await TeamsArchiveService.conversationSenderMessages(user, {
      chatId: 'chat-explicit-sender',
      senderUserId: '0428da6a-d030-4547-bbbb-3ae6514fdf2b',
      limit: 4,
    });

    expect(db.findTeamsArchiveMessages).toHaveBeenCalledWith(
      expect.objectContaining({
        user: 'user-1',
        graphChatId: 'chat-explicit-sender',
        $or: [{ fromUserId: '0428da6a-d030-4547-bbbb-3ae6514fdf2b' }],
        $and: expect.any(Array),
      }),
      expect.anything(),
    );
    expect(result).toMatchObject({
      senderResolution: {
        senderScope: 'person',
        senderUserId: '0428da6a-d030-4547-bbbb-3ae6514fdf2b',
        confidence: 'high',
      },
      messages: [
        expect.objectContaining({
          graphMessageId: 'g-explicit-sender',
          senderMatch: {
            matchedBy: 'fromUserId',
            confidence: 'high',
          },
        }),
      ],
    });
  });

  it('includes legacy aadUser sender messages with bodyText/bodyPreview despite missing normalized fields', async () => {
    const conversation = {
      _id: 'conv-legacy-sender',
      graphChatId: 'chat-legacy-sender',
      chatType: 'meeting',
      topic: 'Legacy sender test',
    };
    db.findTeamsArchiveConversations
      .mockResolvedValueOnce([conversation])
      .mockResolvedValueOnce([conversation]);
    db.findTeamsArchiveMessages.mockResolvedValue([
      {
        _id: 'legacy-body-text',
        graphMessageId: 'g-legacy-body-text',
        graphChatId: 'chat-legacy-sender',
        fromEmail: 'aadUser',
        fromDisplayName: 'Test User',
        bodyText: 'Legacy bodyText should still be searchable',
        sentDateTime: new Date('2026-06-02T10:00:00.000Z'),
      },
      {
        _id: 'legacy-preview',
        graphMessageId: 'g-legacy-preview',
        graphChatId: 'chat-legacy-sender',
        fromEmail: 'aadUser',
        fromDisplayName: 'Test User',
        bodyPreview: 'Legacy bodyPreview should still be readable',
        sentDateTime: new Date('2026-06-01T10:00:00.000Z'),
      },
    ]);
    db.countTeamsArchiveMessages.mockResolvedValue(2);

    const result = await TeamsArchiveService.conversationSenderMessages(user, {
      chatId: 'chat-legacy-sender',
      senderName: 'Test User',
      limit: 4,
    });

    expect(db.findTeamsArchiveMessages).toHaveBeenCalledWith(
      expect.objectContaining({
        user: 'user-1',
        graphChatId: 'chat-legacy-sender',
        $or: [{ fromDisplayName: 'Test User' }],
        $and: expect.arrayContaining([
          expect.objectContaining({
            $or: expect.arrayContaining([
              expect.objectContaining({ bodyText: expect.any(RegExp) }),
              expect.objectContaining({ bodyPreview: expect.any(RegExp) }),
            ]),
          }),
        ]),
      }),
      expect.anything(),
    );
    expect(result).toMatchObject({
      senderResolution: {
        senderScope: 'person',
        senderName: 'Test User',
        confidence: 'medium',
      },
      messages: [
        expect.objectContaining({
          graphMessageId: 'g-legacy-body-text',
          senderMatch: {
            matchedBy: 'fromDisplayName',
            confidence: 'medium',
          },
        }),
        expect.objectContaining({
          graphMessageId: 'g-legacy-preview',
          senderMatch: {
            matchedBy: 'fromDisplayName',
            confidence: 'medium',
          },
        }),
      ],
    });
  });

  it('returns zero-result sender diagnostics with observed senders and same-title alternatives', async () => {
    const selectedConversation = {
      _id: 'conv-zero-selected',
      graphChatId: 'meeting-zero-selected',
      chatType: 'meeting',
      topic: 'IT Eng Standup',
      syncStatus: 'complete',
    };
    const alternativeConversation = {
      _id: 'conv-zero-alt',
      graphChatId: 'meeting-zero-alt',
      chatType: 'meeting',
      topic: 'IT Eng Standup',
      lastMeaningfulMessageAt: new Date('2026-05-30T16:25:17.385Z'),
    };
    db.findTeamsArchiveConversations
      .mockResolvedValueOnce([selectedConversation])
      .mockResolvedValueOnce([selectedConversation])
      .mockResolvedValueOnce([selectedConversation, alternativeConversation]);
    db.findTeamsArchiveMessages
      .mockResolvedValueOnce([])
      .mockResolvedValueOnce([
        {
          _id: 'observed-sender-msg',
          graphMessageId: 'g-observed-sender',
          graphChatId: 'meeting-zero-selected',
          fromUserId: 'other-sender',
          fromEmail: 'other@example.com',
          fromDisplayName: 'Other Sender',
          bodyPreview: 'Someone else posted here',
          sentDateTime: new Date('2026-06-02T10:00:00.000Z'),
          isChunkable: true,
          normalizedTextLength: 24,
          isSystemLikeMessage: false,
        },
      ]);
    db.countTeamsArchiveMessages
      .mockResolvedValueOnce(0)
      .mockResolvedValueOnce(0)
      .mockResolvedValueOnce(0)
      .mockResolvedValueOnce(1)
      .mockResolvedValueOnce(1);

    const result = await TeamsArchiveService.conversationSenderMessages(user, {
      chatId: 'meeting-zero-selected',
      senderUserId: 'missing-sender',
      limit: 4,
    });

    expect(result).toMatchObject({
      retrievalMode: 'conversation_sender_messages',
      resolved: true,
      messages: [],
      zeroResultDiagnostics: {
        selectedConversation: {
          graphChatId: 'meeting-zero-selected',
        },
        totalMessagesInConversation: 1,
        humanReadableMessagesInConversation: 1,
        sameTitleAlternativeGraphChatIds: ['meeting-zero-alt'],
        likelyReasons: expect.arrayContaining([
          'selected wrong recurring meeting instance',
          'messages exist only in a different Teams chat/thread',
        ]),
        recommendedNextActions: expect.arrayContaining([
          'Run sender_identity_report for this graphChatId and compare observed sender identities.',
        ]),
      },
    });
    expect(result.zeroResultDiagnostics.uniqueObservedSenders).toEqual([
      expect.objectContaining({
        fromUserId: 'other-sender',
        fromEmail: 'other@example.com',
        fromDisplayName: 'Other Sender',
        count: 1,
        sampleMessageIds: ['g-observed-sender'],
      }),
    ]);
  });

  it('falls back to display-name alias candidates when sender identity report finds no direct me matches', async () => {
    db.findTeamsArchiveMessages
      .mockResolvedValueOnce([])
      .mockResolvedValueOnce([
        {
          _id: 'alias-a',
          graphMessageId: 'g-alias-a',
          graphChatId: 'chat-alias',
          fromUserId: '0428da6a-d030-4547-bbbb-3ae6514fdf2b',
          fromEmail: 'aadUser',
          fromDisplayName: 'Test User',
          bodyPreview: 'Alias sample',
          sentDateTime: new Date('2026-06-02T10:00:00.000Z'),
        },
      ]);
    db.findTeamsArchiveConversations.mockResolvedValue([
      {
        _id: 'conv-alias',
        graphChatId: 'chat-alias',
        topic: 'Alias meeting',
      },
    ]);

    const result = await TeamsArchiveService.senderIdentityReport(user, {
      chatId: 'chat-alias',
      senderScope: 'me',
    });

    expect(result).toMatchObject({
      retrievalMode: 'sender_identity_report',
      senderScope: 'me',
      warnings: expect.arrayContaining([
        'direct_me_match_returned_zero_used_alias_fallback',
        'invalid_fromEmail_values_detected',
      ]),
      identityCandidates: [
        expect.objectContaining({
          fromUserId: '0428da6a-d030-4547-bbbb-3ae6514fdf2b',
          fromEmail: 'aadUser',
          fromDisplayName: 'Test User',
          senderMatch: {
            matchedBy: 'alias',
            confidence: 'medium',
          },
        }),
      ],
    });
  });

  it('explains system-only recent activity in conversation diagnostics', async () => {
    const conversation = {
      _id: 'conv-diagnostics',
      graphChatId: 'chat-diagnostics',
      chatType: 'meeting',
      topic: 'Diagnostics meeting',
      lastMessageAt: new Date('2026-06-03T10:00:00.000Z'),
      lastMeaningfulMessageAt: null,
      lastSystemMessageAt: new Date('2026-06-03T10:00:00.000Z'),
      messageCount: 8,
      meaningfulMessageCount: 0,
      systemMessageCount: 8,
      emptyMessageCount: 0,
      participantDegraded: true,
      syncStatus: 'complete',
    };
    db.findTeamsArchiveConversations
      .mockResolvedValueOnce([conversation])
      .mockResolvedValueOnce([conversation]);

    const result = await TeamsArchiveService.conversationActivityDiagnostics(user, {
      chatId: 'chat-diagnostics',
      includeRecentMessages: true,
      includeSystem: true,
    });

    expect(result).toMatchObject({
      retrievalMode: 'conversation_activity_diagnostics',
      resolved: true,
      diagnosis: {
        recentBecauseOfSystemActivity: true,
        noMeaningfulMessages: true,
        participantMetadataDegraded: true,
        zeroChunkRisk: true,
      },
    });
    expect(result.explanation).toContain('system-driven');
  });

  it('reports sender identity candidates and invalid aadUser email values', async () => {
    db.findTeamsArchiveMessages.mockResolvedValue([
      {
        _id: 'identity-a',
        graphMessageId: 'g-identity-a',
        graphChatId: 'chat-identity',
        fromUserId: 'aad-user-1',
        fromEmail: 'aadUser',
        fromDisplayName: 'Test User',
        bodyPreview: 'Identity sample',
        sentDateTime: new Date('2026-06-02T10:00:00.000Z'),
      },
      {
        _id: 'identity-b',
        graphMessageId: 'g-identity-b',
        graphChatId: 'chat-identity',
        fromUserId: 'entra-user-1',
        fromEmail: 'user@example.com',
        fromDisplayName: 'Test User',
        bodyPreview: 'Identity sample 2',
        sentDateTime: new Date('2026-06-01T10:00:00.000Z'),
      },
    ]);

    const result = await TeamsArchiveService.senderIdentityReport(user, {
      senderScope: 'me',
    });

    expect(result).toMatchObject({
      retrievalMode: 'sender_identity_report',
      senderScope: 'me',
      confidence: 'high',
      warnings: expect.arrayContaining(['invalid_fromEmail_values_detected']),
    });
    expect(result.identityCandidates).toHaveLength(2);
    expect(result.recommendedSenderFilters).toMatchObject({
      fromUserId: expect.any(String),
    });
  });

  it('marks broad completeness-sensitive preview retrieval as insufficient evidence', async () => {
    db.findTeamsArchiveConversations.mockResolvedValue([]);
    db.findTeamsArchiveMessages.mockResolvedValue([
      {
        _id: 'broad-a',
        graphMessageId: 'g-broad-a',
        graphChatId: 'chat-a',
        bodyPreview: 'Action item preview',
        bodyText: 'Action item preview',
        sentDateTime: new Date('2026-06-01T10:00:00.000Z'),
        isChunkable: true,
        normalizedTextLength: 19,
        isSystemLikeMessage: false,
      },
      {
        _id: 'broad-b',
        graphMessageId: 'g-broad-b',
        graphChatId: 'chat-b',
        bodyPreview: 'Decision preview',
        bodyText: 'Decision preview',
        sentDateTime: new Date('2026-06-01T11:00:00.000Z'),
        isChunkable: true,
        normalizedTextLength: 16,
        isSystemLikeMessage: false,
      },
    ]);

    const result = await TeamsArchiveService.searchMessages(user, {
      query: 'all action items and decisions',
      limit: 2,
    });

    expect(result.evidenceBudget).toMatchObject({
      requestedCompleteness: true,
      conversationsScoped: 2,
      evidenceSufficient: false,
      insufficiencyReasons: expect.arrayContaining(['conversation_not_uniquely_scoped']),
    });
  });
});
