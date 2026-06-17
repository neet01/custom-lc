jest.mock('~/models', () => ({
  findEnterpriseMemoryChunks: jest.fn(),
  findTeamsArchiveConversations: jest.fn(),
  findSlackArchiveConversations: jest.fn(),
  findSlackIdentityLink: jest.fn(),
}));

const db = require('~/models');
const { searchTeamsMemoryChunks, searchSlackMemoryChunks } = require('./retrieval');

describe('EnterpriseMemory retrieval', () => {
  const user = {
    id: 'user-1',
    openidId: 'entra-user-1',
    email: 'user@example.com',
    name: 'Test User',
    username: 'test.user',
  };

  beforeEach(() => {
    jest.clearAllMocks();
    delete process.env.TEAMS_ARCHIVE_TEXT_SEARCH_ENABLED;
    delete process.env.SLACK_ARCHIVE_TEXT_SEARCH_ENABLED;
    db.findEnterpriseMemoryChunks.mockResolvedValue([]);
    db.findTeamsArchiveConversations.mockResolvedValue([]);
    db.findSlackArchiveConversations.mockResolvedValue([]);
    db.findSlackIdentityLink.mockResolvedValue(null);
  });

  afterEach(() => {
    delete process.env.TEAMS_ARCHIVE_TEXT_SEARCH_ENABLED;
    delete process.env.SLACK_ARCHIVE_TEXT_SEARCH_ENABLED;
  });

  it('does not prefilter conversations by topic before chunk search for broad topic queries', async () => {
    db.findEnterpriseMemoryChunks.mockResolvedValue([
      {
        _id: 'chunk-1',
        sourceRecordId: 'graph-msg-1',
        sourceParentRecordId: 'chat-1',
        summary: 'Condor thermal issue',
        text: 'Condor thermal issue and heat shield rework details',
        sourceTimestamp: new Date('2026-04-20T12:00:00.000Z'),
        metadata: {
          fromDisplayName: 'Lead',
          fromEmail: 'lead@example.com',
        },
      },
    ]);

    const result = await searchTeamsMemoryChunks(user, {
      topic: 'condor thermal',
      limit: 4,
    });

    expect(db.findTeamsArchiveConversations).toHaveBeenCalledWith(
      {
        user: 'user-1',
        graphChatId: { $in: ['chat-1'] },
      },
      { limit: 1 },
    );
    expect(db.findEnterpriseMemoryChunks).toHaveBeenCalledWith(
      expect.objectContaining({
        user: 'user-1',
        source: 'teams',
        sourceRecordType: 'teams_message',
        $and: expect.any(Array),
      }),
      expect.objectContaining({
        limit: 4,
      }),
    );
    expect(result).toMatchObject({
      retrievalMode: 'enterprise_memory',
      resultCount: 1,
      trace: {
        conversationPrefilterApplied: false,
        topicPrefilterApplied: false,
        matchedConversationCount: 0,
      },
    });
  });

  it('scopes memory chunk search by Teams graph chat id when chatId is provided', async () => {
    const graphChatId = '19:meeting_YTY0ZWU3NTEtNWJjNi00NmNkLTgxODAtZjdjMmQxZTEzZDQz@thread.v2';
    db.findEnterpriseMemoryChunks.mockResolvedValue([
      {
        _id: 'chunk-chat-scoped',
        sourceRecordId: 'graph-msg-chat-scoped',
        sourceParentRecordId: graphChatId,
        summary: 'Telemetry replay update',
        text: 'Telemetry replay update from the meeting chat',
        sourceTimestamp: new Date('2026-05-20T12:00:00.000Z'),
        metadata: {
          fromDisplayName: 'Test User',
          fromEmail: 'user@example.com',
        },
      },
    ]);

    const result = await searchTeamsMemoryChunks(user, {
      chatId: graphChatId,
      topic: 'telemetry replay',
      chatType: 'meeting',
      limit: 4,
    });

    expect(db.findTeamsArchiveConversations).toHaveBeenCalledWith(
      {
        user: 'user-1',
        graphChatId: { $in: [graphChatId] },
      },
      { limit: 1 },
    );
    expect(db.findEnterpriseMemoryChunks).toHaveBeenCalledWith(
      expect.objectContaining({
        user: 'user-1',
        source: 'teams',
        sourceRecordType: 'teams_message',
        sourceParentRecordId: { $in: [graphChatId] },
        $and: expect.any(Array),
      }),
      expect.objectContaining({
        limit: 4,
      }),
    );
    expect(result).toMatchObject({
      chatId: graphChatId,
      graphChatId,
      trace: expect.objectContaining({
        conversationPrefilterApplied: false,
        chatIdScoped: true,
        matchedConversationCount: 1,
      }),
    });
  });

  it('uses the $text index with relevance ranking when text search is enabled', async () => {
    process.env.TEAMS_ARCHIVE_TEXT_SEARCH_ENABLED = 'true';
    db.findEnterpriseMemoryChunks.mockResolvedValue([
      {
        _id: 'chunk-text',
        sourceRecordId: 'graph-msg-text',
        sourceParentRecordId: 'chat-text',
        summary: 'Condor thermal issue',
        text: 'Condor thermal issue and heat shield rework details',
        sourceTimestamp: new Date('2026-04-20T12:00:00.000Z'),
        metadata: { fromDisplayName: 'Lead', fromEmail: 'lead@example.com' },
      },
    ]);

    const result = await searchTeamsMemoryChunks(user, { topic: 'condor thermal', limit: 4 });

    expect(db.findEnterpriseMemoryChunks).toHaveBeenCalledTimes(1);
    const [filter, queryOptions] = db.findEnterpriseMemoryChunks.mock.calls[0];
    expect(filter).toMatchObject({
      user: 'user-1',
      source: 'teams',
      sourceRecordType: 'teams_message',
      $text: { $search: 'condor thermal' },
    });
    expect(filter.$and).toBeUndefined();
    expect(queryOptions).toMatchObject({ limit: 4, textScore: true });
    expect(result.trace).toMatchObject({ searchBackend: 'text', textSearchFellBack: false });
    expect(result.resultCount).toBe(1);
  });

  it('falls back to regex search when the $text query returns no results', async () => {
    process.env.TEAMS_ARCHIVE_TEXT_SEARCH_ENABLED = 'true';
    db.findEnterpriseMemoryChunks
      .mockResolvedValueOnce([])
      .mockResolvedValueOnce([
        {
          _id: 'chunk-regex',
          sourceRecordId: 'graph-msg-regex',
          sourceParentRecordId: 'chat-regex',
          summary: 'Condor thermal issue',
          text: 'Condor thermal issue and heat shield rework details',
          sourceTimestamp: new Date('2026-04-20T12:00:00.000Z'),
          metadata: { fromDisplayName: 'Lead', fromEmail: 'lead@example.com' },
        },
      ]);

    const result = await searchTeamsMemoryChunks(user, { topic: 'condor thermal', limit: 4 });

    expect(db.findEnterpriseMemoryChunks).toHaveBeenCalledTimes(2);
    expect(db.findEnterpriseMemoryChunks.mock.calls[0][0]).toMatchObject({
      $text: { $search: 'condor thermal' },
    });
    const fallbackFilter = db.findEnterpriseMemoryChunks.mock.calls[1][0];
    expect(fallbackFilter.$text).toBeUndefined();
    expect(fallbackFilter.$and).toEqual(expect.any(Array));
    expect(result.trace).toMatchObject({ searchBackend: 'regex', textSearchFellBack: true });
    expect(result.resultCount).toBe(1);
  });

  it('uses the Slack enterprise memory text index by default', async () => {
    db.findEnterpriseMemoryChunks.mockResolvedValue([
      {
        _id: 'slack-chunk-text',
        sourceRecordId: '1714521600.000100',
        sourceParentRecordId: 'COPS',
        summary: 'Budget approval update',
        text: 'Budget approval update from the ops channel',
        sourceTimestamp: new Date('2026-05-20T12:00:00.000Z'),
        metadata: {
          displayName: 'Test User',
          slackUserId: 'U123',
          conversationType: 'public_channel',
        },
      },
    ]);
    db.findSlackArchiveConversations.mockResolvedValue([
      {
        slackConversationId: 'COPS',
        name: 'ops',
        topic: 'Operations',
        conversationType: 'public_channel',
        participants: [{ slackUserId: 'U123', displayName: 'Test User' }],
      },
    ]);

    const result = await searchSlackMemoryChunks(user, {
      topic: 'budget approval',
      limit: 4,
    });

    expect(db.findEnterpriseMemoryChunks).toHaveBeenCalledWith(
      expect.objectContaining({
        user: 'user-1',
        source: 'slack',
        sourceRecordType: { $in: ['slack_message', 'slack_conversation'] },
        $text: { $search: 'budget approval' },
      }),
      expect.objectContaining({
        limit: 4,
        textScore: true,
      }),
    );
    expect(db.findSlackArchiveConversations).toHaveBeenCalledWith(
      {
        user: 'user-1',
        slackConversationId: { $in: ['COPS'] },
      },
      { limit: 1 },
    );
    expect(result).toMatchObject({
      retrievalMode: 'enterprise_memory',
      source: 'slack',
      resultCount: 1,
      trace: {
        searchBackend: 'text',
        textSearchEnabled: true,
        textSearchFellBack: false,
        searchedSourceRecordTypes: ['slack_message', 'slack_conversation'],
      },
      results: [
        expect.objectContaining({
          slackConversationId: 'COPS',
          conversationName: 'ops',
          fromDisplayName: 'Test User',
        }),
      ],
    });
  });

  it('prefilters Slack conversations before indexed search when participants are provided', async () => {
    db.findSlackArchiveConversations
      .mockResolvedValueOnce([
        {
          slackConversationId: 'DTEAM',
          conversationType: 'im',
          participants: [{ displayName: 'Taylor', slackUserId: 'U456' }],
        },
      ])
      .mockResolvedValueOnce([
        {
          slackConversationId: 'DTEAM',
          conversationType: 'im',
          participants: [{ displayName: 'Taylor', slackUserId: 'U456' }],
        },
      ]);
    db.findEnterpriseMemoryChunks.mockResolvedValue([
      {
        _id: 'slack-chunk-scoped',
        sourceRecordId: '1714521600.000200',
        sourceParentRecordId: 'DTEAM',
        summary: 'Launch checklist',
        text: 'Launch checklist discussion',
        sourceTimestamp: new Date('2026-05-21T12:00:00.000Z'),
        metadata: {
          displayName: 'Taylor',
          slackUserId: 'U456',
          conversationType: 'im',
        },
      },
    ]);

    const result = await searchSlackMemoryChunks(user, {
      topic: 'launch checklist',
      conversationType: 'im',
      participants: ['Taylor'],
      limit: 3,
    });

    expect(db.findSlackArchiveConversations).toHaveBeenNthCalledWith(
      1,
      expect.objectContaining({
        user: 'user-1',
        conversationType: 'im',
        $and: expect.any(Array),
      }),
      { limit: 1000 },
    );
    expect(db.findEnterpriseMemoryChunks).toHaveBeenCalledWith(
      expect.objectContaining({
        source: 'slack',
        sourceRecordType: { $in: ['slack_message', 'slack_conversation'] },
        sourceParentRecordId: { $in: ['DTEAM'] },
        $text: { $search: 'launch checklist' },
      }),
      expect.objectContaining({
        limit: 3,
      }),
    );
    expect(result.trace).toMatchObject({
      searchBackend: 'text',
      conversationPrefilterApplied: true,
      matchedConversationCount: 1,
    });
  });

  it('keeps sender-scoped Slack indexed search on per-message chunks only', async () => {
    db.findEnterpriseMemoryChunks.mockResolvedValue([
      {
        _id: 'slack-sender-chunk',
        sourceRecordType: 'slack_message',
        sourceRecordId: '1714521600.000300',
        sourceParentRecordId: 'COPS',
        summary: 'I will follow up',
        text: 'I will follow up',
        sourceTimestamp: new Date('2026-05-22T12:00:00.000Z'),
        metadata: {
          displayName: 'Test User',
          slackUserId: 'U123',
        },
      },
    ]);

    await searchSlackMemoryChunks(user, {
      topic: 'follow up',
      senderUserId: 'U123',
      limit: 3,
    });

    expect(db.findEnterpriseMemoryChunks).toHaveBeenCalledWith(
      expect.objectContaining({
        source: 'slack',
        sourceRecordType: 'slack_message',
        'metadata.slackUserId': 'U123',
        $text: { $search: 'follow up' },
      }),
      expect.objectContaining({
        limit: 3,
      }),
    );
  });

  it('resolves a responsibility query with intent re-ranking and message-level attribution', async () => {
    db.findEnterpriseMemoryChunks.mockResolvedValue([
      {
        _id: 'topic-only',
        sourceRecordType: 'slack_message',
        sourceRecordId: '1.1',
        sourceParentRecordId: 'CENG',
        text: 'the landing gear inspection is overdue',
        summary: 'the landing gear inspection is overdue',
        sourceTimestamp: new Date('2026-06-01T10:00:00Z'),
        metadata: { displayName: 'Sam', slackUserId: 'U010' },
      },
      {
        _id: 'intent-and-topic',
        sourceRecordType: 'slack_message',
        sourceRecordId: '1.2',
        sourceParentRecordId: 'CENG',
        text: "I'll own landing gear going forward",
        summary: "I'll own landing gear going forward",
        sourceTimestamp: new Date('2026-06-02T10:00:00Z'),
        metadata: { displayName: 'Dana', slackUserId: 'U777' },
      },
      {
        _id: 'intent-only',
        sourceRecordType: 'slack_message',
        sourceRecordId: '1.3',
        sourceParentRecordId: 'CENG',
        text: "I'll own the deployment pipeline",
        summary: "I'll own the deployment pipeline",
        sourceTimestamp: new Date('2026-06-03T10:00:00Z'),
        metadata: { displayName: 'Lee', slackUserId: 'U999' },
      },
    ]);

    const result = await searchSlackMemoryChunks(user, {
      topic: 'who said they are responsible for landing gear',
    });

    expect(db.findSlackIdentityLink).not.toHaveBeenCalled();
    const [candidateFilter, candidateOptions] = db.findEnterpriseMemoryChunks.mock.calls[0];
    expect(candidateFilter).toMatchObject({
      source: 'slack',
      sourceRecordType: 'slack_message',
      $or: expect.any(Array),
    });
    expect(candidateOptions.limit).toBe(32);
    expect(result.trace).toMatchObject({
      searchBackend: 'hybrid_intent',
      queryArchetype: 'responsibility',
      intentDriven: true,
      searchedSourceRecordTypes: ['slack_message'],
    });
    expect(result.results[0]).toMatchObject({
      id: 'intent-and-topic',
      fromDisplayName: 'Dana',
    });
    expect(result.results[0].matchedIntent).toBeTruthy();
    expect(result.results[0].relevanceScore).toBeGreaterThan(result.results[1].relevanceScore);
  });

  it('resolves "me" from the Slack identity link and infers today for a self-commitment query', async () => {
    db.findSlackIdentityLink.mockResolvedValue({ slackUserId: 'U123', status: 'linked' });
    db.findEnterpriseMemoryChunks.mockResolvedValue([
      {
        _id: 'greeting',
        sourceRecordType: 'slack_message',
        sourceRecordId: '2.1',
        sourceParentRecordId: 'CGEN',
        text: 'good morning everyone',
        summary: 'good morning everyone',
        sourceTimestamp: new Date('2026-06-17T09:00:00Z'),
        metadata: { displayName: 'Test User', slackUserId: 'U123' },
      },
      {
        _id: 'commitment',
        sourceRecordType: 'slack_message',
        sourceRecordId: '2.2',
        sourceParentRecordId: 'CGEN',
        text: "I'll send the status report by EOD",
        summary: "I'll send the status report by EOD",
        sourceTimestamp: new Date('2026-06-17T08:00:00Z'),
        metadata: { displayName: 'Test User', slackUserId: 'U123' },
      },
    ]);

    const result = await searchSlackMemoryChunks(user, {
      topic: 'what did I tell people I would do today',
    });

    expect(db.findSlackIdentityLink).toHaveBeenCalledWith({ user: 'user-1', status: 'linked' });
    const [candidateFilter] = db.findEnterpriseMemoryChunks.mock.calls[0];
    expect(candidateFilter).toMatchObject({
      source: 'slack',
      sourceRecordType: 'slack_message',
      sourceTimestamp: expect.objectContaining({ $gte: expect.any(Date) }),
    });
    expect(candidateFilter.$or).toEqual(
      expect.arrayContaining([{ 'metadata.slackUserId': 'U123' }]),
    );
    expect(result.trace).toMatchObject({
      queryArchetype: 'commitment',
      appliedSenderScope: 'me',
      appliedDaysBack: 1,
      senderResolvedFromIdentityLink: true,
    });
    expect(result.results[0]).toMatchObject({ id: 'commitment' });
    expect(result.results[0].matchedIntent).toBeTruthy();
  });
});
