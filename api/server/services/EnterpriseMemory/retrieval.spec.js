jest.mock('~/models', () => ({
  findEnterpriseMemoryChunks: jest.fn(),
  findTeamsArchiveConversations: jest.fn(),
}));

const db = require('~/models');
const { searchTeamsMemoryChunks } = require('./retrieval');

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
    db.findEnterpriseMemoryChunks.mockResolvedValue([]);
    db.findTeamsArchiveConversations.mockResolvedValue([]);
  });

  afterEach(() => {
    delete process.env.TEAMS_ARCHIVE_TEXT_SEARCH_ENABLED;
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
});
