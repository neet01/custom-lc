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
    db.findEnterpriseMemoryChunks.mockResolvedValue([]);
    db.findTeamsArchiveConversations.mockResolvedValue([]);
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
});
