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
      error: jest.fn(),
    },
    runAsSystem: async (fn) => fn(),
  }),
  { virtual: true },
);

jest.mock('~/server/services/ArchiveFeatureAccess', () => ({
  isArchiveFeatureAllowed: jest.fn(),
}));

jest.mock('~/server/services/PluginService', () => ({
  getUserPluginAuthValue: jest.fn(),
}));

jest.mock('~/server/services/EnterpriseMemory/slackProjection', () => ({
  projectSlackArchiveSyncToMemory: jest.fn(),
}));

jest.mock('~/server/services/EnterpriseMemory/retrieval', () => ({
  searchSlackMemoryChunks: jest.fn(),
}));

jest.mock('~/models', () => ({
  findSlackArchiveConversations: jest.fn(),
  findSlackArchiveMessages: jest.fn(),
}));

const db = require('~/models');
const { isArchiveFeatureAllowed } = require('~/server/services/ArchiveFeatureAccess');
const { searchSlackMemoryChunks } = require('~/server/services/EnterpriseMemory/retrieval');
const SlackArchiveService = require('./SlackArchiveService');

describe('SlackArchiveService', () => {
  const originalEnv = process.env;
  const user = {
    id: 'user-1',
    name: 'Test User',
    username: 'test.user',
  };

  beforeEach(() => {
    jest.clearAllMocks();
    process.env = {
      ...originalEnv,
      SLACK_ARCHIVE_ENABLED: 'true',
    };
    isArchiveFeatureAllowed.mockResolvedValue(true);
    db.findSlackArchiveConversations.mockResolvedValue([]);
    db.findSlackArchiveMessages.mockResolvedValue([]);
  });

  afterEach(() => {
    process.env = originalEnv;
  });

  it('uses indexed Slack enterprise-memory results for advanced search', async () => {
    searchSlackMemoryChunks.mockResolvedValue({
      retrievalMode: 'enterprise_memory',
      source: 'slack',
      trace: {
        backend: 'enterprise_memory',
        searchBackend: 'text',
      },
      resultCount: 1,
      results: [
        {
          slackConversationId: 'COPS',
          slackMessageTs: '1714521600.000100',
          excerpt: 'Budget approval update',
        },
      ],
    });

    const result = await SlackArchiveService.advancedSearchMessages(user, {
      topic: 'budget approval',
      limit: 4,
    });

    expect(searchSlackMemoryChunks).toHaveBeenCalledWith(
      user,
      expect.objectContaining({
        topic: 'budget approval',
        limit: 4,
      }),
    );
    expect(db.findSlackArchiveMessages).not.toHaveBeenCalled();
    expect(result).toMatchObject({
      retrievalMode: 'advanced_indexed_slack_memory',
      resultCount: 1,
      trace: {
        searchBackend: 'text',
        archiveFallbackRan: false,
      },
    });
  });

  it('does not fall back to raw archive search for advanced search by default', async () => {
    searchSlackMemoryChunks.mockResolvedValue({
      retrievalMode: 'enterprise_memory',
      source: 'slack',
      trace: {
        backend: 'enterprise_memory',
        searchBackend: 'text',
      },
      resultCount: 0,
      results: [],
    });

    const result = await SlackArchiveService.advancedSearchMessages(user, {
      topic: 'budget approval',
      limit: 4,
    });

    expect(searchSlackMemoryChunks).toHaveBeenCalled();
    expect(db.findSlackArchiveMessages).not.toHaveBeenCalled();
    expect(result).toMatchObject({
      retrievalMode: 'advanced_indexed_slack_memory',
      resultCount: 0,
      trace: {
        backend: 'enterprise_memory',
        archiveFallbackRan: false,
        archiveFallbackDisabled: true,
      },
    });
  });

  it('preserves conversation fields when merging a hydrated Mongoose doc (regression: empty projection)', () => {
    const mongoose = require('mongoose');
    const schema = new mongoose.Schema({
      slackConversationId: String,
      name: String,
      syncAttemptCount: Number,
    });
    const Model =
      mongoose.models.SlackArchiveConvRegression ||
      mongoose.model('SlackArchiveConvRegression', schema);
    const hydratedDoc = new Model({ slackConversationId: 'C001', name: 'ops', syncAttemptCount: 0 });

    expect({ ...hydratedDoc }.slackConversationId).toBeUndefined();

    const channel = { id: 'C001', name: 'ops' };
    const merged = SlackArchiveService.toDiscoveredConversation(channel, hydratedDoc);

    expect(merged.slackConversationId).toBe('C001');
    expect(merged._id).toBeDefined();
    expect(merged.channel).toBe(channel);
  });

  it('resolves an internal conversation id before listing exact Slack messages', async () => {
    db.findSlackArchiveConversations.mockResolvedValue([
      {
        _id: '665f1d7f4e0a7a0012a34567',
        slackConversationId: 'COPS',
        name: 'ops',
      },
    ]);
    db.findSlackArchiveMessages.mockResolvedValue([
      {
        _id: 'msg-1',
        slackConversationId: 'COPS',
        slackMessageTs: '1714521600.000100',
        displayName: 'Manager',
        text: 'Budget approval update',
        normalizedText: 'Budget approval update',
        sentAt: new Date('2026-05-01T00:00:00.000Z'),
      },
    ]);

    const result = await SlackArchiveService.listConversationMessages(
      user,
      '665f1d7f4e0a7a0012a34567',
      { limit: 10 },
    );

    expect(db.findSlackArchiveConversations).toHaveBeenCalledWith(
      {
        user: 'user-1',
        $or: [
          { slackConversationId: '665f1d7f4e0a7a0012a34567' },
          { _id: '665f1d7f4e0a7a0012a34567' },
          { name: /^665f1d7f4e0a7a0012a34567$/i },
          { topic: /^665f1d7f4e0a7a0012a34567$/i },
        ],
      },
      { limit: 1 },
    );
    expect(db.findSlackArchiveMessages).toHaveBeenCalledWith(
      { user: 'user-1', slackConversationId: 'COPS' },
      expect.objectContaining({ limit: 10 }),
    );
    expect(result).toMatchObject({
      conversationId: 'COPS',
      requestedConversationId: '665f1d7f4e0a7a0012a34567',
      count: 1,
    });
  });
});
