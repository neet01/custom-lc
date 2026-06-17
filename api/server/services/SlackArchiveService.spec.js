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
});
