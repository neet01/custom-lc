const mockFind = jest.fn();
const mockFindOne = jest.fn();
const mockConversationCountDocuments = jest.fn();
const mockMessageCountDocuments = jest.fn();
const mockChunkCountDocuments = jest.fn();
const mockConversationAggregate = jest.fn();
const mockMessageAggregate = jest.fn();
const mockChunkAggregate = jest.fn();

function mockObjectId(value) {
  this.value = value;
}

mockObjectId.prototype.toString = function toString() {
  return this.value;
};

mockObjectId.isValid = jest.fn((value) => /^[a-f\d]{24}$/i.test(String(value || '')));

jest.mock('mongoose', () => ({
  Types: { ObjectId: mockObjectId },
  models: {
    SlackArchiveConversation: {
      find: mockFind,
      countDocuments: mockConversationCountDocuments,
      aggregate: mockConversationAggregate,
    },
    SlackArchiveMessage: {
      countDocuments: mockMessageCountDocuments,
      aggregate: mockMessageAggregate,
    },
    SlackArchiveSyncJob: {
      findOne: mockFindOne,
    },
    EnterpriseMemoryChunk: {
      countDocuments: mockChunkCountDocuments,
      aggregate: mockChunkAggregate,
    },
    EnterpriseMemoryJob: {
      findOne: mockFindOne,
    },
  },
}));

const { getArchiveDiagnostics } = require('./ArchiveDiagnosticsService');

function createLeanFindResult(rows) {
  return {
    sort: jest.fn().mockReturnThis(),
    skip: jest.fn().mockReturnThis(),
    limit: jest.fn().mockReturnThis(),
    lean: jest.fn().mockResolvedValue(rows),
  };
}

function createLeanFindOneResult(job) {
  return {
    sort: jest.fn().mockReturnThis(),
    lean: jest.fn().mockResolvedValue(job),
  };
}

function getGroupId(pipeline) {
  return pipeline.find((stage) => stage.$group)?.$group?._id;
}

function resetAggregateDefaults({
  messageStats = [],
  chunkStats = [],
  conversationsByType = [],
  conversationsByStatus = [],
  chunksByRecordType = [],
  chunksByChunkType = [],
  skippedReasons = [],
} = {}) {
  mockConversationAggregate.mockImplementation(async (pipeline) => {
    const groupId = getGroupId(pipeline);
    if (groupId === '$conversationType') {
      return conversationsByType;
    }
    if (groupId === '$syncStatus') {
      return conversationsByStatus;
    }
    return [];
  });

  mockMessageAggregate.mockImplementation(async (pipeline) => {
    const matchStage = pipeline.find((stage) => stage.$match)?.$match || {};
    if (matchStage.isChunkable === false) {
      return skippedReasons;
    }
    return messageStats;
  });

  mockChunkAggregate.mockImplementation(async (pipeline) => {
    const groupId = getGroupId(pipeline);
    if (groupId === '$sourceParentRecordId') {
      return chunkStats;
    }
    if (groupId === '$sourceRecordType') {
      return chunksByRecordType;
    }
    if (groupId === '$chunkType') {
      return chunksByChunkType;
    }
    return [];
  });
}

describe('ArchiveDiagnosticsService', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    mockFindOne.mockReturnValue(createLeanFindOneResult(null));
    mockConversationCountDocuments.mockResolvedValue(1);
    mockMessageCountDocuments.mockResolvedValue(2);
    mockChunkCountDocuments.mockResolvedValue(3);
    resetAggregateDefaults();
  });

  it('marks a Slack conversation healthy when archived messages have projected chunks', async () => {
    mockFind.mockReturnValue(
      createLeanFindResult([
        {
          _id: 'conv-1',
          user: 'user-1',
          slackConversationId: 'COPS',
          name: 'ops',
          conversationType: 'private_channel',
          syncStatus: 'complete',
          messageCount: 2,
          meaningfulMessageCount: 2,
          lastMessageAt: new Date('2026-06-17T12:00:00.000Z'),
          lastMeaningfulMessageAt: new Date('2026-06-17T12:00:00.000Z'),
          updatedAt: new Date('2026-06-17T12:05:00.000Z'),
        },
      ]),
    );
    resetAggregateDefaults({
      messageStats: [
        {
          _id: 'COPS',
          actualMessageCount: 2,
          chunkableMessageCount: 2,
          skippedMessageCount: 0,
        },
      ],
      chunkStats: [
        {
          _id: 'COPS',
          chunkCount: 3,
          messageChunkCount: 2,
          windowChunkCount: 1,
          latestChunkAt: new Date('2026-06-17T12:01:00.000Z'),
        },
      ],
      conversationsByType: [{ _id: 'private_channel', count: 1 }],
      conversationsByStatus: [{ _id: 'complete', count: 1 }],
      chunksByRecordType: [{ _id: 'slack_message', count: 3 }],
      chunksByChunkType: [{ _id: 'message', count: 2 }],
    });

    const result = await getArchiveDiagnostics({ source: 'slack' });

    expect(result.summary.totalConversations).toBe(1);
    expect(result.summary.totalMessages).toBe(2);
    expect(result.summary.totalChunks).toBe(3);
    expect(result.conversations).toHaveLength(1);
    expect(result.conversations[0]).toMatchObject({
      sourceConversationId: 'COPS',
      displayName: 'ops',
      health: {
        state: 'healthy',
        severity: 'ok',
      },
      messageCount: 2,
      chunkCount: 3,
    });
    expect(result.breakdowns.conversationsByType).toEqual([{ key: 'private_channel', count: 1 }]);
  });

  it('marks chunkable Slack conversations without chunks as not projected', async () => {
    mockFind.mockReturnValue(
      createLeanFindResult([
        {
          _id: 'conv-2',
          user: 'user-1',
          slackConversationId: 'CENG',
          name: 'engineering',
          conversationType: 'public_channel',
          syncStatus: 'complete',
          messageCount: 4,
          meaningfulMessageCount: 3,
          lastMessageAt: new Date('2026-06-17T13:00:00.000Z'),
          lastMeaningfulMessageAt: new Date('2026-06-17T13:00:00.000Z'),
          updatedAt: new Date('2026-06-17T13:05:00.000Z'),
        },
      ]),
    );
    mockChunkCountDocuments.mockResolvedValue(0);
    resetAggregateDefaults({
      messageStats: [
        {
          _id: 'CENG',
          actualMessageCount: 4,
          chunkableMessageCount: 3,
          skippedMessageCount: 1,
        },
      ],
      chunkStats: [],
      skippedReasons: [{ _id: 'empty_text', count: 1 }],
    });

    const result = await getArchiveDiagnostics({ source: 'slack' });

    expect(result.summary.errorConversationCount).toBe(1);
    expect(result.conversations[0]).toMatchObject({
      sourceConversationId: 'CENG',
      chunkableMessageCount: 3,
      chunkCount: 0,
      health: {
        state: 'not_projected',
        severity: 'error',
      },
    });
    expect(result.breakdowns.skippedMessageReasons).toEqual([{ key: 'empty_text', count: 1 }]);
  });

  it('marks discovered conversations with no stored messages as no messages', async () => {
    mockFind.mockReturnValue(
      createLeanFindResult([
        {
          _id: 'conv-3',
          user: 'user-1',
          slackConversationId: 'CMISSING',
          name: 'missing-history',
          conversationType: 'public_channel',
          syncStatus: 'complete',
          messageCount: 0,
          meaningfulMessageCount: 0,
          updatedAt: new Date('2026-06-17T14:05:00.000Z'),
        },
      ]),
    );
    mockMessageCountDocuments.mockResolvedValue(0);
    mockChunkCountDocuments.mockResolvedValue(0);
    resetAggregateDefaults();

    const result = await getArchiveDiagnostics({ source: 'slack' });

    expect(result.summary.warningConversationCount).toBe(1);
    expect(result.conversations[0]).toMatchObject({
      sourceConversationId: 'CMISSING',
      messageCount: 0,
      health: {
        state: 'no_messages',
        severity: 'warning',
      },
    });
  });
});
