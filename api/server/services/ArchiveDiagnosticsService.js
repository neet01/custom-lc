const mongoose = require('mongoose');

const SOURCE_CONFIG = {
  slack: {
    source: 'slack',
    conversationModel: 'SlackArchiveConversation',
    messageModel: 'SlackArchiveMessage',
    syncJobModel: 'SlackArchiveSyncJob',
    parentField: 'slackConversationId',
    typeField: 'conversationType',
    displayFields: ['name', 'topic', 'purpose', 'slackConversationId'],
    sourceRecordTypes: ['slack_message', 'slack_conversation'],
  },
  teams: {
    source: 'teams',
    conversationModel: 'TeamsArchiveConversation',
    messageModel: 'TeamsArchiveMessage',
    syncJobModel: 'TeamsArchiveSyncJob',
    parentField: 'graphChatId',
    typeField: 'chatType',
    displayFields: ['topic', 'graphChatId'],
    sourceRecordTypes: ['teams_message', 'teams_chat'],
  },
};

function getModel(name) {
  return mongoose.models?.[name] || null;
}

function toObjectIdIfValid(value) {
  const stringValue = String(value || '').trim();
  if (!stringValue) {
    return null;
  }
  if (mongoose.Types.ObjectId.isValid(stringValue)) {
    return new mongoose.Types.ObjectId(stringValue);
  }
  return stringValue;
}

function toIso(value) {
  if (!value) {
    return null;
  }
  const date = value instanceof Date ? value : new Date(value);
  return Number.isFinite(date.getTime()) ? date.toISOString() : null;
}

function escapeRegex(value) {
  return String(value || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function clampInteger(value, fallback, { min = 0, max = 1000 } = {}) {
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) {
    return fallback;
  }
  return Math.min(Math.max(Math.floor(parsed), min), max);
}

function getUserMatch(userId) {
  const normalized = String(userId || '').trim();
  if (!normalized) {
    return {};
  }
  return { user: toObjectIdIfValid(normalized) };
}

function getSearchClauses(config, query) {
  const normalized = String(query || '').trim();
  if (!normalized) {
    return [];
  }
  const regex = new RegExp(escapeRegex(normalized), 'i');
  return config.displayFields.map((field) => ({ [field]: regex }));
}

async function countDocuments(model, filter) {
  if (!model) {
    return 0;
  }
  return model.countDocuments(filter);
}

async function findLatestJob(model, filter) {
  if (!model) {
    return null;
  }
  return model.findOne(filter).sort({ createdAt: -1 }).lean();
}

async function aggregateCounts(model, match, groupField) {
  if (!model) {
    return [];
  }
  return model.aggregate([
    { $match: match },
    { $group: { _id: `$${groupField}`, count: { $sum: 1 } } },
    { $sort: { count: -1 } },
  ]);
}

async function aggregateSkipReasons(model, match) {
  if (!model) {
    return [];
  }
  return model.aggregate([
    { $match: { ...match, isChunkable: false } },
    {
      $group: {
        _id: { $ifNull: ['$skipChunkReason', 'unknown'] },
        count: { $sum: 1 },
      },
    },
    { $sort: { count: -1 } },
  ]);
}

async function aggregateMessageStats(model, match, parentField, parentIds) {
  if (!model || parentIds.length === 0) {
    return new Map();
  }
  const results = await model.aggregate([
    { $match: { ...match, [parentField]: { $in: parentIds } } },
    {
      $group: {
        _id: `$${parentField}`,
        actualMessageCount: { $sum: 1 },
        chunkableMessageCount: { $sum: { $cond: ['$isChunkable', 1, 0] } },
        skippedMessageCount: { $sum: { $cond: ['$isChunkable', 0, 1] } },
      },
    },
  ]);

  return new Map(results.map((result) => [result._id, result]));
}

async function aggregateChunkStats(match, parentIds) {
  const EnterpriseMemoryChunk = getModel('EnterpriseMemoryChunk');
  if (!EnterpriseMemoryChunk || parentIds.length === 0) {
    return new Map();
  }
  const results = await EnterpriseMemoryChunk.aggregate([
    { $match: { ...match, sourceParentRecordId: { $in: parentIds } } },
    {
      $group: {
        _id: '$sourceParentRecordId',
        chunkCount: { $sum: 1 },
        messageChunkCount: {
          $sum: { $cond: [{ $eq: ['$chunkType', 'message'] }, 1, 0] },
        },
        windowChunkCount: {
          $sum: { $cond: [{ $eq: ['$chunkType', 'conversation_window'] }, 1, 0] },
        },
        latestChunkAt: { $max: '$sourceTimestamp' },
      },
    },
  ]);

  return new Map(results.map((result) => [result._id, result]));
}

function formatJob(job) {
  if (!job) {
    return null;
  }
  return {
    id: job._id?.toString?.() || job.id || null,
    status: job.status || null,
    phase: job.phase || null,
    jobType: job.jobType || null,
    sourceRecordType: job.sourceRecordType || null,
    errorMessage: job.errorMessage || null,
    stats: job.stats || null,
    createdAt: toIso(job.createdAt),
    startedAt: toIso(job.startedAt),
    completedAt: toIso(job.completedAt),
    updatedAt: toIso(job.updatedAt),
  };
}

function classifyConversationHealth({ conversation, messageStats, chunkStats }) {
  const syncStatus = conversation.syncStatus || 'unknown';
  const messageCount = Number(messageStats?.actualMessageCount ?? conversation.messageCount ?? 0);
  const meaningfulMessageCount = Number(
    conversation.meaningfulMessageCount ?? messageStats?.chunkableMessageCount ?? 0,
  );
  const chunkCount = Number(chunkStats?.chunkCount || 0);
  const latestChunkTime = chunkStats?.latestChunkAt
    ? new Date(chunkStats.latestChunkAt).getTime()
    : 0;
  const latestMeaningfulTime = conversation.lastMeaningfulMessageAt
    ? new Date(conversation.lastMeaningfulMessageAt).getTime()
    : 0;

  if (['failed', 'deferred_failed'].includes(syncStatus)) {
    return {
      state: 'sync_failed',
      severity: 'error',
      reason: conversation.syncError || syncStatus,
    };
  }
  if (syncStatus === 'running' || syncStatus === 'pending') {
    return { state: syncStatus, severity: 'warning', reason: `Conversation sync is ${syncStatus}` };
  }
  if (messageCount === 0) {
    return {
      state: 'no_messages',
      severity: 'warning',
      reason: 'No archived messages were stored.',
    };
  }
  if (meaningfulMessageCount === 0) {
    return {
      state: 'no_chunkable_messages',
      severity: 'warning',
      reason: 'Messages exist, but none were marked chunkable.',
    };
  }
  if (chunkCount === 0) {
    return {
      state: 'not_projected',
      severity: 'error',
      reason: 'Chunkable messages exist, but no EnterpriseMemory chunks were found.',
    };
  }
  if (latestMeaningfulTime > 0 && latestChunkTime > 0 && latestChunkTime < latestMeaningfulTime) {
    return {
      state: 'stale_projection',
      severity: 'warning',
      reason: 'Archive has newer meaningful messages than the latest indexed chunk.',
    };
  }
  return {
    state: 'healthy',
    severity: 'ok',
    reason: 'Archive messages are projected into chunks.',
  };
}

function formatConversation({ source, config, conversation, messageStats, chunkStats }) {
  const health = classifyConversationHealth({ conversation, messageStats, chunkStats });
  const parentId = conversation[config.parentField];
  const displayName =
    source === 'slack'
      ? conversation.name || conversation.topic || parentId
      : conversation.topic || parentId;

  return {
    id: conversation._id?.toString?.() || conversation.id || null,
    userId: conversation.user?.toString?.() || String(conversation.user || ''),
    source,
    sourceConversationId: parentId,
    displayName,
    type: conversation[config.typeField] || 'unknown',
    syncStatus: conversation.syncStatus || 'unknown',
    syncError: conversation.syncError || conversation.syncDeferredReason || '',
    messageCount: Number(messageStats?.actualMessageCount ?? conversation.messageCount ?? 0),
    declaredMessageCount: Number(conversation.messageCount || 0),
    meaningfulMessageCount: Number(
      conversation.meaningfulMessageCount ?? messageStats?.chunkableMessageCount ?? 0,
    ),
    chunkableMessageCount: Number(messageStats?.chunkableMessageCount || 0),
    skippedMessageCount: Number(messageStats?.skippedMessageCount || 0),
    chunkCount: Number(chunkStats?.chunkCount || 0),
    messageChunkCount: Number(chunkStats?.messageChunkCount || 0),
    windowChunkCount: Number(chunkStats?.windowChunkCount || 0),
    lastMessageAt: toIso(conversation.lastMessageAt),
    lastMeaningfulMessageAt: toIso(conversation.lastMeaningfulMessageAt),
    lastMessageSyncAt: toIso(conversation.lastMessageSyncAt),
    latestChunkAt: toIso(chunkStats?.latestChunkAt),
    updatedAt: toIso(conversation.updatedAt),
    health,
  };
}

function toCountItems(results) {
  return results.map((result) => ({
    key: result._id == null || result._id === '' ? 'unknown' : String(result._id),
    count: Number(result.count || 0),
  }));
}

async function getArchiveDiagnostics(options = {}) {
  const source = SOURCE_CONFIG[options.source] ? options.source : 'slack';
  const config = SOURCE_CONFIG[source];
  const Conversation = getModel(config.conversationModel);
  const Message = getModel(config.messageModel);
  const SyncJob = getModel(config.syncJobModel);
  const EnterpriseMemoryChunk = getModel('EnterpriseMemoryChunk');
  const EnterpriseMemoryJob = getModel('EnterpriseMemoryJob');
  const limit = clampInteger(options.limit, 50, { min: 1, max: 200 });
  const offset = clampInteger(options.offset, 0, { min: 0, max: 100000 });
  const userMatch = getUserMatch(options.userId);
  const searchClauses = getSearchClauses(config, options.q);
  const type = String(options.type || '').trim();
  const status = String(options.status || '').trim();
  const conversationFilter = {
    ...userMatch,
    ...(type ? { [config.typeField]: type } : {}),
    ...(status ? { syncStatus: status } : {}),
    ...(searchClauses.length > 0 ? { $or: searchClauses } : {}),
  };
  const baseConversationFilter = { ...userMatch };
  const baseMessageFilter = { ...userMatch };
  const baseChunkFilter = {
    ...userMatch,
    source,
    sourceRecordType: { $in: config.sourceRecordTypes },
  };

  const conversations = Conversation
    ? await Conversation.find(conversationFilter)
        .sort({ lastMeaningfulMessageAt: -1, lastMessageAt: -1, updatedAt: -1 })
        .skip(offset)
        .limit(limit)
        .lean()
    : [];

  const parentIds = conversations
    .map((conversation) => conversation[config.parentField])
    .filter(Boolean);
  const [messageStats, chunkStats] = await Promise.all([
    aggregateMessageStats(Message, baseMessageFilter, config.parentField, parentIds),
    aggregateChunkStats(baseChunkFilter, parentIds),
  ]);

  const [
    totalConversations,
    filteredConversations,
    totalMessages,
    totalChunks,
    latestSync,
    latestProjection,
    conversationsByType,
    conversationsByStatus,
    chunksByRecordType,
    chunksByChunkType,
    skippedMessageReasons,
  ] = await Promise.all([
    countDocuments(Conversation, baseConversationFilter),
    countDocuments(Conversation, conversationFilter),
    countDocuments(Message, baseMessageFilter),
    countDocuments(EnterpriseMemoryChunk, baseChunkFilter),
    findLatestJob(SyncJob, baseConversationFilter),
    findLatestJob(EnterpriseMemoryJob, { ...userMatch, source, jobType: 'projection' }),
    aggregateCounts(Conversation, baseConversationFilter, config.typeField),
    aggregateCounts(Conversation, baseConversationFilter, 'syncStatus'),
    aggregateCounts(EnterpriseMemoryChunk, baseChunkFilter, 'sourceRecordType'),
    aggregateCounts(EnterpriseMemoryChunk, baseChunkFilter, 'chunkType'),
    aggregateSkipReasons(Message, baseMessageFilter),
  ]);

  const rows = conversations.map((conversation) =>
    formatConversation({
      source,
      config,
      conversation,
      messageStats: messageStats.get(conversation[config.parentField]),
      chunkStats: chunkStats.get(conversation[config.parentField]),
    }),
  );

  return {
    source,
    generatedAt: new Date().toISOString(),
    filters: {
      userId: String(options.userId || '').trim() || null,
      q: String(options.q || '').trim() || null,
      type: type || null,
      status: status || null,
      limit,
      offset,
    },
    summary: {
      totalConversations,
      filteredConversations,
      totalMessages,
      totalChunks,
      healthyConversationCount: rows.filter((row) => row.health.severity === 'ok').length,
      warningConversationCount: rows.filter((row) => row.health.severity === 'warning').length,
      errorConversationCount: rows.filter((row) => row.health.severity === 'error').length,
    },
    breakdowns: {
      conversationsByType: toCountItems(conversationsByType),
      conversationsByStatus: toCountItems(conversationsByStatus),
      chunksByRecordType: toCountItems(chunksByRecordType),
      chunksByChunkType: toCountItems(chunksByChunkType),
      skippedMessageReasons: toCountItems(skippedMessageReasons),
    },
    latestSync: formatJob(latestSync),
    latestProjection: formatJob(latestProjection),
    conversations: rows,
  };
}

module.exports = {
  getArchiveDiagnostics,
};
