const db = require('~/models');

function toArray(value) {
  return Array.isArray(value) ? value : [];
}

function looksLikeMongoObjectId(value) {
  return /^[a-f0-9]{24}$/i.test(String(value || '').trim());
}

function clampInteger(value, fallback, { min = 1, max = 500 } = {}) {
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) {
    return fallback;
  }
  return Math.min(Math.max(Math.trunc(parsed), min), max);
}

function escapeRegex(value) {
  return String(value || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function buildSearchRegex(value) {
  const normalized = String(value || '')
    .trim()
    .replace(/\s+/g, ' ');

  if (!normalized) {
    return null;
  }

  return new RegExp(escapeRegex(normalized).replace(/\\ /g, '\\s+'), 'i');
}

function truncateText(value, max = 280) {
  const normalized = String(value || '').replace(/\s+/g, ' ').trim();
  if (!normalized) {
    return '';
  }

  if (normalized.length <= max) {
    return normalized;
  }

  return `${normalized.slice(0, Math.max(0, max - 1)).trimEnd()}…`;
}

function buildTopicTerms(value) {
  const stopWords = new Set([
    'what',
    'has',
    'have',
    'been',
    'about',
    'with',
    'from',
    'into',
    'that',
    'this',
    'those',
    'these',
    'discussion',
    'discussed',
    'messages',
    'message',
    'chat',
    'chats',
    'teams',
    'recently',
    'recent',
    'show',
    'find',
    'look',
    'search',
    'around',
    'regarding',
  ]);

  return [...new Set(
    String(value || '')
      .toLowerCase()
      .split(/[^a-z0-9._-]+/i)
      .map((term) => term.trim())
      .filter((term) => term.length >= 2 && !stopWords.has(term)),
  )].slice(0, 8);
}

function buildChunkSearchClauses(value) {
  const phraseRegex = buildSearchRegex(value);
  const termRegexes = buildTopicTerms(value).map((term) => buildSearchRegex(term)).filter(Boolean);
  const fields = ['text', 'summary', 'title'];
  const fieldOr = (regex) => ({
    $or: fields.map((field) => ({ [field]: regex })),
  });

  const clauses = [];
  if (phraseRegex) {
    clauses.push(fieldOr(phraseRegex));
  }

  if (termRegexes.length > 1) {
    clauses.push({
      $and: termRegexes.map((regex) => fieldOr(regex)),
    });
  }

  return { phraseRegex, termRegexes, clauses };
}

function isTextSearchEnabled() {
  return /^(true|1|yes|on)$/i.test(String(process.env.TEAMS_ARCHIVE_TEXT_SEARCH_ENABLED || '').trim());
}

function isSlackTextSearchEnabled() {
  return !/^(false|0|no|off)$/i.test(String(process.env.SLACK_ARCHIVE_TEXT_SEARCH_ENABLED || '').trim());
}

function buildTextSearchString(value) {
  const terms = buildTopicTerms(value);
  if (terms.length > 0) {
    return terms.join(' ');
  }
  return String(value || '').trim();
}

function getUserSenderClauses(user) {
  const normalizedEmail = String(user?.email || '')
    .trim()
    .toLowerCase();
  const normalizedName = String(user?.name || '')
    .trim();
  const normalizedUsername = String(user?.username || '')
    .trim();
  const openidId = String(user?.openidId || '')
    .trim();

  return [
    ...(openidId ? [{ 'metadata.fromUserId': openidId }] : []),
    ...(normalizedEmail
      ? [{ 'metadata.fromEmail': normalizedEmail }, { 'metadata.fromEmail': user?.email }]
      : []),
    ...(normalizedName ? [{ 'metadata.fromDisplayName': normalizedName }] : []),
    ...(normalizedUsername ? [{ 'metadata.fromDisplayName': normalizedUsername }] : []),
  ];
}

function buildParticipantConversationClauses(participants = []) {
  return toArray(participants)
    .map((participant) => String(participant || '').trim())
    .filter(Boolean)
    .slice(0, 10)
    .map((participant) => {
      const regex = buildSearchRegex(participant);
      return {
        $or: [{ 'participants.displayName': regex }, { 'participants.email': regex }],
      };
    });
}

function mapCompactParticipants(participants = [], max = 4) {
  return toArray(participants)
    .slice(0, max)
    .map((participant) => ({
      displayName: participant?.displayName || '',
      email: participant?.email || '',
      slackUserId: participant?.slackUserId || '',
    }))
    .filter((participant) => participant.displayName || participant.email || participant.slackUserId);
}

function getSlackSenderClauses(user) {
  const normalizedName = String(user?.name || '').trim();
  const normalizedUsername = String(user?.username || '').trim();

  return [
    ...(normalizedName ? [{ 'metadata.displayName': normalizedName }] : []),
    ...(normalizedUsername ? [{ 'metadata.username': normalizedUsername }] : []),
  ];
}

function buildSlackParticipantConversationClauses(participants = []) {
  return toArray(participants)
    .map((participant) => String(participant || '').trim())
    .filter(Boolean)
    .slice(0, 10)
    .map((participant) => {
      const regex = buildSearchRegex(participant);
      return {
        $or: [
          { name: regex },
          { topic: regex },
          { purpose: regex },
          { 'participants.displayName': regex },
          { 'participants.realName': regex },
          { 'participants.username': regex },
          { 'participants.email': regex },
          { 'participants.slackUserId': participant },
        ],
      };
    });
}

async function searchTeamsMemoryChunks(user, options = {}) {
  if (typeof db.findEnterpriseMemoryChunks !== 'function') {
    return null;
  }

  const userId = user?.id || user?._id?.toString();
  const topic = String(options.topic || options.query || '').trim();
  const chatId = String(options.chatId || options.graphChatId || '').trim();
  const scopedGraphChatId = chatId && !looksLikeMongoObjectId(chatId) ? chatId : '';
  const senderScope = String(options.senderScope || 'any').trim();
  const chatType = String(options.chatType || 'any').trim();
  const sortBy = String(options.sortBy || 'recent').trim();
  const limit = clampInteger(options.limit, 8, { max: 12 });
  const offset = clampInteger(options.offset, 0, { min: 0, max: 100000 });
  const daysBack = options.daysBack
    ? clampInteger(options.daysBack, 30, { min: 1, max: 3650 })
    : undefined;

  const validSenderScopes = new Set(['any', 'me', 'others']);
  const validChatTypes = new Set(['any', 'oneOnOne', 'group', 'meeting']);
  const normalizedSenderScope = validSenderScopes.has(senderScope) ? senderScope : 'any';
  const normalizedChatType = validChatTypes.has(chatType) ? chatType : 'any';

  const participantClauses = buildParticipantConversationClauses(options.participants);
  const { phraseRegex, termRegexes, clauses: topicChunkClauses } = buildChunkSearchClauses(topic);

  let matchedConversationIds = scopedGraphChatId ? [scopedGraphChatId] : [];
  const shouldPrefilterConversations =
    !scopedGraphChatId && (normalizedChatType !== 'any' || participantClauses.length > 0);

  if (shouldPrefilterConversations) {
    const conversationFilter = {
      user: userId,
      ...(normalizedChatType !== 'any' ? { chatType: normalizedChatType } : {}),
      ...(participantClauses.length > 0
        ? {
            $and: [...participantClauses],
          }
        : {}),
    };

    const matchedConversations = await db.findTeamsArchiveConversations(conversationFilter, {
      limit: 1000,
    });
    matchedConversationIds = matchedConversations
      .map((conversation) => conversation.graphChatId)
      .filter(Boolean);

    if (matchedConversationIds.length === 0) {
      return {
        retrievalMode: 'enterprise_memory',
        topic: topic || undefined,
        senderScope: normalizedSenderScope,
        chatType: normalizedChatType,
        daysBack,
        participants: toArray(options.participants).filter(Boolean),
        guidance:
          'No matching memory chunks were found. Consider broadening the participants, chat type, or timeframe.',
        trace: {
          backend: 'enterprise_memory',
          conversationPrefilterApplied: true,
          topicPrefilterApplied: false,
          matchedConversationCount: 0,
        },
        results: [],
      };
    }
  }

  const senderClauses = getUserSenderClauses(user);
  const baseChunkFilter = {
    user: userId,
    source: 'teams',
    sourceRecordType: 'teams_message',
    ...(matchedConversationIds.length > 0 ? { sourceParentRecordId: { $in: matchedConversationIds } } : {}),
    ...(daysBack ? { sourceTimestamp: { $gte: new Date(Date.now() - daysBack * 24 * 60 * 60 * 1000) } } : {}),
    ...(normalizedSenderScope === 'me'
      ? senderClauses.length > 0
        ? { $or: senderClauses }
        : {}
      : normalizedSenderScope === 'others'
        ? senderClauses.length > 0
          ? { $nor: senderClauses }
          : {}
        : {}),
  };

  const regexChunkFilter = {
    ...baseChunkFilter,
    ...(topicChunkClauses.length > 0 ? { $and: topicChunkClauses } : {}),
  };

  const recencySort =
    sortBy === 'oldest' ? { sourceTimestamp: 1, orderIndex: 1 } : { sourceTimestamp: -1, orderIndex: -1 };
  const textSearchString = buildTextSearchString(topic);
  const useTextSearch =
    isTextSearchEnabled() && Boolean(textSearchString) && topicChunkClauses.length > 0;

  let chunks;
  let searchBackend = 'regex';
  let textSearchFellBack = false;

  if (useTextSearch) {
    const textChunkFilter = { ...baseChunkFilter, $text: { $search: textSearchString } };
    chunks = await db.findEnterpriseMemoryChunks(textChunkFilter, {
      limit,
      offset,
      sort: sortBy === 'oldest' ? { sourceTimestamp: 1 } : { sourceTimestamp: -1 },
      textScore: sortBy !== 'oldest',
    });
    searchBackend = 'text';

    if (chunks.length === 0) {
      chunks = await db.findEnterpriseMemoryChunks(regexChunkFilter, { limit, offset, sort: recencySort });
      searchBackend = 'regex';
      textSearchFellBack = true;
    }
  } else {
    chunks = await db.findEnterpriseMemoryChunks(regexChunkFilter, { limit, offset, sort: recencySort });
  }

  const conversationIds = [...new Set(chunks.map((chunk) => chunk.sourceParentRecordId).filter(Boolean))];
  const conversations = conversationIds.length
    ? await db.findTeamsArchiveConversations(
        { user: userId, graphChatId: { $in: conversationIds } },
        { limit: Math.max(conversationIds.length, 1) },
      )
    : [];
  const conversationMap = new Map(
    conversations.map((conversation) => [conversation.graphChatId, conversation]),
  );

  return {
    retrievalMode: 'enterprise_memory',
    topic: topic || undefined,
    ...(scopedGraphChatId ? { chatId, graphChatId: scopedGraphChatId } : {}),
    senderScope: normalizedSenderScope,
    chatType: normalizedChatType,
    daysBack,
    participants: toArray(options.participants).filter(Boolean),
    guidance:
      'These are compact enterprise-memory previews. If one conversation stands out, summarize that conversation before expanding to a bounded message window.',
    trace: {
      backend: 'enterprise_memory',
      searchBackend,
      textSearchFellBack,
      conversationPrefilterApplied: shouldPrefilterConversations,
      chatIdScoped: Boolean(scopedGraphChatId),
      topicPrefilterApplied: false,
      matchedConversationCount: matchedConversationIds.length,
    },
    resultCount: chunks.length,
    results: chunks.map((chunk) => {
      const conversation = conversationMap.get(chunk.sourceParentRecordId);
      return {
        id: chunk._id?.toString?.() || chunk.id,
        graphMessageId: chunk.sourceRecordId,
        graphChatId: chunk.sourceParentRecordId,
        topic: conversation?.topic || '',
        chatType: conversation?.chatType || '',
        participants: mapCompactParticipants(conversation?.participants || []),
        fromDisplayName: chunk?.metadata?.fromDisplayName || '',
        fromEmail: chunk?.metadata?.fromEmail || '',
        subject: chunk?.metadata?.subject || '',
        summary: truncateText(chunk.summary || '', 180),
        excerpt: truncateText(chunk.summary || chunk.text || '', 260),
        sentDateTime: chunk.sourceTimestamp,
        webUrl: chunk?.metadata?.webUrl || '',
      };
    }),
  };
}

async function searchSlackMemoryChunks(user, options = {}) {
  if (typeof db.findEnterpriseMemoryChunks !== 'function') {
    return null;
  }

  const userId = user?.id || user?._id?.toString();
  const topic = String(options.topic || options.query || '').trim();
  const conversationId = String(
    options.conversationId || options.channelId || options.slackConversationId || '',
  ).trim();
  const senderScope = String(options.senderScope || 'any').trim();
  const senderUserId = String(options.senderUserId || options.slackUserId || '').trim();
  const conversationType = String(options.conversationType || options.channelType || 'any').trim();
  const sortBy = String(options.sortBy || 'relevance').trim();
  const limit = clampInteger(options.limit, 8, { max: 12 });
  const offset = clampInteger(options.offset, 0, { min: 0, max: 100000 });
  const daysBack = options.daysBack
    ? clampInteger(options.daysBack, 30, { min: 1, max: 3650 })
    : undefined;

  const validSenderScopes = new Set(['any', 'me', 'others']);
  const validConversationTypes = new Set(['any', 'public_channel', 'private_channel', 'im', 'mpim']);
  const normalizedSenderScope = validSenderScopes.has(senderScope) ? senderScope : 'any';
  const normalizedConversationType = validConversationTypes.has(conversationType)
    ? conversationType
    : 'any';

  const participantClauses = buildSlackParticipantConversationClauses(options.participants);
  const { clauses: topicChunkClauses } = buildChunkSearchClauses(topic);

  let matchedConversationIds = conversationId ? [conversationId] : [];
  const shouldPrefilterConversations =
    !conversationId && (normalizedConversationType !== 'any' || participantClauses.length > 0);

  if (shouldPrefilterConversations) {
    const conversationFilter = {
      user: userId,
      ...(normalizedConversationType !== 'any' ? { conversationType: normalizedConversationType } : {}),
      ...(participantClauses.length > 0 ? { $and: participantClauses } : {}),
    };

    const matchedConversations = await db.findSlackArchiveConversations(conversationFilter, {
      limit: 1000,
    });
    matchedConversationIds = matchedConversations
      .map((conversation) => conversation.slackConversationId)
      .filter(Boolean);

    if (matchedConversationIds.length === 0) {
      return {
        retrievalMode: 'enterprise_memory',
        source: 'slack',
        topic: topic || undefined,
        senderScope: normalizedSenderScope,
        conversationType: normalizedConversationType,
        daysBack,
        participants: toArray(options.participants).filter(Boolean),
        guidance:
          'No matching Slack memory chunks were found. Broaden the participants, channel type, or timeframe.',
        trace: {
          backend: 'enterprise_memory',
          searchBackend: 'text',
          textSearchEnabled: isSlackTextSearchEnabled(),
          conversationPrefilterApplied: true,
          conversationIdScoped: false,
          matchedConversationCount: 0,
        },
        resultCount: 0,
        results: [],
      };
    }
  }

  const senderClauses = senderUserId
    ? [{ 'metadata.slackUserId': senderUserId }]
    : getSlackSenderClauses(user);
  const baseChunkFilter = {
    user: userId,
    source: 'slack',
    sourceRecordType: 'slack_message',
    ...(matchedConversationIds.length > 0
      ? { sourceParentRecordId: { $in: matchedConversationIds } }
      : {}),
    ...(daysBack ? { sourceTimestamp: { $gte: new Date(Date.now() - daysBack * 24 * 60 * 60 * 1000) } } : {}),
    ...(senderUserId
      ? { 'metadata.slackUserId': senderUserId }
      : normalizedSenderScope === 'me'
        ? senderClauses.length > 0
          ? { $or: senderClauses }
          : {}
        : normalizedSenderScope === 'others'
          ? senderClauses.length > 0
            ? { $nor: senderClauses }
            : {}
          : {}),
  };

  const textSearchString = buildTextSearchString(topic);
  const textSearchEnabled = isSlackTextSearchEnabled();
  const recencySort =
    sortBy === 'oldest' ? { sourceTimestamp: 1, orderIndex: 1 } : { sourceTimestamp: -1, orderIndex: -1 };
  const regexChunkFilter = {
    ...baseChunkFilter,
    ...(topicChunkClauses.length > 0 ? { $and: topicChunkClauses } : {}),
  };

  let chunks = [];
  let searchBackend = 'none';
  let textSearchFellBack = false;

  if (textSearchEnabled && textSearchString) {
    const textChunkFilter = { ...baseChunkFilter, $text: { $search: textSearchString } };
    chunks = await db.findEnterpriseMemoryChunks(textChunkFilter, {
      limit,
      offset,
      sort: sortBy === 'oldest' ? { sourceTimestamp: 1 } : { sourceTimestamp: -1 },
      textScore: sortBy !== 'oldest' && sortBy !== 'recent',
    });
    searchBackend = 'text';

    if (chunks.length === 0 && topicChunkClauses.length > 0) {
      chunks = await db.findEnterpriseMemoryChunks(regexChunkFilter, { limit, offset, sort: recencySort });
      searchBackend = 'regex';
      textSearchFellBack = true;
    }
  } else if (topicChunkClauses.length > 0) {
    chunks = await db.findEnterpriseMemoryChunks(regexChunkFilter, { limit, offset, sort: recencySort });
    searchBackend = 'regex';
  } else {
    chunks = await db.findEnterpriseMemoryChunks(baseChunkFilter, { limit, offset, sort: recencySort });
    searchBackend = 'recency';
  }

  const conversationIds = [...new Set(chunks.map((chunk) => chunk.sourceParentRecordId).filter(Boolean))];
  const conversations = conversationIds.length
    ? await db.findSlackArchiveConversations(
        { user: userId, slackConversationId: { $in: conversationIds } },
        { limit: Math.max(conversationIds.length, 1) },
      )
    : [];
  const conversationMap = new Map(
    conversations.map((conversation) => [conversation.slackConversationId, conversation]),
  );

  return {
    retrievalMode: 'enterprise_memory',
    source: 'slack',
    topic: topic || undefined,
    ...(conversationId ? { conversationId } : {}),
    senderScope: normalizedSenderScope,
    ...(senderUserId ? { senderUserId } : {}),
    conversationType: normalizedConversationType,
    daysBack,
    participants: toArray(options.participants).filter(Boolean),
    guidance:
      'These are indexed Slack enterprise-memory results. Prefer these for semantic/hybrid discovery; use get_messages for exact thread/channel context after selecting a conversation.',
    trace: {
      backend: 'enterprise_memory',
      searchBackend,
      textSearchEnabled,
      textSearchFellBack,
      conversationPrefilterApplied: shouldPrefilterConversations,
      conversationIdScoped: Boolean(conversationId),
      matchedConversationCount: matchedConversationIds.length,
    },
    resultCount: chunks.length,
    results: chunks.map((chunk) => {
      const conversation = conversationMap.get(chunk.sourceParentRecordId);
      return {
        id: chunk._id?.toString?.() || chunk.id,
        slackMessageTs: chunk.sourceRecordId,
        slackConversationId: chunk.sourceParentRecordId,
        conversationId: chunk.sourceParentRecordId,
        conversationName: conversation?.name || '',
        topic: conversation?.topic || '',
        conversationType: conversation?.conversationType || chunk?.metadata?.conversationType || '',
        participants: mapCompactParticipants(conversation?.participants || []),
        fromDisplayName: chunk?.metadata?.displayName || chunk?.metadata?.username || '',
        slackUserId: chunk?.metadata?.slackUserId || '',
        threadTs: chunk?.metadata?.threadTs || '',
        summary: truncateText(chunk.summary || '', 180),
        excerpt: truncateText(chunk.summary || chunk.text || '', 320),
        sentAt: chunk.sourceTimestamp,
      };
    }),
  };
}

module.exports = {
  searchTeamsMemoryChunks,
  searchSlackMemoryChunks,
};
