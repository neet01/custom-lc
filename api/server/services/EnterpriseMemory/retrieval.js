const db = require('~/models');

function toArray(value) {
  return Array.isArray(value) ? value : [];
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
    }))
    .filter((participant) => participant.displayName || participant.email);
}

async function searchTeamsMemoryChunks(user, options = {}) {
  if (typeof db.findEnterpriseMemoryChunks !== 'function') {
    return null;
  }

  const userId = user?.id || user?._id?.toString();
  const topic = String(options.topic || options.query || '').trim();
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

  let matchedConversationIds = [];
  if (normalizedChatType !== 'any' || topic || participantClauses.length > 0) {
    const conversationFilter = {
      user: userId,
      ...(normalizedChatType !== 'any' ? { chatType: normalizedChatType } : {}),
      ...((topic || participantClauses.length > 0)
        ? {
            $and: [
              ...(topic ? [{ topic: phraseRegex || termRegexes[0] || buildSearchRegex(topic) }] : []),
              ...participantClauses,
            ],
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
          'No matching memory chunks were found. Consider broadening the topic, participants, or timeframe.',
        results: [],
      };
    }
  }

  const senderClauses = getUserSenderClauses(user);
  const chunkFilter = {
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
    ...(topicChunkClauses.length > 0 ? { $and: topicChunkClauses } : {}),
  };

  const chunks = await db.findEnterpriseMemoryChunks(chunkFilter, {
    limit,
    offset,
    sort: sortBy === 'oldest' ? { sourceTimestamp: 1, orderIndex: 1 } : { sourceTimestamp: -1, orderIndex: -1 },
  });

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
    senderScope: normalizedSenderScope,
    chatType: normalizedChatType,
    daysBack,
    participants: toArray(options.participants).filter(Boolean),
    guidance:
      'These are compact enterprise-memory previews. If one conversation stands out, summarize that conversation before expanding to a bounded message window.',
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

module.exports = {
  searchTeamsMemoryChunks,
};
