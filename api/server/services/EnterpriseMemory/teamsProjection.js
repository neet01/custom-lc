const { logger } = require('@librechat/data-schemas');
const db = require('~/models');

const PROJECTION_MESSAGE_FETCH_LIMIT = 5000;
const CONVERSATION_WINDOW_MESSAGE_TARGET = 20;
const CONVERSATION_WINDOW_MAX_TIME_GAP_MS = 12 * 60 * 60 * 1000;

function uniqueStrings(values = []) {
  const seen = new Set();
  const results = [];

  for (const value of values) {
    const normalized = String(value || '').trim();
    if (!normalized) {
      continue;
    }
    const key = normalized.toLowerCase();
    if (seen.has(key)) {
      continue;
    }
    seen.add(key);
    results.push(normalized);
  }

  return results;
}

function buildPersonCanonicalKey({ userId, email, displayName }) {
  if (userId) {
    return `aad:${String(userId).trim().toLowerCase()}`;
  }

  if (email) {
    return `email:${String(email).trim().toLowerCase()}`;
  }

  return `name:${String(displayName || 'unknown').trim().toLowerCase()}`;
}

function buildParticipantMap(conversation, messages) {
  const participants = new Map();

  const addParticipant = (participant = {}) => {
    const displayName = String(participant.displayName || '').trim();
    const email = String(participant.email || '').trim();
    const userId = String(participant.userId || participant.mentionedUserId || '').trim();
    const canonicalKey = buildPersonCanonicalKey({ userId, email, displayName });

    const existing = participants.get(canonicalKey) || {
      canonicalKey,
      displayName: '',
      email: '',
      userId: '',
      aliases: [],
    };

    existing.displayName = existing.displayName || displayName || email || userId || canonicalKey;
    existing.email = existing.email || email;
    existing.userId = existing.userId || userId;
    existing.aliases = uniqueStrings([...(existing.aliases || []), displayName, email]);
    participants.set(canonicalKey, existing);
  };

  for (const participant of conversation?.participants || []) {
    addParticipant(participant);
  }

  for (const message of messages) {
    addParticipant({
      displayName: message?.fromDisplayName,
      email: message?.fromEmail,
      userId: message?.fromUserId,
    });

    for (const mention of message?.mentions || []) {
      addParticipant({
        displayName: mention?.displayName,
        userId: mention?.mentionedUserId,
      });
    }
  }

  return participants;
}

function formatConversationTitle(conversation) {
  if (conversation?.topic) {
    return conversation.topic;
  }

  const participants = uniqueStrings(
    (conversation?.participants || []).map((participant) => participant?.displayName || participant?.email),
  );

  if (participants.length > 0) {
    return participants.slice(0, 4).join(', ');
  }

  return `Teams chat ${conversation?.graphChatId || ''}`.trim();
}

function getChunkableMessageText(message) {
  return String(message?.bodyText || message?.summary || message?.subject || '').trim();
}

function pushConversationWindow(currentWindow, windows) {
  if (currentWindow.length > 0) {
    windows.push(currentWindow);
  }
}

function buildConversationWindows(sortedMessages = []) {
  const chunkableMessages = sortedMessages.filter((message) => {
    if (typeof message?.isChunkable === 'boolean') {
      return message.isChunkable;
    }
    return Boolean(getChunkableMessageText(message));
  });

  const windows = [];
  let currentWindow = [];

  for (const message of chunkableMessages) {
    const lastMessage = currentWindow[currentWindow.length - 1];
    const lastTime = lastMessage?.sentDateTime ? new Date(lastMessage.sentDateTime).getTime() : 0;
    const currentTime = message?.sentDateTime ? new Date(message.sentDateTime).getTime() : 0;
    const timeGapExceeded =
      lastMessage &&
      currentTime > 0 &&
      lastTime > 0 &&
      currentTime - lastTime > CONVERSATION_WINDOW_MAX_TIME_GAP_MS;

    if (
      currentWindow.length >= CONVERSATION_WINDOW_MESSAGE_TARGET ||
      timeGapExceeded
    ) {
      pushConversationWindow(currentWindow, windows);
      currentWindow = [];
    }

    currentWindow.push(message);
  }

  pushConversationWindow(currentWindow, windows);
  return windows;
}

function formatConversationWindowText(messages = []) {
  return messages
    .map((message) => {
      const timestamp = message?.sentDateTime?.toISOString?.() || '';
      const sender = String(message?.fromDisplayName || message?.fromEmail || 'Unknown').trim();
      const text = getChunkableMessageText(message);
      return `[${timestamp}] ${sender}: ${text}`.trim();
    })
    .join('\n');
}

async function projectTeamsConversationToMemory({
  userId,
  tenantId,
  visibilityScope = 'user',
  conversation,
  messages,
}) {
  const conversationEntity = await db.upsertEnterpriseMemoryEntity({
    user: userId,
    tenantId,
    visibilityScope,
    source: 'teams',
    entityType: 'conversation',
    canonicalKey: `teams_chat:${conversation.graphChatId}`,
    displayName: formatConversationTitle(conversation),
    aliases: uniqueStrings([conversation.topic, conversation.graphChatId]),
    summary: `Teams ${conversation.chatType || 'chat'} conversation`,
    sourceRecordType: 'teams_chat',
    sourceRecordId: conversation.graphChatId,
    sourceUpdatedAt: conversation.sourceUpdatedAt || conversation.updatedAt || conversation.lastMessageAt,
    attributes: {
      chatType: conversation.chatType || '',
      webUrl: conversation.webUrl || '',
      messageCount: conversation.messageCount || 0,
      participantCount: Array.isArray(conversation.participants) ? conversation.participants.length : 0,
      lastMessageAt: conversation.lastMessageAt || null,
    },
  });

  const participants = buildParticipantMap(conversation, messages);
  const participantEntityMap = new Map();

  for (const participant of participants.values()) {
    const entity = await db.upsertEnterpriseMemoryEntity({
      user: userId,
      tenantId,
      visibilityScope,
      source: 'teams',
      entityType: 'person',
      canonicalKey: participant.canonicalKey,
      displayName: participant.displayName,
      aliases: participant.aliases,
      summary: participant.email || participant.userId || '',
      sourceRecordType: 'teams_participant',
      sourceRecordId: participant.userId || participant.email || participant.displayName,
      sourceParentRecordId: conversation.graphChatId,
      sourceUpdatedAt: conversation.sourceUpdatedAt || conversation.updatedAt || conversation.lastMessageAt,
      attributes: {
        email: participant.email || '',
        aadUserId: participant.userId || '',
      },
    });

    participantEntityMap.set(participant.canonicalKey, entity);
  }

  const relationships = [];

  for (const participant of participants.values()) {
    const entity = participantEntityMap.get(participant.canonicalKey);
    if (!entity?._id) {
      continue;
    }

    relationships.push({
      user: userId,
      tenantId,
      visibilityScope,
      source: 'teams',
      relationshipType: 'has_participant',
      fromEntityId: conversationEntity._id.toString(),
      toEntityId: entity._id.toString(),
      sourceRecordType: 'teams_chat',
      sourceRecordId: conversation.graphChatId,
      sourceUpdatedAt: conversation.sourceUpdatedAt || conversation.updatedAt || conversation.lastMessageAt,
      attributes: {
        chatType: conversation.chatType || '',
      },
    });
  }

  const sortedMessages = [...messages].sort((a, b) => {
    const aTime = a?.sentDateTime ? new Date(a.sentDateTime).getTime() : 0;
    const bTime = b?.sentDateTime ? new Date(b.sentDateTime).getTime() : 0;
    return aTime - bTime;
  });

  const chunks = [];
  let skippedEmptyTextMessages = 0;
  const skippedReasonCounts = {};
  let messageChunkCount = 0;

  for (const [index, message] of sortedMessages.entries()) {
    const text = getChunkableMessageText(message);
    const messageIsChunkable =
      typeof message?.isChunkable === 'boolean' ? Boolean(message.isChunkable) : Boolean(text);

    if (!messageIsChunkable || !text) {
      skippedEmptyTextMessages += 1;
      const skipReason = String(message?.skipChunkReason || 'empty_normalized_text').trim() || 'empty_normalized_text';
      skippedReasonCounts[skipReason] = (skippedReasonCounts[skipReason] || 0) + 1;
      continue;
    }

    const senderKey = buildPersonCanonicalKey({
      userId: message?.fromUserId,
      email: message?.fromEmail,
      displayName: message?.fromDisplayName,
    });

    const senderEntity = participantEntityMap.get(senderKey);
    const mentionEntityIds = (message?.mentions || [])
      .map((mention) =>
        participantEntityMap.get(
          buildPersonCanonicalKey({
            userId: mention?.mentionedUserId,
            displayName: mention?.displayName,
          }),
        ),
      )
      .filter(Boolean)
      .map((entity) => entity._id.toString());

    chunks.push({
      user: userId,
      tenantId,
      visibilityScope,
      source: 'teams',
      sourceRecordType: 'teams_message',
      sourceRecordId: message.graphMessageId,
      sourceParentRecordId: conversation.graphChatId,
      parentEntityId: conversationEntity._id.toString(),
      entityIds: uniqueStrings([
        conversationEntity._id.toString(),
        senderEntity?._id?.toString(),
        ...mentionEntityIds,
      ]),
      chunkType: 'message',
      title: `${message?.fromDisplayName || 'Unknown'} @ ${message?.sentDateTime?.toISOString?.() || ''}`.trim(),
      text,
      summary: String(message?.bodyPreview || '').trim().slice(0, 512),
      orderIndex: index,
      sourceTimestamp: message?.sentDateTime,
      metadata: {
        graphChatId: conversation.graphChatId,
        graphMessageId: message.graphMessageId,
        replyToId: message.replyToId || '',
        fromDisplayName: message.fromDisplayName || '',
        fromEmail: message.fromEmail || '',
        fromUserId: message.fromUserId || '',
        subject: message.subject || '',
        webUrl: message.webUrl || '',
        importance: message.importance || '',
        messageType: message.messageType || '',
        attachmentCount: Array.isArray(message.attachments) ? message.attachments.length : 0,
        mentionCount: Array.isArray(message.mentions) ? message.mentions.length : 0,
        chatType: conversation.chatType || '',
      },
    });
    messageChunkCount += 1;
  }

  const conversationWindows = buildConversationWindows(sortedMessages);
  let conversationWindowChunkCount = 0;
  for (const [windowIndex, windowMessages] of conversationWindows.entries()) {
    const windowText = formatConversationWindowText(windowMessages);
    if (!windowText.trim()) {
      continue;
    }

    const windowEntityIds = uniqueStrings([
      conversationEntity._id.toString(),
      ...windowMessages.flatMap((message) => {
        const senderKey = buildPersonCanonicalKey({
          userId: message?.fromUserId,
          email: message?.fromEmail,
          displayName: message?.fromDisplayName,
        });
        const senderEntity = participantEntityMap.get(senderKey);
        return senderEntity?._id?.toString?.() ? [senderEntity._id.toString()] : [];
      }),
    ]);
    const senderLabels = uniqueStrings(
      windowMessages.map((message) => message?.fromDisplayName || message?.fromEmail),
    );
    const firstMessage = windowMessages[0];
    const lastMessage = windowMessages[windowMessages.length - 1];

    chunks.push({
      user: userId,
      tenantId,
      visibilityScope,
      source: 'teams',
      sourceRecordType: 'teams_chat',
      sourceRecordId: conversation.graphChatId,
      sourceParentRecordId: conversation.graphChatId,
      parentEntityId: conversationEntity._id.toString(),
      entityIds: windowEntityIds,
      chunkType: 'conversation_window',
      title: `${formatConversationTitle(conversation)} window ${windowIndex + 1}`.trim(),
      text: windowText,
      summary: `${windowMessages.length} messages from ${senderLabels.slice(0, 4).join(', ')}`.trim(),
      orderIndex: windowIndex,
      sourceTimestamp: lastMessage?.sentDateTime || firstMessage?.sentDateTime,
      metadata: {
        graphChatId: conversation.graphChatId,
        chatType: conversation.chatType || '',
        participants: uniqueStrings(
          (conversation?.participants || []).map(
            (participant) => participant?.displayName || participant?.email,
          ),
        ),
        senders: senderLabels,
        firstMessageAt: firstMessage?.sentDateTime || null,
        lastMessageAt: lastMessage?.sentDateTime || null,
        messageIds: windowMessages.map((message) => message.graphMessageId).filter(Boolean),
        includedMessageCount: windowMessages.length,
      },
    });
    conversationWindowChunkCount += 1;
  }

  await Promise.all([
    db.bulkUpsertEnterpriseMemoryRelationships(relationships),
    db.bulkUpsertEnterpriseMemoryChunks(chunks),
  ]);

  return {
    entityCount: 1 + participantEntityMap.size,
    relationshipCount: relationships.length,
    chunkCount: chunks.length,
    messageChunkCount,
    conversationWindowChunkCount,
    totalMessages: sortedMessages.length,
    chunkableMessageCount: messageChunkCount,
    skippedEmptyTextMessages,
    participantEntityCount: participantEntityMap.size,
    skippedReasonCounts,
    zeroChunkReasonCounts:
      messageChunkCount === 0 && conversationWindowChunkCount === 0 ? skippedReasonCounts : {},
    searchable: chunks.length > 0,
    chatType: conversation?.chatType || 'unknown',
    participantDegraded: Boolean(conversation?.participantDegraded),
  };
}

async function projectTeamsArchiveSyncToMemory({
  userId,
  tenantId,
  syncJobId,
  graphChatIds = [],
  visibilityScope = 'user',
  runStatus = 'success',
  deferredGraphChatIds = [],
}) {
  const deferredGraphChatIdSet = new Set(
    (Array.isArray(deferredGraphChatIds) ? deferredGraphChatIds : []).map((id) => String(id || '').trim()),
  );
  const requestedGraphChatIds = Array.isArray(graphChatIds) ? graphChatIds : [];
  const projectableGraphChatIds = requestedGraphChatIds.filter(
    (graphChatId) => !deferredGraphChatIdSet.has(String(graphChatId || '').trim()),
  );
  const excludedDeferredCount = requestedGraphChatIds.length - projectableGraphChatIds.length;
  const runIsPartial = runStatus === 'partial' || deferredGraphChatIdSet.size > 0;

  const projectionJob = await db.createEnterpriseMemoryJob({
    user: userId,
    tenantId,
    visibilityScope,
    source: 'teams',
    jobType: 'projection',
    status: 'running',
    sourceRecordType: 'teams_sync',
    sourceRecordId: syncJobId,
    stats: {
      requestedConversationCount: projectableGraphChatIds.length,
      projectedConversationCount: 0,
      entityCount: 0,
      relationshipCount: 0,
      chunkCount: 0,
      sourceRunStatus: runStatus,
      sourceRunPartial: runIsPartial,
      deferredConversationCount: deferredGraphChatIdSet.size,
      excludedDeferredConversationCount: excludedDeferredCount,
    },
    startedAt: new Date(),
    lastHeartbeatAt: new Date(),
  });

  if (runIsPartial) {
    logger.warn('[EnterpriseMemory] Projecting a partial Teams sync run; deferred conversations excluded', {
      userId,
      syncJobId,
      runStatus,
      deferredConversationCount: deferredGraphChatIdSet.size,
      excludedDeferredConversationCount: excludedDeferredCount,
      projectableConversationCount: projectableGraphChatIds.length,
    });
  }

  try {
    let projectedConversationCount = 0;
    let entityCount = 0;
    let relationshipCount = 0;
    let chunkCount = 0;
    let missingConversationCount = 0;
    let zeroMessageConversationCount = 0;
    let zeroChunkConversationCount = 0;
    let truncatedConversationCount = 0;
    let totalMessagesLoaded = 0;
    let totalChunkableMessages = 0;
    let totalSkippedEmptyTextMessages = 0;
    let messageChunkCount = 0;
    let conversationWindowChunkCount = 0;
    let participantDegradedConversationCount = 0;
    const zeroChunkReasonCounts = {};
    const searchableConversationCountsByChatType = {
      oneOnOne: 0,
      group: 0,
      meeting: 0,
      unknown: 0,
    };

    for (const graphChatId of projectableGraphChatIds) {
      const [conversation] = await db.findTeamsArchiveConversations({ user: userId, graphChatId }, { limit: 1 });
      if (!conversation) {
        missingConversationCount += 1;
        logger.warn('[EnterpriseMemory] Teams projection skipped missing archived conversation', {
          userId,
          syncJobId,
          graphChatId,
        });
        continue;
      }

      const messages = await db.findTeamsArchiveMessages(
        { user: userId, graphChatId },
        { limit: PROJECTION_MESSAGE_FETCH_LIMIT, sort: { sentDateTime: 1, createdAt: 1 } },
      );
      const archivedMessageCount = Number(conversation?.messageCount || 0);
      const truncatedByFetchCap =
        archivedMessageCount > messages.length ||
        messages.length >= PROJECTION_MESSAGE_FETCH_LIMIT;

      const result = await projectTeamsConversationToMemory({
        userId,
        tenantId,
        visibilityScope,
        conversation,
        messages,
      });

      projectedConversationCount += 1;
      entityCount += result.entityCount;
      relationshipCount += result.relationshipCount;
      chunkCount += result.chunkCount;
      messageChunkCount += result.messageChunkCount || 0;
      conversationWindowChunkCount += result.conversationWindowChunkCount || 0;
      totalMessagesLoaded += result.totalMessages || 0;
      totalChunkableMessages += result.chunkableMessageCount || 0;
      totalSkippedEmptyTextMessages += result.skippedEmptyTextMessages || 0;
      if (result.participantDegraded) {
        participantDegradedConversationCount += 1;
      }
      if (result.searchable) {
        searchableConversationCountsByChatType[result.chatType] =
          (searchableConversationCountsByChatType[result.chatType] || 0) + 1;
      }

      if ((result.totalMessages || 0) === 0) {
        zeroMessageConversationCount += 1;
      }

      if ((result.chunkCount || 0) === 0) {
        zeroChunkConversationCount += 1;
        for (const [reason, count] of Object.entries(result.zeroChunkReasonCounts || {})) {
          zeroChunkReasonCounts[reason] = (zeroChunkReasonCounts[reason] || 0) + Number(count || 0);
        }
        logger.info('[EnterpriseMemory] Teams projection produced zero searchable chunks for conversation', {
          userId,
          syncJobId,
          graphChatId,
          chatType: conversation?.chatType || '',
          topic: conversation?.topic || '',
          archivedMessageCount,
          loadedMessageCount: result.totalMessages || 0,
          skippedEmptyTextMessages: result.skippedEmptyTextMessages || 0,
          zeroChunkReasonCounts: result.zeroChunkReasonCounts || {},
          participantCount: Array.isArray(conversation?.participants)
            ? conversation.participants.length
            : 0,
        });
      }

      if (truncatedByFetchCap) {
        truncatedConversationCount += 1;
        logger.warn('[EnterpriseMemory] Teams projection truncated conversation by fetch cap', {
          userId,
          syncJobId,
          graphChatId,
          archivedMessageCount,
          loadedMessageCount: result.totalMessages || 0,
          projectionMessageFetchLimit: PROJECTION_MESSAGE_FETCH_LIMIT,
        });
      }
    }

    const projectionDiagnostics = {
      missingConversationCount,
      zeroMessageConversationCount,
      zeroChunkConversationCount,
      truncatedConversationCount,
      totalMessagesLoaded,
      totalChunkableMessages,
      totalSkippedEmptyTextMessages,
      projectionMessageFetchLimit: PROJECTION_MESSAGE_FETCH_LIMIT,
      zeroChunkReasonCounts,
      searchableConversationCountsByChatType,
      participantDegradedConversationCount,
      messageChunkCount,
      conversationWindowChunkCount,
      sourceRunStatus: runStatus,
      sourceRunPartial: runIsPartial,
      deferredConversationCount: deferredGraphChatIdSet.size,
      excludedDeferredConversationCount: excludedDeferredCount,
    };

    const updatedJob = await db.updateEnterpriseMemoryJob(
      projectionJob._id?.toString?.() || projectionJob.id,
      {
        status: 'success',
        completedAt: new Date(),
        lastHeartbeatAt: new Date(),
        stats: {
          requestedConversationCount: projectableGraphChatIds.length,
          projectedConversationCount,
          entityCount,
          relationshipCount,
          chunkCount,
          sourceRunStatus: runStatus,
          sourceRunPartial: runIsPartial,
          deferredConversationCount: deferredGraphChatIdSet.size,
          excludedDeferredConversationCount: excludedDeferredCount,
          projectionDiagnostics,
        },
      },
    );

    logger.info('[EnterpriseMemory] Teams projection completed', {
      userId,
      syncJobId,
      requestedConversationCount: projectableGraphChatIds.length,
      projectedConversationCount,
      entityCount,
      relationshipCount,
      chunkCount,
      sourceRunStatus: runStatus,
      sourceRunPartial: runIsPartial,
      projectionDiagnostics,
    });

    return {
      status: 'success',
      jobId: updatedJob?._id?.toString?.() || projectionJob._id?.toString?.() || projectionJob.id,
      projectedConversationCount,
      entityCount,
      relationshipCount,
      chunkCount,
      sourceRunStatus: runStatus,
      sourceRunPartial: runIsPartial,
      deferredConversationCount: deferredGraphChatIdSet.size,
      projectionDiagnostics,
    };
  } catch (error) {
    await db.updateEnterpriseMemoryJob(projectionJob._id?.toString?.() || projectionJob.id, {
      status: 'failure',
      completedAt: new Date(),
      lastHeartbeatAt: new Date(),
      errorMessage: error?.message || 'Teams enterprise memory projection failed',
    });

    logger.error('[EnterpriseMemory] Teams projection failed', {
      userId,
      syncJobId,
      error: error?.message || error,
    });

    throw error;
  }
}

module.exports = {
  projectTeamsArchiveSyncToMemory,
};
