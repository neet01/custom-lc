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

function buildPersonCanonicalKey({ teamId, slackUserId, botId, email, displayName }) {
  const workspaceKey = String(teamId || 'workspace').trim().toLowerCase();
  if (slackUserId) {
    return `slack_user:${workspaceKey}:${String(slackUserId).trim().toLowerCase()}`;
  }

  if (botId) {
    return `slack_bot:${workspaceKey}:${String(botId).trim().toLowerCase()}`;
  }

  if (email) {
    return `email:${String(email).trim().toLowerCase()}`;
  }

  return `name:${String(displayName || 'unknown').trim().toLowerCase()}`;
}

function buildParticipantMap(conversation, messages) {
  const participants = new Map();
  const teamId = conversation?.teamId || '';

  const addParticipant = (participant = {}) => {
    const displayName = String(participant.displayName || participant.realName || participant.username || '').trim();
    const email = String(participant.email || '').trim();
    const slackUserId = String(participant.slackUserId || participant.mentionedUserId || '').trim();
    const botId = String(participant.botId || '').trim();
    const canonicalKey = buildPersonCanonicalKey({
      teamId,
      slackUserId,
      botId,
      email,
      displayName,
    });

    const existing = participants.get(canonicalKey) || {
      canonicalKey,
      displayName: '',
      realName: '',
      username: '',
      email: '',
      slackUserId: '',
      botId: '',
      isBot: false,
      isAppUser: false,
      aliases: [],
    };

    existing.displayName = existing.displayName || displayName || email || slackUserId || botId || canonicalKey;
    existing.realName = existing.realName || String(participant.realName || '').trim();
    existing.username = existing.username || String(participant.username || '').trim();
    existing.email = existing.email || email;
    existing.slackUserId = existing.slackUserId || slackUserId;
    existing.botId = existing.botId || botId;
    existing.isBot = existing.isBot || Boolean(participant.isBot || botId);
    existing.isAppUser = existing.isAppUser || Boolean(participant.isAppUser);
    existing.aliases = uniqueStrings([
      ...(existing.aliases || []),
      displayName,
      participant.realName,
      participant.username,
      email,
    ]);
    participants.set(canonicalKey, existing);
  };

  for (const participant of conversation?.participants || []) {
    addParticipant(participant);
  }

  for (const message of messages) {
    addParticipant({
      displayName: message?.displayName,
      username: message?.username,
      slackUserId: message?.slackUserId,
      botId: message?.botId,
    });

    for (const mention of message?.mentions || []) {
      addParticipant({
        displayName: mention?.displayName,
        slackUserId: mention?.slackUserId,
      });
    }
  }

  return participants;
}

function formatConversationTitle(conversation) {
  if (conversation?.name) {
    return `#${conversation.name}`;
  }

  if (conversation?.topic) {
    return conversation.topic;
  }

  const participants = uniqueStrings(
    (conversation?.participants || []).map(
      (participant) => participant?.displayName || participant?.realName || participant?.username || participant?.email,
    ),
  );

  if (participants.length > 0) {
    return participants.slice(0, 4).join(', ');
  }

  return `Slack conversation ${conversation?.slackConversationId || ''}`.trim();
}

function getChunkableMessageText(message) {
  return String(message?.normalizedText || message?.text || '').trim();
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
    const lastTime = lastMessage?.sentAt ? new Date(lastMessage.sentAt).getTime() : 0;
    const currentTime = message?.sentAt ? new Date(message.sentAt).getTime() : 0;
    const timeGapExceeded =
      lastMessage &&
      currentTime > 0 &&
      lastTime > 0 &&
      currentTime - lastTime > CONVERSATION_WINDOW_MAX_TIME_GAP_MS;

    if (currentWindow.length >= CONVERSATION_WINDOW_MESSAGE_TARGET || timeGapExceeded) {
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
      const timestamp = message?.sentAt?.toISOString?.() || '';
      const sender = String(message?.displayName || message?.username || message?.slackUserId || 'Unknown').trim();
      const text = getChunkableMessageText(message);
      return `[${timestamp}] ${sender}: ${text}`.trim();
    })
    .join('\n');
}

async function projectSlackConversationToMemory({
  userId,
  tenantId,
  visibilityScope = 'user',
  conversation,
  messages,
}) {
  const workspaceScope = conversation?.teamId || conversation?.enterpriseId || 'workspace';
  const conversationEntity = await db.upsertEnterpriseMemoryEntity({
    user: userId,
    tenantId,
    visibilityScope,
    source: 'slack',
    entityType: 'conversation',
    canonicalKey: `slack_conversation:${workspaceScope}:${conversation.slackConversationId}`,
    displayName: formatConversationTitle(conversation),
    aliases: uniqueStrings([conversation.name, conversation.topic, conversation.purpose, conversation.slackConversationId]),
    summary: `Slack ${conversation.conversationType || 'conversation'}`,
    sourceRecordType: 'slack_conversation',
    sourceRecordId: conversation.slackConversationId,
    sourceUpdatedAt: conversation.sourceUpdatedAt || conversation.updatedAt || conversation.lastMessageAt,
    attributes: {
      teamId: conversation.teamId || '',
      enterpriseId: conversation.enterpriseId || '',
      conversationType: conversation.conversationType || '',
      topic: conversation.topic || '',
      purpose: conversation.purpose || '',
      isArchived: Boolean(conversation.isArchived),
      isSlackConnect: Boolean(conversation.isSlackConnect),
      messageCount: conversation.messageCount || 0,
      participantCount: Array.isArray(conversation.participants) ? conversation.participants.length : 0,
      lastMessageAt: conversation.lastMessageAt || null,
      lastMeaningfulMessageAt: conversation.lastMeaningfulMessageAt || null,
    },
  });

  const participants = buildParticipantMap(conversation, messages);
  const participantEntityMap = new Map();

  for (const participant of participants.values()) {
    const entity = await db.upsertEnterpriseMemoryEntity({
      user: userId,
      tenantId,
      visibilityScope,
      source: 'slack',
      entityType: 'person',
      canonicalKey: participant.canonicalKey,
      displayName: participant.displayName,
      aliases: participant.aliases,
      summary: participant.email || participant.username || participant.slackUserId || participant.botId || '',
      sourceRecordType: 'slack_participant',
      sourceRecordId: participant.slackUserId || participant.botId || participant.email || participant.displayName,
      sourceParentRecordId: conversation.slackConversationId,
      sourceUpdatedAt: conversation.sourceUpdatedAt || conversation.updatedAt || conversation.lastMessageAt,
      attributes: {
        slackUserId: participant.slackUserId || '',
        slackBotId: participant.botId || '',
        email: participant.email || '',
        username: participant.username || '',
        realName: participant.realName || '',
        isBot: Boolean(participant.isBot),
        isAppUser: Boolean(participant.isAppUser),
        teamId: conversation.teamId || '',
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
      source: 'slack',
      relationshipType: 'has_participant',
      fromEntityId: conversationEntity._id.toString(),
      toEntityId: entity._id.toString(),
      sourceRecordType: 'slack_conversation',
      sourceRecordId: conversation.slackConversationId,
      sourceUpdatedAt: conversation.sourceUpdatedAt || conversation.updatedAt || conversation.lastMessageAt,
      attributes: {
        conversationType: conversation.conversationType || '',
        teamId: conversation.teamId || '',
      },
    });
  }

  const sortedMessages = [...messages].sort((a, b) => {
    const aTime = a?.sentAt ? new Date(a.sentAt).getTime() : 0;
    const bTime = b?.sentAt ? new Date(b.sentAt).getTime() : 0;
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
      teamId: conversation?.teamId,
      slackUserId: message?.slackUserId,
      botId: message?.botId,
      displayName: message?.displayName || message?.username,
    });
    const senderEntity = participantEntityMap.get(senderKey);
    const mentionEntityIds = (message?.mentions || [])
      .map((mention) =>
        participantEntityMap.get(
          buildPersonCanonicalKey({
            teamId: conversation?.teamId,
            slackUserId: mention?.slackUserId,
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
      source: 'slack',
      sourceRecordType: 'slack_message',
      sourceRecordId: message.slackMessageTs,
      sourceParentRecordId: conversation.slackConversationId,
      parentEntityId: conversationEntity._id.toString(),
      entityIds: uniqueStrings([
        conversationEntity._id.toString(),
        senderEntity?._id?.toString(),
        ...mentionEntityIds,
      ]),
      chunkType: 'message',
      title: `${message?.displayName || message?.username || 'Unknown'} @ ${message?.sentAt?.toISOString?.() || ''}`.trim(),
      text,
      summary: text.slice(0, 512),
      orderIndex: index,
      sourceTimestamp: message?.sentAt,
      metadata: {
        slackConversationId: conversation.slackConversationId,
        slackMessageTs: message.slackMessageTs,
        threadTs: message.threadTs || '',
        slackUserId: message.slackUserId || '',
        botId: message.botId || '',
        displayName: message.displayName || '',
        username: message.username || '',
        subtype: message.subtype || '',
        replyCount: Number(message.replyCount || 0),
        reactionCount: Array.isArray(message.reactions) ? message.reactions.length : 0,
        attachmentCount: Array.isArray(message.attachments) ? message.attachments.length : 0,
        fileCount: Array.isArray(message.files) ? message.files.length : 0,
        conversationType: conversation.conversationType || '',
        teamId: conversation.teamId || '',
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
          teamId: conversation?.teamId,
          slackUserId: message?.slackUserId,
          botId: message?.botId,
          displayName: message?.displayName || message?.username,
        });
        const senderEntity = participantEntityMap.get(senderKey);
        return senderEntity?._id?.toString?.() ? [senderEntity._id.toString()] : [];
      }),
    ]);
    const senderLabels = uniqueStrings(
      windowMessages.map((message) => message?.displayName || message?.username || message?.slackUserId),
    );
    const firstMessage = windowMessages[0];
    const lastMessage = windowMessages[windowMessages.length - 1];

    chunks.push({
      user: userId,
      tenantId,
      visibilityScope,
      source: 'slack',
      sourceRecordType: 'slack_conversation',
      sourceRecordId: conversation.slackConversationId,
      sourceParentRecordId: conversation.slackConversationId,
      parentEntityId: conversationEntity._id.toString(),
      entityIds: windowEntityIds,
      chunkType: 'conversation_window',
      title: `${formatConversationTitle(conversation)} window ${windowIndex + 1}`.trim(),
      text: windowText,
      summary: `${windowMessages.length} messages from ${senderLabels.slice(0, 4).join(', ')}`.trim(),
      orderIndex: windowIndex,
      sourceTimestamp: lastMessage?.sentAt || firstMessage?.sentAt,
      metadata: {
        slackConversationId: conversation.slackConversationId,
        conversationType: conversation.conversationType || '',
        senders: senderLabels,
        participants: uniqueStrings(
          (conversation?.participants || []).map(
            (participant) =>
              participant?.displayName || participant?.realName || participant?.username || participant?.email,
          ),
        ),
        firstMessageAt: firstMessage?.sentAt || null,
        lastMessageAt: lastMessage?.sentAt || null,
        messageIds: windowMessages.map((message) => message.slackMessageTs).filter(Boolean),
        includedMessageCount: windowMessages.length,
        teamId: conversation.teamId || '',
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
    conversationType: conversation?.conversationType || 'unknown',
  };
}

async function projectSlackArchiveSyncToMemory({
  userId,
  tenantId,
  syncJobId,
  slackConversationIds = [],
  visibilityScope = 'user',
  runStatus = 'success',
}) {
  const requestedConversationIds = Array.isArray(slackConversationIds)
    ? slackConversationIds.map((id) => String(id || '').trim()).filter(Boolean)
    : [];

  const projectionJob = await db.createEnterpriseMemoryJob({
    user: userId,
    tenantId,
    visibilityScope,
    source: 'slack',
    jobType: 'projection',
    status: 'running',
    sourceRecordType: 'slack_sync',
    sourceRecordId: syncJobId,
    stats: {
      requestedConversationCount: requestedConversationIds.length,
      projectedConversationCount: 0,
      entityCount: 0,
      relationshipCount: 0,
      chunkCount: 0,
      sourceRunStatus: runStatus,
    },
    startedAt: new Date(),
    lastHeartbeatAt: new Date(),
  });

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
    const zeroChunkReasonCounts = {};
    const searchableConversationCountsByType = {
      public_channel: 0,
      private_channel: 0,
      im: 0,
      mpim: 0,
      unknown: 0,
    };

    for (const slackConversationId of requestedConversationIds) {
      const [conversation] = await db.findSlackArchiveConversations(
        { user: userId, slackConversationId },
        { limit: 1 },
      );
      if (!conversation) {
        missingConversationCount += 1;
        logger.warn('[EnterpriseMemory] Slack projection skipped missing archived conversation', {
          userId,
          syncJobId,
          slackConversationId,
        });
        continue;
      }

      const messages = await db.findSlackArchiveMessages(
        { user: userId, slackConversationId },
        { limit: PROJECTION_MESSAGE_FETCH_LIMIT, sort: { sentAt: 1, createdAt: 1 } },
      );
      const archivedMessageCount = Number(conversation?.messageCount || 0);
      const truncatedByFetchCap =
        archivedMessageCount > messages.length || messages.length >= PROJECTION_MESSAGE_FETCH_LIMIT;

      const result = await projectSlackConversationToMemory({
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

      if (result.searchable) {
        searchableConversationCountsByType[result.conversationType] =
          (searchableConversationCountsByType[result.conversationType] || 0) + 1;
      }

      if ((result.totalMessages || 0) === 0) {
        zeroMessageConversationCount += 1;
      }

      if ((result.chunkCount || 0) === 0) {
        zeroChunkConversationCount += 1;
        for (const [reason, count] of Object.entries(result.zeroChunkReasonCounts || {})) {
          zeroChunkReasonCounts[reason] = (zeroChunkReasonCounts[reason] || 0) + Number(count || 0);
        }
        logger.info('[EnterpriseMemory] Slack projection produced zero searchable chunks for conversation', {
          userId,
          syncJobId,
          slackConversationId,
          conversationType: conversation?.conversationType || '',
          name: conversation?.name || '',
          archivedMessageCount,
          loadedMessageCount: result.totalMessages || 0,
          skippedEmptyTextMessages: result.skippedEmptyTextMessages || 0,
          zeroChunkReasonCounts: result.zeroChunkReasonCounts || {},
          participantCount: Array.isArray(conversation?.participants) ? conversation.participants.length : 0,
        });
      }

      if (truncatedByFetchCap) {
        truncatedConversationCount += 1;
        logger.warn('[EnterpriseMemory] Slack projection truncated conversation by fetch cap', {
          userId,
          syncJobId,
          slackConversationId,
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
      searchableConversationCountsByType,
      messageChunkCount,
      conversationWindowChunkCount,
      sourceRunStatus: runStatus,
    };

    const updatedJob = await db.updateEnterpriseMemoryJob(
      projectionJob._id?.toString?.() || projectionJob.id,
      {
        status: 'success',
        completedAt: new Date(),
        lastHeartbeatAt: new Date(),
        stats: {
          requestedConversationCount: requestedConversationIds.length,
          projectedConversationCount,
          entityCount,
          relationshipCount,
          chunkCount,
          sourceRunStatus: runStatus,
          projectionDiagnostics,
        },
      },
    );

    logger.info('[EnterpriseMemory] Slack projection completed', {
      userId,
      syncJobId,
      requestedConversationCount: requestedConversationIds.length,
      projectedConversationCount,
      entityCount,
      relationshipCount,
      chunkCount,
      sourceRunStatus: runStatus,
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
      projectionDiagnostics,
    };
  } catch (error) {
    await db.updateEnterpriseMemoryJob(projectionJob._id?.toString?.() || projectionJob.id, {
      status: 'failure',
      completedAt: new Date(),
      lastHeartbeatAt: new Date(),
      errorMessage: error?.message || 'Slack enterprise memory projection failed',
    });

    logger.error('[EnterpriseMemory] Slack projection failed', {
      userId,
      syncJobId,
      error: error?.message || error,
    });

    throw error;
  }
}

module.exports = {
  projectSlackArchiveSyncToMemory,
};
