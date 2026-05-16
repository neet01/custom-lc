const { logger } = require('@librechat/data-schemas');
const db = require('~/models');

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

  for (const [index, message] of sortedMessages.entries()) {
    const text = String(message?.bodyText || message?.summary || message?.subject || '').trim();
    if (!text) {
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
        webUrl: message.webUrl || '',
        importance: message.importance || '',
        messageType: message.messageType || '',
        attachmentCount: Array.isArray(message.attachments) ? message.attachments.length : 0,
        mentionCount: Array.isArray(message.mentions) ? message.mentions.length : 0,
      },
    });
  }

  await Promise.all([
    db.bulkUpsertEnterpriseMemoryRelationships(relationships),
    db.bulkUpsertEnterpriseMemoryChunks(chunks),
  ]);

  return {
    entityCount: 1 + participantEntityMap.size,
    relationshipCount: relationships.length,
    chunkCount: chunks.length,
  };
}

async function projectTeamsArchiveSyncToMemory({
  userId,
  tenantId,
  syncJobId,
  graphChatIds = [],
  visibilityScope = 'user',
}) {
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
      requestedConversationCount: graphChatIds.length,
      projectedConversationCount: 0,
      entityCount: 0,
      relationshipCount: 0,
      chunkCount: 0,
    },
    startedAt: new Date(),
    lastHeartbeatAt: new Date(),
  });

  try {
    let projectedConversationCount = 0;
    let entityCount = 0;
    let relationshipCount = 0;
    let chunkCount = 0;

    for (const graphChatId of graphChatIds) {
      const [conversation] = await db.findTeamsArchiveConversations({ user: userId, graphChatId }, { limit: 1 });
      if (!conversation) {
        continue;
      }

      const messages = await db.findTeamsArchiveMessages(
        { user: userId, graphChatId },
        { limit: 5000, sort: { sentDateTime: 1, createdAt: 1 } },
      );

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
    }

    const updatedJob = await db.updateEnterpriseMemoryJob(
      projectionJob._id?.toString?.() || projectionJob.id,
      {
        status: 'success',
        completedAt: new Date(),
        lastHeartbeatAt: new Date(),
        stats: {
          requestedConversationCount: graphChatIds.length,
          projectedConversationCount,
          entityCount,
          relationshipCount,
          chunkCount,
        },
      },
    );

    return {
      status: 'success',
      jobId: updatedJob?._id?.toString?.() || projectionJob._id?.toString?.() || projectionJob.id,
      projectedConversationCount,
      entityCount,
      relationshipCount,
      chunkCount,
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
