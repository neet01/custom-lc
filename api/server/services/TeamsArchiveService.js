const { isEnabled } = require('@librechat/api');
const { logger } = require('@librechat/data-schemas');
const { getGraphApiToken } = require('~/server/services/GraphTokenService');
const db = require('~/models');

const DEFAULT_GRAPH_BASE_URL = 'https://graph.microsoft.us/v1.0';
const DEFAULT_SCOPES = 'https://graph.microsoft.us/.default';
const DEFAULT_CHAT_LIMIT = 50;
const DEFAULT_MESSAGES_PER_CHAT = 250;
const DEFAULT_SEARCH_LIMIT = 25;

class TeamsArchiveServiceError extends Error {
  constructor(message, status = 500, details) {
    super(message);
    this.name = 'TeamsArchiveServiceError';
    this.status = status;
    this.details = details;
  }
}

function isTeamsArchiveEnabled() {
  return isEnabled(process.env.TEAMS_ARCHIVE_ENABLED);
}

function normalizeGraphBaseUrl(baseUrl = DEFAULT_GRAPH_BASE_URL) {
  const trimmed = String(baseUrl || DEFAULT_GRAPH_BASE_URL)
    .trim()
    .replace(/\/+$/, '');
  if (/\/(v1\.0|beta)$/i.test(trimmed)) {
    return trimmed;
  }
  return `${trimmed}/v1.0`;
}

function getTeamsArchiveConfig() {
  return {
    enabled: isTeamsArchiveEnabled(),
    graphBaseUrl: normalizeGraphBaseUrl(
      process.env.TEAMS_ARCHIVE_GRAPH_BASE_URL || DEFAULT_GRAPH_BASE_URL,
    ),
    scopes: process.env.TEAMS_ARCHIVE_GRAPH_SCOPES || DEFAULT_SCOPES,
    defaultChatLimit: Number(process.env.TEAMS_ARCHIVE_MAX_SYNC_CHATS) || DEFAULT_CHAT_LIMIT,
    defaultMessagesPerChat:
      Number(process.env.TEAMS_ARCHIVE_MAX_MESSAGES_PER_CHAT) || DEFAULT_MESSAGES_PER_CHAT,
    defaultSearchLimit: Number(process.env.TEAMS_ARCHIVE_SEARCH_LIMIT) || DEFAULT_SEARCH_LIMIT,
  };
}

function assertEnabled() {
  if (!isTeamsArchiveEnabled()) {
    throw new TeamsArchiveServiceError('Teams archive is not enabled', 403);
  }
}

function assertDelegatedUser(user) {
  if (!user?.openidId || user?.provider !== 'openid') {
    throw new TeamsArchiveServiceError(
      'Teams archive access requires Entra ID authentication',
      403,
    );
  }

  if (!isEnabled(process.env.OPENID_REUSE_TOKENS)) {
    throw new TeamsArchiveServiceError('Teams archive requires OPENID_REUSE_TOKENS=true', 403);
  }

  if (!user?.federatedTokens?.access_token) {
    throw new TeamsArchiveServiceError(
      'No delegated OpenID token is available for Microsoft Graph',
      401,
    );
  }
}

async function getDelegatedGraphToken(user, scopes = getTeamsArchiveConfig().scopes) {
  assertEnabled();
  assertDelegatedUser(user);
  const tokenResponse = await getGraphApiToken(user, user.federatedTokens.access_token, scopes);
  if (!tokenResponse?.access_token) {
    throw new TeamsArchiveServiceError(
      'Microsoft Graph token exchange did not return an access token',
      502,
    );
  }
  return tokenResponse.access_token;
}

function buildGraphUrl(pathname, query) {
  if (/^https?:\/\//i.test(pathname)) {
    const url = new URL(pathname);
    if (query) {
      for (const [key, value] of Object.entries(query)) {
        if (value !== undefined && value !== null && value !== '') {
          url.searchParams.set(key, String(value));
        }
      }
    }
    return url;
  }

  const { graphBaseUrl } = getTeamsArchiveConfig();
  const base = graphBaseUrl.endsWith('/') ? graphBaseUrl : `${graphBaseUrl}/`;
  const url = new URL(pathname.replace(/^\//, ''), base);

  if (query) {
    for (const [key, value] of Object.entries(query)) {
      if (value !== undefined && value !== null && value !== '') {
        url.searchParams.set(key, String(value));
      }
    }
  }

  return url;
}

async function parseGraphError(response) {
  try {
    const payload = await response.json();
    return payload?.error?.message || payload?.message || response.statusText;
  } catch {
    return response.statusText;
  }
}

async function graphRequest(user, pathname, options = {}) {
  const token = await getDelegatedGraphToken(user, options.scopes);
  const url = buildGraphUrl(pathname, options.query);
  const headers = {
    Authorization: `Bearer ${token}`,
    Accept: 'application/json',
    ...options.headers,
  };

  if (options.body !== undefined) {
    headers['Content-Type'] = 'application/json';
  }

  const response = await fetch(url, {
    method: options.method || 'GET',
    headers,
    body: options.body !== undefined ? JSON.stringify(options.body) : undefined,
  });

  if (!response.ok) {
    const graphMessage = await parseGraphError(response);
    if (!options.suppressErrorLog) {
      logger.warn('[TeamsArchiveService] Microsoft Graph request failed', {
        status: response.status,
        path: pathname,
        graphMessage,
      });
    }
    throw new TeamsArchiveServiceError(
      'Microsoft Graph request failed',
      response.status,
      graphMessage,
    );
  }

  if (response.status === 204) {
    return null;
  }

  return response.json();
}

function decodeHtmlEntities(value) {
  return String(value || '')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/&#x27;/gi, "'");
}

function normalizeHtmlText(content, contentType = 'text') {
  const raw = String(content || '');
  if (String(contentType || '').toLowerCase() !== 'html') {
    return raw.trim();
  }

  return decodeHtmlEntities(
    raw
      .replace(/<style[\s\S]*?<\/style>/gi, ' ')
      .replace(/<script[\s\S]*?<\/script>/gi, ' ')
      .replace(/<\/(p|div|tr|li|h[1-6]|table|section|article|br)>/gi, '\n')
      .replace(/<[^>]+>/g, ' ')
      .replace(/[ \t]+\n/g, '\n')
      .replace(/\n{3,}/g, '\n\n')
      .replace(/[ \t]{2,}/g, ' '),
  ).trim();
}

function toDate(value) {
  if (!value) {
    return undefined;
  }
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? undefined : date;
}

function toArray(value) {
  return Array.isArray(value) ? value : [];
}

function uniqueParticipants(participants) {
  const seen = new Set();
  return participants.filter((participant) => {
    const key = `${participant.userId || ''}:${participant.email || ''}:${participant.displayName || ''}`;
    if (!key || seen.has(key)) {
      return false;
    }
    seen.add(key);
    return true;
  });
}

function normalizeConversation(chat, members = []) {
  const participants = uniqueParticipants(
    members
      .map((member) => ({
        displayName: member?.displayName || member?.email || '',
        email: member?.email || member?.userId || '',
        userId: member?.userId,
      }))
      .filter((participant) => participant.displayName || participant.email || participant.userId),
  );

  return {
    graphChatId: chat.id,
    chatType: chat.chatType,
    topic: chat.topic || chat.subject || '',
    webUrl: chat.webUrl || '',
    participants,
    sourceUpdatedAt: toDate(chat.lastUpdatedDateTime),
  };
}

function normalizeMessage(chatId, message) {
  const fromUser = message?.from?.user || message?.from?.application || {};
  const from = {
    displayName: message?.from?.user?.displayName || message?.from?.application?.displayName || '',
    userId: fromUser?.id || '',
  };
  const bodyContentType = String(message?.body?.contentType || 'html').toLowerCase();
  const bodyContent = String(message?.body?.content || '');
  const bodyText = normalizeHtmlText(bodyContent, bodyContentType);
  const bodyPreview = bodyText.slice(0, 500);

  return {
    graphChatId: chatId,
    graphMessageId: message.id,
    replyToId: message.replyToId || '',
    fromDisplayName: from.displayName,
    fromEmail: fromUser?.email || fromUser?.userIdentityType || '',
    fromUserId: from.userId,
    subject: message.subject || '',
    summary: message.summary || '',
    importance: message.importance || '',
    messageType: message.messageType || '',
    bodyContentType,
    bodyPreview,
    bodyContent,
    bodyText,
    attachments: toArray(message.attachments).map((attachment) => ({
      id: attachment?.id,
      name: attachment?.name,
      contentType: attachment?.contentType,
      contentUrl: attachment?.contentUrl,
    })),
    mentions: toArray(message.mentions).map((mention) => ({
      id: mention?.id,
      displayName: mention?.mentionText || mention?.mentioned?.user?.displayName,
      mentionedUserId: mention?.mentioned?.user?.id,
    })),
    webUrl: message.webUrl || '',
    sentDateTime: toDate(message.createdDateTime),
    lastModifiedDateTime: toDate(message.lastModifiedDateTime),
    deletedDateTime: toDate(message.deletedDateTime),
    etag: message.etag || message['@odata.etag'] || '',
  };
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

async function listChatsPage(user, { top = DEFAULT_CHAT_LIMIT, nextLink } = {}) {
  if (nextLink) {
    return graphRequest(user, nextLink);
  }

  return graphRequest(user, '/me/chats', {
    query: {
      $top: top,
      $orderby: 'lastUpdatedDateTime desc',
    },
  });
}

async function listChatMembers(user, chatId) {
  try {
    const response = await graphRequest(user, `/chats/${encodeURIComponent(chatId)}/members`, {
      query: { $top: 50 },
      suppressErrorLog: true,
    });

    return toArray(response?.value).map((member) => ({
      displayName: member?.displayName || member?.email || '',
      email: member?.email || '',
      userId: member?.userId || member?.id || '',
    }));
  } catch (error) {
    logger.warn('[TeamsArchiveService] Failed to list chat members', {
      chatId,
      error: error?.message,
    });
    return [];
  }
}

async function listChatMessages(user, chatId, { top = DEFAULT_MESSAGES_PER_CHAT } = {}) {
  const messages = [];
  let nextLink = null;

  do {
    const response = await graphRequest(
      user,
      nextLink || `/chats/${encodeURIComponent(chatId)}/messages`,
      nextLink
        ? {}
        : {
            query: {
              $top: Math.min(top - messages.length, 50),
            },
          },
    );

    messages.push(...toArray(response?.value));
    nextLink = response?.['@odata.nextLink'] && messages.length < top ? response['@odata.nextLink'] : null;
  } while (nextLink && messages.length < top);

  return messages.slice(0, top);
}

async function getStatus(user) {
  const config = getTeamsArchiveConfig();
  const userId = user?.id || user?._id?.toString();
  const [conversationCount, messageCount, latestSync] = await Promise.all([
    userId ? db.countTeamsArchiveConversations({ user: userId }) : 0,
    userId ? db.countTeamsArchiveMessages({ user: userId }) : 0,
    userId ? db.findLatestTeamsArchiveSyncJob({ user: userId }) : null,
  ]);

  return {
    enabled: config.enabled,
    graphBaseUrl: config.graphBaseUrl,
    graphScopes: config.scopes,
    syncModes: ['chats'],
    channelSyncSupported: false,
    conversationCount,
    messageCount,
    latestSync: latestSync
      ? {
          id: latestSync._id?.toString?.() || latestSync.id,
          status: latestSync.status,
          mode: latestSync.mode,
          conversationCount: latestSync.conversationCount || 0,
          messageCount: latestSync.messageCount || 0,
          startedAt: latestSync.startedAt,
          completedAt: latestSync.completedAt,
          errorMessage: latestSync.errorMessage,
        }
      : null,
  };
}

async function syncUserArchive(user, options = {}) {
  assertEnabled();
  assertDelegatedUser(user);

  const userId = user?.id || user?._id?.toString();
  if (!userId) {
    throw new TeamsArchiveServiceError('User id is required for Teams archive sync', 400);
  }

  const config = getTeamsArchiveConfig();
  const chatLimit = clampInteger(options.chatLimit, config.defaultChatLimit, { max: 250 });
  const messagesPerChat = clampInteger(options.messagesPerChat, config.defaultMessagesPerChat, {
    max: 1000,
  });
  const mode = options.mode === 'chats' ? 'chats' : 'chats';

  const syncJob = await db.createTeamsArchiveSyncJob({
    user: userId,
    status: 'running',
    mode,
    conversationCount: 0,
    messageCount: 0,
    startedAt: new Date(),
  });

  try {
    const syncedConversations = [];
    let nextLink = null;
    let processedChats = 0;
    let persistedMessages = 0;

    while (processedChats < chatLimit) {
      const response = await listChatsPage(user, {
        top: Math.min(chatLimit - processedChats, 50),
        nextLink,
      });
      const chats = toArray(response?.value);
      if (chats.length === 0) {
        break;
      }

      for (const chat of chats) {
        if (processedChats >= chatLimit) {
          break;
        }

        const members = await listChatMembers(user, chat.id);
        const normalizedConversation = normalizeConversation(chat, members);
        const messages = await listChatMessages(user, chat.id, { top: messagesPerChat });
        const normalizedMessages = messages.map((message) => ({
          user: userId,
          ...normalizeMessage(chat.id, message),
        }));

        if (normalizedMessages.length > 0) {
          persistedMessages += await db.bulkUpsertTeamsArchiveMessages(normalizedMessages);
        }

        const sortedMessages = [...normalizedMessages]
          .filter((message) => message.sentDateTime instanceof Date)
          .sort((a, b) => b.sentDateTime.getTime() - a.sentDateTime.getTime());

        const conversationRecord = await db.upsertTeamsArchiveConversation({
          user: userId,
          ...normalizedConversation,
          lastMessageAt: sortedMessages[0]?.sentDateTime,
          lastSyncedAt: new Date(),
          messageCount: normalizedMessages.length,
        });

        syncedConversations.push({
          id: conversationRecord._id?.toString?.() || conversationRecord.id,
          graphChatId: conversationRecord.graphChatId,
          topic: conversationRecord.topic || '',
          chatType: conversationRecord.chatType || '',
          messageCount: conversationRecord.messageCount || 0,
          lastMessageAt: conversationRecord.lastMessageAt,
        });

        processedChats += 1;
      }

      nextLink =
        response?.['@odata.nextLink'] && processedChats < chatLimit ? response['@odata.nextLink'] : null;

      if (!nextLink) {
        break;
      }
    }

    const updatedJob = await db.updateTeamsArchiveSyncJob(syncJob._id?.toString?.() || syncJob.id, {
      status: 'success',
      conversationCount: syncedConversations.length,
      messageCount: persistedMessages,
      completedAt: new Date(),
    });

    return {
      syncJob: updatedJob || syncJob,
      mode,
      conversationCount: syncedConversations.length,
      messageCount: persistedMessages,
      conversations: syncedConversations,
    };
  } catch (error) {
    await db.updateTeamsArchiveSyncJob(syncJob._id?.toString?.() || syncJob.id, {
      status: 'failure',
      errorMessage: error?.message || 'Teams archive sync failed',
      completedAt: new Date(),
    });
    throw error;
  }
}

async function listConversations(user, options = {}) {
  assertEnabled();
  const userId = user?.id || user?._id?.toString();
  const limit = clampInteger(options.limit, 50, { max: 200 });
  const offset = clampInteger(options.offset, 0, { min: 0, max: 100000 });

  const conversations = await db.findTeamsArchiveConversations(
    { user: userId },
    { limit, offset, sort: { lastMessageAt: -1, updatedAt: -1 } },
  );

  return {
    conversations: conversations.map((conversation) => ({
      id: conversation._id?.toString?.() || conversation.id,
      graphChatId: conversation.graphChatId,
      chatType: conversation.chatType || '',
      topic: conversation.topic || '',
      participants: conversation.participants || [],
      webUrl: conversation.webUrl || '',
      lastMessageAt: conversation.lastMessageAt,
      lastSyncedAt: conversation.lastSyncedAt,
      sourceUpdatedAt: conversation.sourceUpdatedAt,
      messageCount: conversation.messageCount || 0,
    })),
  };
}

async function listConversationMessages(user, chatId, options = {}) {
  assertEnabled();
  const userId = user?.id || user?._id?.toString();
  if (!chatId) {
    throw new TeamsArchiveServiceError('Chat id is required', 400);
  }

  const limit = clampInteger(options.limit, 100, { max: 500 });
  const offset = clampInteger(options.offset, 0, { min: 0, max: 100000 });

  const messages = await db.findTeamsArchiveMessages(
    { user: userId, graphChatId: chatId },
    { limit, offset, sort: { sentDateTime: 1, createdAt: 1 } },
  );

  return {
    chatId,
    messages: messages.map((message) => ({
      id: message._id?.toString?.() || message.id,
      graphMessageId: message.graphMessageId,
      replyToId: message.replyToId || '',
      fromDisplayName: message.fromDisplayName || '',
      fromEmail: message.fromEmail || '',
      subject: message.subject || '',
      summary: message.summary || '',
      importance: message.importance || '',
      bodyPreview: message.bodyPreview || '',
      bodyText: message.bodyText || '',
      bodyContentType: message.bodyContentType || 'html',
      bodyContent: message.bodyContent || '',
      webUrl: message.webUrl || '',
      sentDateTime: message.sentDateTime,
      lastModifiedDateTime: message.lastModifiedDateTime,
      attachments: message.attachments || [],
      mentions: message.mentions || [],
    })),
  };
}

async function searchMessages(user, options = {}) {
  assertEnabled();
  const userId = user?.id || user?._id?.toString();
  const query = String(options.query || '').trim();
  if (!query) {
    throw new TeamsArchiveServiceError('Search query is required', 400);
  }

  const limit = clampInteger(options.limit, getTeamsArchiveConfig().defaultSearchLimit, {
    max: 100,
  });
  const offset = clampInteger(options.offset, 0, { min: 0, max: 100000 });
  const chatId = options.chatId ? String(options.chatId) : undefined;
  const regex = new RegExp(escapeRegex(query), 'i');
  const filter = {
    user: userId,
    ...(chatId ? { graphChatId: chatId } : {}),
    $or: [
      { bodyText: regex },
      { bodyPreview: regex },
      { summary: regex },
      { subject: regex },
      { fromDisplayName: regex },
      { fromEmail: regex },
    ],
  };

  const messages = await db.findTeamsArchiveMessages(filter, {
    limit,
    offset,
    sort: { sentDateTime: -1, createdAt: -1 },
  });

  const conversationIds = [...new Set(messages.map((message) => message.graphChatId).filter(Boolean))];
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
    query,
    results: messages.map((message) => {
      const conversation = conversationMap.get(message.graphChatId);
      return {
        id: message._id?.toString?.() || message.id,
        graphMessageId: message.graphMessageId,
        graphChatId: message.graphChatId,
        topic: conversation?.topic || '',
        chatType: conversation?.chatType || '',
        fromDisplayName: message.fromDisplayName || '',
        fromEmail: message.fromEmail || '',
        bodyPreview: message.bodyPreview || '',
        bodyText: message.bodyText || '',
        sentDateTime: message.sentDateTime,
        webUrl: message.webUrl || '',
      };
    }),
  };
}

module.exports = {
  TeamsArchiveServiceError,
  getStatus,
  syncUserArchive,
  listConversations,
  listConversationMessages,
  searchMessages,
};
