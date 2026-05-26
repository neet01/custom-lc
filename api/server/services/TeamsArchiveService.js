const { isEnabled } = require('@librechat/api');
const { randomUUID } = require('crypto');
const { logger, runAsSystem } = require('@librechat/data-schemas');
const { getGraphApiToken } = require('~/server/services/GraphTokenService');
const db = require('~/models');

let projectTeamsArchiveSyncToMemory = null;
let searchTeamsMemoryChunks = null;

try {
  ({ projectTeamsArchiveSyncToMemory } = require('~/server/services/EnterpriseMemory/teamsProjection'));
} catch (error) {
  if (error?.code !== 'MODULE_NOT_FOUND') {
    throw error;
  }
}

try {
  ({ searchTeamsMemoryChunks } = require('~/server/services/EnterpriseMemory/retrieval'));
} catch (error) {
  if (error?.code !== 'MODULE_NOT_FOUND') {
    throw error;
  }
}

const DEFAULT_GRAPH_BASE_URL = 'https://graph.microsoft.us/v1.0';
const DEFAULT_SCOPES = 'https://graph.microsoft.us/.default';
const DEFAULT_CHAT_LIMIT = 50;
const DEFAULT_MESSAGES_PER_CHAT = 250;
const DEFAULT_SEARCH_LIMIT = 25;
const DEFAULT_SYNC_STALE_MINUTES = 45;
const DEFAULT_MAX_CONCURRENT_SYNCS = 3;
const DEFAULT_DISCOVERY_REFRESH_HOURS = 12;
const DEFAULT_DISCOVERY_CONCURRENCY = 4;
const DEFAULT_MEMBER_ENRICHMENT_MODE = 'adaptive';
const DEFAULT_MEMBER_ENRICHMENT_FAILURE_THRESHOLD = 8;
const DEFAULT_CONVERSATION_DOSSIER_MAX_MESSAGES = 20000;
const HEARTBEAT_MIN_INTERVAL_MS = 30 * 1000;
const HEARTBEAT_CHAT_INTERVAL = 5;

class TeamsArchiveServiceError extends Error {
  constructor(message, status = 500, details) {
    super(message);
    this.name = 'TeamsArchiveServiceError';
    this.status = status;
    this.details = details;
  }
}

class TeamsArchiveSyncCancelledError extends Error {
  constructor(message = 'Teams archive sync cancelled by user') {
    super(message);
    this.name = 'TeamsArchiveSyncCancelledError';
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
  const parsedMaxConcurrentSyncs = Number(process.env.TEAMS_ARCHIVE_MAX_CONCURRENT_SYNCS);
  const parsedDiscoveryRefreshHours = Number(process.env.TEAMS_ARCHIVE_DISCOVERY_REFRESH_HOURS);
  const parsedDiscoveryConcurrency = Number(process.env.TEAMS_ARCHIVE_DISCOVERY_CONCURRENCY);
  const parsedMemberLookupFailureThreshold = Number(
    process.env.TEAMS_ARCHIVE_MEMBER_ENRICHMENT_FAILURE_THRESHOLD,
  );
  const memberEnrichmentMode = String(
    process.env.TEAMS_ARCHIVE_MEMBER_ENRICHMENT_MODE || DEFAULT_MEMBER_ENRICHMENT_MODE,
  )
    .trim()
    .toLowerCase();

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
    syncStaleMinutes:
      Number(process.env.TEAMS_ARCHIVE_SYNC_STALE_MINUTES) || DEFAULT_SYNC_STALE_MINUTES,
    discoveryRefreshHours:
      Number.isFinite(parsedDiscoveryRefreshHours) && parsedDiscoveryRefreshHours > 0
        ? parsedDiscoveryRefreshHours
        : DEFAULT_DISCOVERY_REFRESH_HOURS,
    discoveryConcurrency:
      Number.isFinite(parsedDiscoveryConcurrency) && parsedDiscoveryConcurrency > 0
        ? Math.min(Math.floor(parsedDiscoveryConcurrency), 16)
        : DEFAULT_DISCOVERY_CONCURRENCY,
    memberEnrichmentMode: ['all', 'non_meeting', 'disabled', 'adaptive'].includes(
      memberEnrichmentMode,
    )
      ? memberEnrichmentMode
      : DEFAULT_MEMBER_ENRICHMENT_MODE,
    memberEnrichmentFailureThreshold:
      Number.isFinite(parsedMemberLookupFailureThreshold) && parsedMemberLookupFailureThreshold > 0
        ? Math.floor(parsedMemberLookupFailureThreshold)
        : DEFAULT_MEMBER_ENRICHMENT_FAILURE_THRESHOLD,
    maxConcurrentSyncs:
      Number.isFinite(parsedMaxConcurrentSyncs) && parsedMaxConcurrentSyncs >= 0
        ? Math.floor(parsedMaxConcurrentSyncs)
        : DEFAULT_MAX_CONCURRENT_SYNCS,
  };
}

function getLeaseDurationMs() {
  return getTeamsArchiveConfig().syncStaleMinutes * 60 * 1000;
}

function getUserLeaseKey(userId) {
  return `user:${userId}`;
}

function getSlotLeaseKey(slotNumber) {
  return `slot:${slotNumber}`;
}

function getLeaseExpiryDate() {
  return new Date(Date.now() + getLeaseDurationMs());
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

function summarizeChatTypes(chats = []) {
  return chats.reduce(
    (acc, chat) => {
      const chatType = String(chat?.chatType || 'unknown');
      acc[chatType] = (acc[chatType] || 0) + 1;
      return acc;
    },
    { oneOnOne: 0, group: 0, meeting: 0, unknown: 0 },
  );
}

function createMemberLookupController(config = getTeamsArchiveConfig()) {
  return {
    mode: config.memberEnrichmentMode,
    failureThreshold: config.memberEnrichmentFailureThreshold,
    disabledChatTypes: new Set(),
    stats: {
      skippedByMode: 0,
      skippedByCircuitBreaker: 0,
      successCount: 0,
      failureCount: 0,
      failureByChatType: {},
    },
  };
}

function shouldAttemptMemberLookup(chatType, controller) {
  if (!controller) {
    return true;
  }

  if (controller.mode === 'disabled') {
    controller.stats.skippedByMode += 1;
    return false;
  }

  if (controller.mode === 'non_meeting' && chatType === 'meeting') {
    controller.stats.skippedByMode += 1;
    return false;
  }

  if (controller.disabledChatTypes.has(chatType)) {
    controller.stats.skippedByCircuitBreaker += 1;
    return false;
  }

  return true;
}

function recordMemberLookupFailure(chatType, controller) {
  if (!controller) {
    return false;
  }

  controller.stats.failureCount += 1;
  controller.stats.failureByChatType[chatType] =
    (controller.stats.failureByChatType[chatType] || 0) + 1;

  if (
    controller.mode === 'adaptive' &&
    controller.stats.failureByChatType[chatType] >= controller.failureThreshold &&
    !controller.disabledChatTypes.has(chatType)
  ) {
    controller.disabledChatTypes.add(chatType);
    return true;
  }

  return false;
}

async function mapWithConcurrency(items, limit, mapper) {
  const results = new Array(items.length);
  let nextIndex = 0;

  async function worker() {
    while (nextIndex < items.length) {
      const currentIndex = nextIndex;
      nextIndex += 1;
      results[currentIndex] = await mapper(items[currentIndex], currentIndex);
    }
  }

  const workerCount = Math.max(1, Math.min(limit, items.length));
  await Promise.all(Array.from({ length: workerCount }, () => worker()));
  return results;
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
    ...(openidId ? [{ fromUserId: openidId }] : []),
    ...(normalizedEmail ? [{ fromEmail: normalizedEmail }, { fromEmail: user?.email }] : []),
    ...(normalizedName ? [{ fromDisplayName: normalizedName }] : []),
    ...(normalizedUsername ? [{ fromDisplayName: normalizedUsername }] : []),
  ];
}

function getUserMessageIdentityFilter(user, userId) {
  const senderOr = getUserSenderClauses(user);
  return {
    user: userId,
    ...(senderOr.length > 0 ? { $or: senderOr } : {}),
  };
}

const SEARCHABLE_MESSAGE_FIELDS = [
  'bodyText',
  'bodyPreview',
  'bodyContent',
  'summary',
  'subject',
  'fromDisplayName',
  'fromEmail',
  'attachments.name',
  'mentions.displayName',
];

function buildFieldOrClause(fields, regex) {
  return {
    $or: fields.map((field) => ({ [field]: regex })),
  };
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

function buildTopicSearchClauses(value) {
  const phraseRegex = buildSearchRegex(value);
  const termRegexes = buildTopicTerms(value).map((term) => buildSearchRegex(term)).filter(Boolean);
  const clauses = [];

  if (phraseRegex) {
    clauses.push(buildFieldOrClause(SEARCHABLE_MESSAGE_FIELDS, phraseRegex));
  }

  if (termRegexes.length > 1) {
    clauses.push({
      $and: termRegexes.map((regex) => buildFieldOrClause(SEARCHABLE_MESSAGE_FIELDS, regex)),
    });
  }

  return {
    phraseRegex,
    termRegexes,
    clauses,
  };
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

function getNestedFieldValues(value, pathSegments) {
  if (pathSegments.length === 0) {
    if (Array.isArray(value)) {
      return value.flatMap((entry) => getNestedFieldValues(entry, []));
    }
    return value === undefined || value === null ? [] : [value];
  }

  if (Array.isArray(value)) {
    return value.flatMap((entry) => getNestedFieldValues(entry, pathSegments));
  }

  if (value === undefined || value === null || typeof value !== 'object') {
    return [];
  }

  const [head, ...tail] = pathSegments;
  return getNestedFieldValues(value[head], tail);
}

function messageMatchesRegex(message, regex) {
  if (!regex) {
    return true;
  }

  return SEARCHABLE_MESSAGE_FIELDS.some((field) => {
    const values = getNestedFieldValues(message, field.split('.'));
    return values.some((value) => regex.test(String(value || '')));
  });
}

function hasNonEmptyMemoryResults(result) {
  return Boolean(result && Array.isArray(result.results) && result.results.length > 0);
}

function mapMessageResult(message, conversation) {
  return {
    id: message._id?.toString?.() || message.id,
    graphMessageId: message.graphMessageId,
    graphChatId: message.graphChatId,
    topic: conversation?.topic || '',
    chatType: conversation?.chatType || '',
    participants: conversation?.participants || [],
    fromDisplayName: message.fromDisplayName || '',
    fromEmail: message.fromEmail || '',
    subject: message.subject || '',
    summary: message.summary || '',
    bodyPreview: message.bodyPreview || '',
    bodyText: message.bodyText || '',
    attachments: message.attachments || [],
    mentions: message.mentions || [],
    sentDateTime: message.sentDateTime,
    webUrl: message.webUrl || '',
  };
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

function mapCompactConversation(conversation) {
  return {
    id: conversation._id?.toString?.() || conversation.id,
    graphChatId: conversation.graphChatId,
    chatType: conversation.chatType || '',
    topic: conversation.topic || '',
    participants: mapCompactParticipants(conversation.participants || []),
    webUrl: conversation.webUrl || '',
    lastMessageAt: conversation.lastMessageAt,
    lastSyncedAt: conversation.lastSyncedAt,
    sourceUpdatedAt: conversation.sourceUpdatedAt,
    messageCount: conversation.messageCount || 0,
  };
}

function mapConversationCandidate(conversation) {
  return {
    id: conversation._id?.toString?.() || conversation.id,
    graphChatId: conversation.graphChatId,
    chatType: conversation.chatType || '',
    topic: conversation.topic || '',
    participants: mapCompactParticipants(conversation.participants || [], 8),
    firstMessageAt: conversation.firstMessageAt || null,
    lastMessageAt: conversation.lastMessageAt || null,
    lastSyncedAt: conversation.lastSyncedAt || null,
    messageCount: conversation.messageCount || 0,
    syncStatus: conversation.syncStatus || '',
    webUrl: conversation.webUrl || '',
  };
}

function mapCompactMessageResult(message, conversation, options = {}) {
  const {
    excerptLimit = 280,
    includeConversation = true,
    includeReplyToId = false,
    includeImportance = false,
  } = options;

  const excerptSource = message.bodyPreview || message.summary || message.bodyText || '';

  return {
    id: message._id?.toString?.() || message.id,
    graphMessageId: message.graphMessageId,
    graphChatId: message.graphChatId,
    ...(includeConversation
      ? {
          topic: conversation?.topic || '',
          chatType: conversation?.chatType || '',
          participants: mapCompactParticipants(conversation?.participants || []),
        }
      : {}),
    fromDisplayName: message.fromDisplayName || '',
    fromEmail: message.fromEmail || '',
    subject: message.subject || '',
    summary: truncateText(message.summary || '', 180),
    excerpt: truncateText(excerptSource, excerptLimit),
    attachmentNames: toArray(message.attachments)
      .map((attachment) => attachment?.name)
      .filter(Boolean)
      .slice(0, 3),
    mentions: toArray(message.mentions)
      .map((mention) => mention?.displayName)
      .filter(Boolean)
      .slice(0, 5),
    sentDateTime: message.sentDateTime,
    webUrl: message.webUrl || '',
    ...(includeReplyToId ? { replyToId: message.replyToId || '' } : {}),
    ...(includeImportance ? { importance: message.importance || '' } : {}),
  };
}

function summarizeSenderCounts(messages = []) {
  const senderCounts = new Map();

  for (const message of messages) {
    const key =
      String(message?.fromDisplayName || '').trim() ||
      String(message?.fromEmail || '').trim() ||
      'Unknown sender';
    senderCounts.set(key, (senderCounts.get(key) || 0) + 1);
  }

  return [...senderCounts.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5)
    .map(([sender, count]) => ({ sender, count }));
}

function summarizeConversationText({
  topic,
  chatType,
  totalMessages,
  matchedMessages,
  daysBack,
  topSenders,
}) {
  const segments = [];
  const normalizedTopic = String(topic || '').trim();

  if (normalizedTopic) {
    segments.push(`Conversation topic: ${normalizedTopic}.`);
  } else {
    segments.push('Conversation topic is not labeled in Teams.');
  }

  if (chatType) {
    segments.push(`Chat type: ${chatType}.`);
  }

  if (typeof daysBack === 'number') {
    segments.push(`Window searched: last ${daysBack} days.`);
  }

  segments.push(
    matchedMessages === totalMessages
      ? `${totalMessages} archived messages were included in the summary.`
      : `${matchedMessages} matching messages were found out of ${totalMessages} archived messages in scope.`,
  );

  if (topSenders.length > 0) {
    segments.push(
      `Most active senders in scope: ${topSenders
        .map(({ sender, count }) => `${sender} (${count})`)
        .join(', ')}.`,
    );
  }

  return segments.join(' ');
}

function summarizeCoverageByMonth(messages = [], regex = null) {
  const byMonth = new Map();

  for (const message of messages) {
    const sentDate = toDate(message?.sentDateTime) || toDate(message?.createdAt);
    if (!sentDate) {
      continue;
    }

    const monthKey = sentDate.toISOString().slice(0, 7);
    const current = byMonth.get(monthKey) || {
      month: monthKey,
      totalMessages: 0,
      matchedMessages: 0,
      firstMessageAt: sentDate,
      lastMessageAt: sentDate,
    };

    current.totalMessages += 1;
    if (regex && messageMatchesRegex(message, regex)) {
      current.matchedMessages += 1;
    }

    if (!current.firstMessageAt || sentDate < current.firstMessageAt) {
      current.firstMessageAt = sentDate;
    }
    if (!current.lastMessageAt || sentDate > current.lastMessageAt) {
      current.lastMessageAt = sentDate;
    }

    byMonth.set(monthKey, current);
  }

  return [...byMonth.values()].sort((a, b) => a.month.localeCompare(b.month));
}

function mapDossierHighlights(messages = [], conversation, limit = 8) {
  return messages.slice(0, limit).map((message) =>
    mapCompactMessageResult(message, conversation, {
      includeReplyToId: true,
      excerptLimit: 360,
    }),
  );
}

function buildConversationLookupFilter(userId, options = {}) {
  const chatType = String(options.chatType || 'any').trim();
  const validChatTypes = new Set(['any', 'oneOnOne', 'group', 'meeting']);
  const normalizedChatType = validChatTypes.has(chatType) ? chatType : 'any';
  const participantClauses = buildParticipantConversationClauses(options.participants);
  const topic = String(options.topic || options.query || '').trim();
  const topicRegex = topic ? buildSearchRegex(topic) : null;
  const daysBack = options.daysBack
    ? clampInteger(options.daysBack, 30, { min: 1, max: 3650 })
    : undefined;

  return {
    normalizedChatType,
    topic,
    topicRegex,
    daysBack,
    participantClauses,
    filter: {
      user: userId,
      ...(normalizedChatType !== 'any' ? { chatType: normalizedChatType } : {}),
      ...(daysBack
        ? { lastMessageAt: { $gte: new Date(Date.now() - daysBack * 24 * 60 * 60 * 1000) } }
        : {}),
      ...((participantClauses.length > 0 || topicRegex)
        ? {
            $and: [
              ...(topicRegex ? [buildFieldOrClause(['topic'], topicRegex)] : []),
              ...participantClauses,
            ],
          }
        : {}),
    },
  };
}

async function findConversationCandidates(userId, options = {}) {
  const requestedLimit =
    options.candidateLimit ?? options.limit ?? 10;
  const limit = clampInteger(requestedLimit, 10, { min: 1, max: 1000 });
  const lookup = buildConversationLookupFilter(userId, options);
  let conversations = await db.findTeamsArchiveConversations(lookup.filter, {
    limit,
    offset: 0,
    sort: { lastMessageAt: -1, updatedAt: -1 },
  });

  if (conversations.length === 0 && lookup.participantClauses.length > 0) {
    const participantRegexes = toArray(options.participants)
      .map((participant) => buildSearchRegex(String(participant || '').trim()))
      .filter(Boolean);

    if (participantRegexes.length > 0) {
      const senderFallbackMessages = await db.findTeamsArchiveMessages(
        {
          user: userId,
          ...(lookup.daysBack
            ? {
                sentDateTime: {
                  $gte: new Date(Date.now() - lookup.daysBack * 24 * 60 * 60 * 1000),
                },
              }
            : {}),
          $or: participantRegexes.flatMap((regex) => [
            { fromDisplayName: regex },
            { fromEmail: regex },
          ]),
        },
        {
          limit: 2000,
          offset: 0,
          sort: { sentDateTime: -1, createdAt: -1 },
        },
      );

      const derivedConversationIds = [
        ...new Set(
          senderFallbackMessages.map((message) => message.graphChatId).filter(Boolean),
        ),
      ];

      if (derivedConversationIds.length > 0) {
        logger.info('[TeamsArchiveService] Resolving conversation candidates via message-sender fallback', {
          userId,
          participantCount: participantRegexes.length,
          derivedConversationCount: derivedConversationIds.length,
          chatType: lookup.normalizedChatType,
          daysBack: lookup.daysBack,
        });

        const fallbackConversationFilter = {
          user: userId,
          graphChatId: { $in: derivedConversationIds },
          ...(lookup.normalizedChatType !== 'any' ? { chatType: lookup.normalizedChatType } : {}),
          ...(lookup.daysBack
            ? {
                lastMessageAt: {
                  $gte: new Date(Date.now() - lookup.daysBack * 24 * 60 * 60 * 1000),
                },
              }
            : {}),
          ...(lookup.topicRegex
            ? { $and: [buildFieldOrClause(['topic'], lookup.topicRegex)] }
            : {}),
        };

        conversations = await db.findTeamsArchiveConversations(
          fallbackConversationFilter,
          {
            limit,
            offset: 0,
            sort: { lastMessageAt: -1, updatedAt: -1 },
          },
        );
      }
    }
  }

  return {
    ...lookup,
    conversations,
  };
}

function isRecoverableChatMessageError(error) {
  return (
    error?.name === 'TeamsArchiveServiceError' &&
    (error?.status === 403 || error?.status === 404)
  );
}

function getSyncHeartbeatReference(job) {
  return (
    toDate(job?.updatedAt) ||
    toDate(job?.startedAt) ||
    toDate(job?.createdAt) ||
    null
  );
}

function isRunningSyncJobStale(job) {
  if (!job || job.status !== 'running') {
    return false;
  }

  const reference = getSyncHeartbeatReference(job);
  if (!reference) {
    return false;
  }

  const staleMinutes = getTeamsArchiveConfig().syncStaleMinutes;
  return Date.now() - reference.getTime() > staleMinutes * 60 * 1000;
}

async function reconcileRunningSyncJob(userId) {
  const latestRunningJob = await db.findLatestTeamsArchiveSyncJob({
    user: userId,
    status: 'running',
  });

  if (!latestRunningJob) {
    return null;
  }

  if (!isRunningSyncJobStale(latestRunningJob)) {
    return latestRunningJob;
  }

  logger.warn('[TeamsArchiveService] Reconciling stale running sync job', {
    userId,
    syncJobId: latestRunningJob._id?.toString?.() || latestRunningJob.id,
    startedAt: latestRunningJob.startedAt,
    updatedAt: latestRunningJob.updatedAt,
  });

  return db.updateTeamsArchiveSyncJob(latestRunningJob._id?.toString?.() || latestRunningJob.id, {
    status: 'failure',
    errorMessage: 'Sync interrupted by restart or worker shutdown',
    completedAt: new Date(),
  });
}

async function heartbeatSyncJob(syncJobId, updates = {}) {
  return db.updateTeamsArchiveSyncJob(syncJobId, updates);
}

async function acquireSyncLease({ leaseKey, leaseType, ownerToken, userId }) {
  return runAsSystem(async () =>
    db.acquireTeamsArchiveSyncLease({
      leaseKey,
      leaseType,
      ownerToken,
      user: userId,
      leaseExpiresAt: getLeaseExpiryDate(),
      lastHeartbeatAt: new Date(),
    }),
  );
}

async function refreshSyncLease(leaseKey, ownerToken) {
  return runAsSystem(async () =>
    db.refreshTeamsArchiveSyncLease(leaseKey, ownerToken, getLeaseExpiryDate()),
  );
}

async function releaseSyncLease(leaseKey, ownerToken) {
  return runAsSystem(async () => db.releaseTeamsArchiveSyncLease(leaseKey, ownerToken));
}

async function countActiveSyncSlots() {
  return runAsSystem(async () =>
    db.countActiveTeamsArchiveSyncLeases({
      leaseType: 'slot',
      leaseExpiresAt: { $gt: new Date() },
    }),
  );
}

async function acquireGlobalSyncSlot(ownerToken, userId, maxConcurrentSyncs) {
  if (!maxConcurrentSyncs || maxConcurrentSyncs <= 0) {
    return null;
  }

  for (let slotNumber = 0; slotNumber < maxConcurrentSyncs; slotNumber += 1) {
    const leaseKey = getSlotLeaseKey(slotNumber);
    const lease = await acquireSyncLease({
      leaseKey,
      leaseType: 'slot',
      ownerToken,
      userId,
    });

    if (lease) {
      return leaseKey;
    }
  }

  return null;
}

async function heartbeatSyncExecution(syncJobId, updates, leaseContext = {}) {
  const operations = [heartbeatSyncJob(syncJobId, updates)];

  if (leaseContext.userLeaseKey && leaseContext.ownerToken) {
    operations.push(refreshSyncLease(leaseContext.userLeaseKey, leaseContext.ownerToken));
  }

  if (leaseContext.slotLeaseKey && leaseContext.ownerToken) {
    operations.push(refreshSyncLease(leaseContext.slotLeaseKey, leaseContext.ownerToken));
  }

  await Promise.all(operations);
}

async function getSyncJobById(syncJobId) {
  if (typeof db.findTeamsArchiveSyncJobById !== 'function') {
    return null;
  }

  return db.findTeamsArchiveSyncJobById(syncJobId);
}

async function ensureSyncJobActive(syncJobId) {
  const currentJob = await getSyncJobById(syncJobId);

  if (currentJob?.status === 'cancelled') {
    throw new TeamsArchiveSyncCancelledError();
  }

  return currentJob;
}

async function listChatsPage(user, { top = DEFAULT_CHAT_LIMIT, nextLink } = {}) {
  if (nextLink) {
    return graphRequest(user, nextLink);
  }

  return graphRequest(user, '/me/chats', {
    query: {
      $top: top,
    },
  });
}

function isDiscoveryRefreshDue(backfillState) {
  if (!backfillState?.discoveryComplete) {
    return false;
  }

  const reference =
    toDate(backfillState?.lastDiscoveredAt) ||
    toDate(backfillState?.updatedAt) ||
    toDate(backfillState?.lastCompletedAt);

  if (!reference) {
    return true;
  }

  const refreshMs = getTeamsArchiveConfig().discoveryRefreshHours * 60 * 60 * 1000;
  return Date.now() - reference.getTime() >= refreshMs;
}

async function listChatMembers(user, chatId, chatType, controller) {
  const normalizedChatType = String(chatType || 'unknown');
  if (!shouldAttemptMemberLookup(normalizedChatType, controller)) {
    return [];
  }

  try {
    const response = await graphRequest(user, `/chats/${encodeURIComponent(chatId)}/members`, {
      query: { $top: 50 },
      suppressErrorLog: true,
    });

    if (controller) {
      controller.stats.successCount += 1;
    }

    return toArray(response?.value).map((member) => ({
      displayName: member?.displayName || member?.email || '',
      email: member?.email || '',
      userId: member?.userId || member?.id || '',
    }));
  } catch (error) {
    const disabledByCircuitBreaker = recordMemberLookupFailure(normalizedChatType, controller);
    logger.warn('[TeamsArchiveService] Failed to list chat members', {
      chatId,
      chatType: normalizedChatType,
      error: error?.message,
    });

    if (disabledByCircuitBreaker) {
      logger.info('[TeamsArchiveService] Disabling member enrichment for chat type during this sync', {
        chatType: normalizedChatType,
        failureThreshold: controller?.failureThreshold,
      });
    }

    return [];
  }
}

async function listChatMessagesPage(user, chatId, { top = DEFAULT_MESSAGES_PER_CHAT, nextLink } = {}) {
  const response = await graphRequest(
    user,
    nextLink || `/chats/${encodeURIComponent(chatId)}/messages`,
    nextLink
      ? {}
      : {
          query: {
            $top: Math.min(top, 50),
          },
        },
  );

  return {
    messages: toArray(response?.value),
    nextLink: response?.['@odata.nextLink'] || null,
  };
}

function computeBackfillLifecycleStatus({
  discoveredChatCount = 0,
  completedChatCount = 0,
  discoveryComplete,
  nextChatPageLink,
  pendingChatCount = 0,
  runningChatCount = 0,
  failedChatCount = 0,
  hasActiveSync = false,
}) {
  const hasOutstandingWork =
    !discoveryComplete || Boolean(nextChatPageLink) || pendingChatCount > 0 || runningChatCount > 0;
  const hasArchivedProgress =
    discoveredChatCount > 0 ||
    completedChatCount > 0 ||
    failedChatCount > 0 ||
    pendingChatCount > 0 ||
    runningChatCount > 0;

  if (hasActiveSync) {
    if (!discoveryComplete || nextChatPageLink) {
      return 'discovering';
    }

    return 'syncing';
  }

  if (hasOutstandingWork && hasArchivedProgress) {
    return 'paused';
  }

  if (failedChatCount > 0) {
    return 'failed';
  }

  if (discoveryComplete && discoveredChatCount > 0) {
    return 'complete';
  }

  return 'idle';
}

async function refreshBackfillStateSnapshot(userId, updates = {}, options = {}) {
  const { includeMessageCount = true, hasActiveSync = false } = options;
  const [discoveredChatCount, completedChatCount, pendingChatCount, runningChatCount, failedChatCount, totalMessageCount] =
    await Promise.all([
      db.countTeamsArchiveConversations({ user: userId }),
      db.countTeamsArchiveConversations({ user: userId, syncStatus: 'complete' }),
      db.countTeamsArchiveConversations({ user: userId, syncStatus: 'pending' }),
      db.countTeamsArchiveConversations({ user: userId, syncStatus: 'running' }),
      db.countTeamsArchiveConversations({ user: userId, syncStatus: 'failed' }),
      includeMessageCount
        ? db.countTeamsArchiveMessages({ user: userId })
        : Promise.resolve(undefined),
    ]);

  const nextState = {
    user: userId,
    discoveredChatCount,
    completedChatCount,
    pendingChatCount,
    runningChatCount,
    failedChatCount,
    ...(includeMessageCount && totalMessageCount !== undefined ? { totalMessageCount } : {}),
    lastHeartbeatAt: new Date(),
    ...updates,
  };

  if (!nextState.status || ['discovering', 'syncing', 'complete', 'paused', 'idle'].includes(nextState.status)) {
    nextState.status = computeBackfillLifecycleStatus({
      discoveredChatCount,
      completedChatCount,
      discoveryComplete: nextState.discoveryComplete,
      nextChatPageLink: nextState.nextChatPageLink,
      pendingChatCount,
      runningChatCount,
      failedChatCount,
      hasActiveSync,
    });
  }

  return db.upsertTeamsArchiveBackfillState(nextState);
}

function buildSyncJobCheckpoint({
  phase,
  nextChatPageLink,
  discoveryComplete,
  pageNumber,
  discoveredThisRun,
  processedChats,
  persistedMessages,
}) {
  return {
    phase,
    nextChatPageLink: nextChatPageLink || null,
    discoveryComplete: Boolean(discoveryComplete),
    pageNumber,
    discoveredThisRun,
    processedChats,
    persistedMessages,
  };
}

function queueTeamsProjection(params) {
  if (typeof projectTeamsArchiveSyncToMemory !== 'function') {
    return null;
  }

  const queuedAt = new Date();
  void projectTeamsArchiveSyncToMemory(params).catch((error) => {
    logger.error('[TeamsArchiveService] Background Teams projection failed', {
      userId: params?.userId,
      syncJobId: params?.syncJobId,
      error: error?.message || error,
    });
  });

  return {
    status: 'queued',
    queuedAt,
    requestedConversationCount: Array.isArray(params?.graphChatIds) ? params.graphChatIds.length : 0,
  };
}

function normalizeProjectionDiagnostics(stats = {}) {
  const diagnostics = stats?.projectionDiagnostics || {};
  const totalMessagesLoaded = Number(diagnostics.totalMessagesLoaded || 0);
  const totalChunkableMessages = Number(diagnostics.totalChunkableMessages || 0);
  const totalSkippedEmptyTextMessages = Number(diagnostics.totalSkippedEmptyTextMessages || 0);

  return {
    missingConversationCount: Number(diagnostics.missingConversationCount || 0),
    zeroMessageConversationCount: Number(diagnostics.zeroMessageConversationCount || 0),
    zeroChunkConversationCount: Number(diagnostics.zeroChunkConversationCount || 0),
    truncatedConversationCount: Number(diagnostics.truncatedConversationCount || 0),
    totalMessagesLoaded,
    totalChunkableMessages,
    totalSkippedEmptyTextMessages,
    projectionMessageFetchLimit: Number(diagnostics.projectionMessageFetchLimit || 0),
    chunkableMessageRate:
      totalMessagesLoaded > 0 ? Number(((totalChunkableMessages / totalMessagesLoaded) * 100).toFixed(1)) : 0,
    skippedEmptyTextRate:
      totalMessagesLoaded > 0
        ? Number(((totalSkippedEmptyTextMessages / totalMessagesLoaded) * 100).toFixed(1))
        : 0,
  };
}

async function getStatus(user) {
  const config = getTeamsArchiveConfig();
  const userId = user?.id || user?._id?.toString();
  await reconcileRunningSyncJob(userId);
  const [
    conversationCount,
    messageCount,
    latestSync,
    latestProjection,
    activeSyncs,
    backfillState,
    projectionChunkCount,
    projectionEntityConversationCount,
    projectionConversationCount,
  ] = await Promise.all([
    userId ? db.countTeamsArchiveConversations({ user: userId }) : 0,
    userId ? db.countTeamsArchiveMessages({ user: userId }) : 0,
    userId ? db.findLatestTeamsArchiveSyncJob({ user: userId }) : null,
    typeof db.findLatestEnterpriseMemoryJob === 'function'
      ? db.findLatestEnterpriseMemoryJob({ user: userId, source: 'teams', jobType: 'projection' })
      : null,
    countActiveSyncSlots(),
    userId && typeof db.getTeamsArchiveBackfillState === 'function'
      ? db.getTeamsArchiveBackfillState(userId)
      : null,
    userId && typeof db.countEnterpriseMemoryChunks === 'function'
      ? db.countEnterpriseMemoryChunks({
          user: userId,
          source: 'teams',
          sourceRecordType: 'teams_message',
        })
      : 0,
    userId && typeof db.countEnterpriseMemoryEntities === 'function'
      ? db.countEnterpriseMemoryEntities({
          user: userId,
          source: 'teams',
          entityType: 'conversation',
          sourceRecordType: 'teams_chat',
        })
      : 0,
    userId && typeof db.countDistinctEnterpriseMemoryChunkField === 'function'
      ? db.countDistinctEnterpriseMemoryChunkField('sourceParentRecordId', {
          user: userId,
          source: 'teams',
          sourceRecordType: 'teams_message',
        })
      : 0,
  ]);

  const normalizedBackfillStatus = backfillState
    ? computeBackfillLifecycleStatus({
        discoveredChatCount: backfillState.discoveredChatCount || 0,
        completedChatCount: backfillState.completedChatCount || 0,
        discoveryComplete: Boolean(backfillState.discoveryComplete),
        nextChatPageLink: backfillState.nextChatPageLink,
        pendingChatCount: backfillState.pendingChatCount || 0,
        runningChatCount: backfillState.runningChatCount || 0,
        failedChatCount: backfillState.failedChatCount || 0,
        hasActiveSync: latestSync?.status === 'running',
      })
    : null;

  return {
    enabled: config.enabled,
    graphBaseUrl: config.graphBaseUrl,
    graphScopes: config.scopes,
    maxConcurrentSyncs: config.maxConcurrentSyncs,
    activeSyncs,
    syncModes: ['chats'],
    channelSyncSupported: false,
    conversationCount,
    messageCount,
    backfillState: backfillState
      ? {
          status: normalizedBackfillStatus || backfillState.status,
          discoveryComplete: Boolean(backfillState.discoveryComplete),
          nextChatPageLinkPresent: Boolean(backfillState.nextChatPageLink),
          discoveredChatCount: backfillState.discoveredChatCount || 0,
          completedChatCount: backfillState.completedChatCount || 0,
          pendingChatCount: backfillState.pendingChatCount || 0,
          runningChatCount: backfillState.runningChatCount || 0,
          failedChatCount: backfillState.failedChatCount || 0,
          totalMessageCount: backfillState.totalMessageCount || 0,
          lastSyncJobId: backfillState.lastSyncJobId || null,
          lastProjectionJobId: backfillState.lastProjectionJobId || null,
          lastDiscoveredAt: backfillState.lastDiscoveredAt || null,
          lastCompletedAt: backfillState.lastCompletedAt || null,
          lastHeartbeatAt: backfillState.lastHeartbeatAt || null,
          errorMessage: backfillState.errorMessage || null,
        }
      : null,
    latestSync: latestSync
      ? {
          id: latestSync._id?.toString?.() || latestSync.id,
          status: latestSync.status,
          mode: latestSync.mode,
          phase: latestSync.phase || null,
          checkpoint: latestSync.checkpoint || {},
          stats: latestSync.stats || {},
          requestedChatLimit: latestSync.requestedChatLimit || 0,
          requestedMessagesPerChat: latestSync.requestedMessagesPerChat || 0,
          discoveredChatCount: latestSync.discoveredChatCount || 0,
          processedChatCount: latestSync.processedChatCount || 0,
          skippedChatCount: latestSync.skippedChatCount || 0,
          projectionJobId: latestSync.projectionJobId || null,
          conversationCount: latestSync.conversationCount || 0,
          messageCount: latestSync.messageCount || 0,
          startedAt: latestSync.startedAt,
          completedAt: latestSync.completedAt,
          errorMessage: latestSync.errorMessage,
        }
      : null,
    latestProjection: latestProjection
      ? {
          id: latestProjection._id?.toString?.() || latestProjection.id,
          status: latestProjection.status,
          startedAt: latestProjection.startedAt,
          completedAt: latestProjection.completedAt,
          errorMessage: latestProjection.errorMessage,
          stats: latestProjection.stats || {},
          projectionDiagnostics: normalizeProjectionDiagnostics(latestProjection.stats || {}),
        }
      : null,
    projectionCoverage: {
      indexedConversationCount: projectionEntityConversationCount || 0,
      totalConversationCount: conversationCount || 0,
      indexedChunkCount: projectionChunkCount || 0,
      searchableConversationCount: projectionConversationCount || 0,
      pendingConversationCount:
        Math.max(0, (conversationCount || 0) - (projectionEntityConversationCount || 0)),
      fullyIndexed:
        conversationCount > 0
          ? (projectionEntityConversationCount || 0) >= conversationCount
          : false,
      coveragePercent:
        conversationCount > 0
          ? Number(
              ((((projectionEntityConversationCount || 0) / conversationCount) * 100).toFixed(1)),
            )
          : 0,
    },
  };
}

async function deleteUserArchive(user) {
  assertEnabled();
  const userId = user?.id || user?._id?.toString();

  if (!userId) {
    throw new TeamsArchiveServiceError('User id is required for Teams archive reset', 400);
  }

  const latestRunningJob = await reconcileRunningSyncJob(userId);
  if (latestRunningJob?.status === 'running') {
    throw new TeamsArchiveServiceError(
      'Cannot delete Teams archive data while a sync is running',
      409,
      {
        reason: 'sync_running',
        syncJobId: latestRunningJob._id?.toString?.() || latestRunningJob.id,
      },
    );
  }

  const deleted = await runAsSystem(async () => {
    const [
      conversations,
      messages,
      syncJobs,
      syncLeases,
      backfillStates,
      projectionJobs,
      chunks,
      entities,
      relationships,
    ] = await Promise.all([
      typeof db.deleteTeamsArchiveConversations === 'function'
        ? db.deleteTeamsArchiveConversations({ user: userId })
        : 0,
      typeof db.deleteTeamsArchiveMessages === 'function'
        ? db.deleteTeamsArchiveMessages({ user: userId })
        : 0,
      typeof db.deleteTeamsArchiveSyncJobs === 'function'
        ? db.deleteTeamsArchiveSyncJobs({ user: userId })
        : 0,
      typeof db.deleteTeamsArchiveSyncLeases === 'function'
        ? db.deleteTeamsArchiveSyncLeases({ user: userId })
        : 0,
      typeof db.deleteTeamsArchiveBackfillStates === 'function'
        ? db.deleteTeamsArchiveBackfillStates({ user: userId })
        : 0,
      typeof db.deleteEnterpriseMemoryJobs === 'function'
        ? db.deleteEnterpriseMemoryJobs({ user: userId, source: 'teams' })
        : 0,
      typeof db.deleteEnterpriseMemoryChunks === 'function'
        ? db.deleteEnterpriseMemoryChunks({ user: userId, source: 'teams' })
        : 0,
      typeof db.deleteEnterpriseMemoryEntities === 'function'
        ? db.deleteEnterpriseMemoryEntities({ user: userId, source: 'teams' })
        : 0,
      typeof db.deleteEnterpriseMemoryRelationships === 'function'
        ? db.deleteEnterpriseMemoryRelationships({ user: userId, source: 'teams' })
        : 0,
    ]);

    return {
      conversations,
      messages,
      syncJobs,
      syncLeases,
      backfillStates,
      projectionJobs,
      chunks,
      entities,
      relationships,
    };
  });

  logger.info('[TeamsArchiveService] Deleted archived Teams data for user', {
    userId,
    deleted,
  });

  return {
    deleted,
    message: 'Archived Teams data cleared. A fresh sync will start from scratch.',
  };
}

async function cancelRunningSync(user) {
  assertEnabled();
  const userId = user?.id || user?._id?.toString();
  const latestRunningJob = await reconcileRunningSyncJob(userId);

  if (!latestRunningJob || latestRunningJob.status !== 'running') {
    return {
      cancelled: false,
      status: 'idle',
      syncJob: null,
      message: 'No running Teams archive sync was found.',
    };
  }

  const cancelledJob = await db.updateTeamsArchiveSyncJob(
    latestRunningJob._id?.toString?.() || latestRunningJob.id,
    {
      status: 'cancelled',
      phase: 'cancelled',
      errorMessage: 'Sync cancelled by user',
      completedAt: new Date(),
    },
  );

  await refreshBackfillStateSnapshot(userId, {
    status: 'paused',
    lastSyncJobId: cancelledJob?._id?.toString?.() || latestRunningJob._id?.toString?.() || latestRunningJob.id,
    errorMessage: 'Sync cancelled by user',
  }, { includeMessageCount: false, hasActiveSync: false });

  return {
    cancelled: true,
    status: 'cancelled',
    syncJob: cancelledJob,
    message: 'Teams archive sync cancellation requested.',
  };
}

async function getSyncStartAvailability(user) {
  assertEnabled();
  assertDelegatedUser(user);

  const userId = user?.id || user?._id?.toString();
  if (!userId) {
    throw new TeamsArchiveServiceError('User id is required for Teams archive sync', 400);
  }

  const config = getTeamsArchiveConfig();
  const latestRunningJob = await reconcileRunningSyncJob(userId);
  if (latestRunningJob?.status === 'running') {
    return {
      allowed: false,
      reason: 'already_running',
      status: 202,
      syncJob: latestRunningJob,
      message: 'A Teams archive sync is already running for this user.',
    };
  }

  const activeUserLeases = await runAsSystem(async () =>
    db.countActiveTeamsArchiveSyncLeases({
      leaseKey: getUserLeaseKey(userId),
      leaseType: 'user',
      leaseExpiresAt: { $gt: new Date() },
    }),
  );

  if (activeUserLeases > 0) {
    return {
      allowed: false,
      reason: 'user_lock',
      status: 409,
      message: 'A Teams archive sync is already in progress or still shutting down for this user.',
    };
  }

  const activeSyncs = await countActiveSyncSlots();
  if (config.maxConcurrentSyncs > 0 && activeSyncs >= config.maxConcurrentSyncs) {
    return {
      allowed: false,
      reason: 'capacity',
      status: 429,
      message: 'Teams archive sync capacity is full. Try again in a few minutes.',
      details: {
        activeSyncs,
        maxConcurrentSyncs: config.maxConcurrentSyncs,
      },
    };
  }

  return {
    allowed: true,
    activeSyncs,
    maxConcurrentSyncs: config.maxConcurrentSyncs,
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
  const chatLimit = clampInteger(options.chatLimit, config.defaultChatLimit, { max: 10000 });
  const messagesPerChat = clampInteger(options.messagesPerChat, config.defaultMessagesPerChat, {
    max: 5000,
  });
  const mode = options.mode === 'chats' ? 'chats' : 'chats';
  const latestRunningJob = await reconcileRunningSyncJob(userId);
  const ownerToken = randomUUID();
  const userLeaseKey = getUserLeaseKey(userId);
  let slotLeaseKey = null;
  let syncJob = null;

  if (latestRunningJob?.status === 'running') {
    return {
      syncJob: latestRunningJob,
      mode,
      conversationCount: latestRunningJob.conversationCount || 0,
      messageCount: latestRunningJob.messageCount || 0,
      conversations: [],
      memoryProjection: null,
      alreadyRunning: true,
    };
  }

  try {
    const userLease = await acquireSyncLease({
      leaseKey: userLeaseKey,
      leaseType: 'user',
      ownerToken,
      userId,
    });

    if (!userLease) {
      const concurrentRunningJob = await reconcileRunningSyncJob(userId);
      if (concurrentRunningJob?.status === 'running') {
        return {
          syncJob: concurrentRunningJob,
          mode,
          conversationCount: concurrentRunningJob.conversationCount || 0,
          messageCount: concurrentRunningJob.messageCount || 0,
          conversations: [],
          memoryProjection: null,
          alreadyRunning: true,
        };
      }

      throw new TeamsArchiveServiceError(
        'A Teams archive sync is already in progress or still shutting down for this user',
        409,
      );
    }

    slotLeaseKey = await acquireGlobalSyncSlot(ownerToken, userId, config.maxConcurrentSyncs);

    if (config.maxConcurrentSyncs > 0 && !slotLeaseKey) {
      const activeSyncs = await countActiveSyncSlots();
      logger.warn('[TeamsArchiveService] Global Teams sync capacity reached', {
        userId,
        activeSyncs,
        maxConcurrentSyncs: config.maxConcurrentSyncs,
      });

      throw new TeamsArchiveServiceError(
        'Teams archive sync capacity is full. Try again in a few minutes.',
        429,
        {
          activeSyncs,
          maxConcurrentSyncs: config.maxConcurrentSyncs,
        },
      );
    }

    const existingBackfillState =
      typeof db.getTeamsArchiveBackfillState === 'function'
        ? await db.getTeamsArchiveBackfillState(userId)
        : null;
    const shouldRefreshDiscovery = isDiscoveryRefreshDue(existingBackfillState);
    let nextChatPageLink =
      shouldRefreshDiscovery || !existingBackfillState?.discoveryComplete
        ? existingBackfillState?.nextChatPageLink || null
        : null;
    let discoveryComplete =
      shouldRefreshDiscovery ? false : Boolean(existingBackfillState?.discoveryComplete);

    syncJob = await db.createTeamsArchiveSyncJob({
      user: userId,
      status: 'running',
      mode,
      phase: 'discovering_chats',
      checkpoint: buildSyncJobCheckpoint({
        phase: 'discovering_chats',
        nextChatPageLink,
        discoveryComplete,
        pageNumber: 0,
        discoveredThisRun: 0,
        processedChats: 0,
        persistedMessages: 0,
      }),
      requestedChatLimit: chatLimit,
      requestedMessagesPerChat: messagesPerChat,
      conversationCount: 0,
      messageCount: 0,
      startedAt: new Date(),
    });

    await refreshBackfillStateSnapshot(userId, {
      status: 'discovering',
      nextChatPageLink,
      discoveryComplete,
      lastSyncJobId: syncJob._id?.toString?.() || syncJob.id,
      errorMessage: '',
      lastDiscoveredAt: shouldRefreshDiscovery ? new Date() : existingBackfillState?.lastDiscoveredAt,
    }, { includeMessageCount: false, hasActiveSync: true });

    const syncedConversations = [];
    let processedChats = 0;
    let persistedMessages = 0;
    let skippedMessageChats = 0;
    let discoveredThisRun = 0;
    let pageNumber = 0;
    let lastHeartbeatAt = Date.now();
    const graphChatTypeSummary = { oneOnOne: 0, group: 0, meeting: 0, unknown: 0 };
    const processedChatTypeSummary = { oneOnOne: 0, group: 0, meeting: 0, unknown: 0 };
    const discoveryDecisionSummary = {
      noExistingConversation: 0,
      resumableCursor: 0,
      incompleteSyncStatus: 0,
      sourceUpdated: 0,
      alreadyComplete: 0,
    };
    const completedGraphChatIds = new Set();
    const memberLookupController = createMemberLookupController(config);

    while (!discoveryComplete && discoveredThisRun < chatLimit) {
      await ensureSyncJobActive(syncJob._id?.toString?.() || syncJob.id);
      pageNumber += 1;

      const response = await listChatsPage(user, {
        top: Math.min(chatLimit - discoveredThisRun, 50),
        nextLink: nextChatPageLink,
      });
      const chats = toArray(response?.value).sort((a, b) => {
        const aTime = toDate(a?.lastUpdatedDateTime)?.getTime() ?? 0;
        const bTime = toDate(b?.lastUpdatedDateTime)?.getTime() ?? 0;
        return bTime - aTime;
      });

      if (chats.length === 0) {
        discoveryComplete = true;
        nextChatPageLink = null;
        break;
      }

      const existingConversations = await db.findTeamsArchiveConversations(
        { user: userId, graphChatId: { $in: chats.map((chat) => chat.id) } },
        { limit: Math.max(chats.length, 1) },
      );
      const existingConversationMap = new Map(
        existingConversations.map((conversation) => [conversation.graphChatId, conversation]),
      );

      const pageChatTypeSummary = summarizeChatTypes(chats);
      for (const [chatType, count] of Object.entries(pageChatTypeSummary)) {
        graphChatTypeSummary[chatType] = (graphChatTypeSummary[chatType] || 0) + count;
      }

      logger.info('[TeamsArchiveService] Sync discovery page loaded', {
        userId,
        syncJobId: syncJob._id?.toString?.() || syncJob.id,
        pageNumber,
        chatsReturned: chats.length,
        discoveredThisRun,
        chatLimit,
        pageChatTypeSummary,
        hasNextPage: Boolean(response?.['@odata.nextLink']),
      });

      await mapWithConcurrency(chats, config.discoveryConcurrency, async (chat) => {
        await ensureSyncJobActive(syncJob._id?.toString?.() || syncJob.id);
        const chatType = String(chat?.chatType || 'unknown');
        const members = await listChatMembers(user, chat.id, chatType, memberLookupController);
        const normalizedConversation = normalizeConversation(chat, members);
        const existingConversation = existingConversationMap.get(chat.id);
        const sourceUpdatedAt =
          normalizedConversation.sourceUpdatedAt || existingConversation?.sourceUpdatedAt;
        const lastMessageSyncAt = existingConversation?.lastMessageSyncAt || null;
        const sourceUpdatedMs = sourceUpdatedAt?.getTime?.() ?? 0;
        const lastMessageSyncMs = lastMessageSyncAt ? new Date(lastMessageSyncAt).getTime() : 0;
        let syncDecision = 'alreadyComplete';

        if (!existingConversation) {
          syncDecision = 'noExistingConversation';
        } else if (Boolean(existingConversation.syncCursor)) {
          syncDecision = 'resumableCursor';
        } else if (existingConversation.syncStatus !== 'complete') {
          syncDecision = 'incompleteSyncStatus';
        } else if (sourceUpdatedMs > 0 && sourceUpdatedMs > lastMessageSyncMs) {
          syncDecision = 'sourceUpdated';
        }

        const needsSync = syncDecision !== 'alreadyComplete';
        discoveryDecisionSummary[syncDecision] =
          (discoveryDecisionSummary[syncDecision] || 0) + 1;

        await db.upsertTeamsArchiveConversation({
          user: userId,
          ...normalizedConversation,
          sourceDiscoveredAt: existingConversation?.sourceDiscoveredAt || new Date(),
          sourceLastMessageAt: normalizedConversation.sourceUpdatedAt,
          syncStatus: needsSync ? 'pending' : existingConversation?.syncStatus || 'complete',
          syncCursor: needsSync ? existingConversation?.syncCursor || undefined : undefined,
          syncError: needsSync ? '' : existingConversation?.syncError,
          syncStartedAt: existingConversation?.syncStartedAt,
          syncCompletedAt: existingConversation?.syncCompletedAt,
          lastMessageSyncAt: existingConversation?.lastMessageSyncAt,
          lastMessageAt: existingConversation?.lastMessageAt,
          lastSyncedAt: existingConversation?.lastSyncedAt,
          messageCount: existingConversation?.messageCount || 0,
        });
      });

      discoveredThisRun += chats.length;

      nextChatPageLink =
        response?.['@odata.nextLink'] && discoveredThisRun < chatLimit
          ? response['@odata.nextLink']
          : response?.['@odata.nextLink'] || null;
      discoveryComplete = !nextChatPageLink;

      await heartbeatSyncExecution(
        syncJob._id?.toString?.() || syncJob.id,
        {
          phase: 'discovering_chats',
          discoveredChatCount: discoveredThisRun,
          checkpoint: buildSyncJobCheckpoint({
            phase: 'discovering_chats',
            nextChatPageLink,
            discoveryComplete,
            pageNumber,
            discoveredThisRun,
            processedChats,
            persistedMessages,
          }),
          stats: {
            graphChatTypeSummary,
            processedChatTypeSummary,
            memberLookup: memberLookupController.stats,
          },
        },
        { userLeaseKey, slotLeaseKey, ownerToken },
      );
      lastHeartbeatAt = Date.now();

      await refreshBackfillStateSnapshot(userId, {
        nextChatPageLink,
        discoveryComplete,
        lastSyncJobId: syncJob._id?.toString?.() || syncJob.id,
        lastDiscoveredAt: new Date(),
      }, { includeMessageCount: false, hasActiveSync: true });

      if (!nextChatPageLink || discoveredThisRun >= chatLimit) {
        break;
      }
    }

    await heartbeatSyncExecution(
      syncJob._id?.toString?.() || syncJob.id,
      {
        phase: 'syncing_messages',
        checkpoint: buildSyncJobCheckpoint({
          phase: 'syncing_messages',
          nextChatPageLink,
          discoveryComplete,
          pageNumber,
          discoveredThisRun,
          processedChats,
          persistedMessages,
        }),
      },
      { userLeaseKey, slotLeaseKey, ownerToken },
    );

    const conversationsToSync = await db.findTeamsArchiveConversations(
      {
        user: userId,
        syncStatus: { $in: ['pending', 'running', 'failed'] },
      },
      {
        limit: chatLimit,
        sort: { sourceUpdatedAt: -1, sourceLastMessageAt: -1, updatedAt: -1 },
      },
    );

    for (const conversation of conversationsToSync) {
      await ensureSyncJobActive(syncJob._id?.toString?.() || syncJob.id);
      const conversationId = conversation._id?.toString?.() || conversation.id;
      const chatType = String(conversation?.chatType || 'unknown');
      const incrementalRefresh = Boolean(conversation?.syncCompletedAt && !conversation?.syncCursor);
      const incrementalCutoff = conversation?.lastMessageSyncAt
        ? new Date(conversation.lastMessageSyncAt)
        : null;
      let remainingMessageBudget = messagesPerChat;
      let nextMessageCursor = conversation?.syncCursor || null;
      let latestMessageAt = conversation?.lastMessageAt || null;
      let conversationFailed = false;
      let reachedIncrementalCutoff = false;

      await db.updateTeamsArchiveConversation(conversationId, {
        syncStatus: 'running',
        syncStartedAt: conversation?.syncStartedAt || new Date(),
        syncError: '',
      });

      try {
        while (remainingMessageBudget > 0) {
          const page = await listChatMessagesPage(user, conversation.graphChatId, {
            top: Math.min(remainingMessageBudget, 50),
            nextLink: nextMessageCursor,
          });

          const normalizedMessages = page.messages.map((message) => ({
            user: userId,
            ...normalizeMessage(conversation.graphChatId, message),
          }));

          if (normalizedMessages.length > 0) {
            persistedMessages += await db.bulkUpsertTeamsArchiveMessages(normalizedMessages);
            const sortedMessages = [...normalizedMessages]
              .filter((message) => message.sentDateTime instanceof Date)
              .sort((a, b) => b.sentDateTime.getTime() - a.sentDateTime.getTime());
            latestMessageAt = sortedMessages[0]?.sentDateTime || latestMessageAt;

            if (incrementalCutoff) {
              const pageHasNewerMessages = normalizedMessages.some(
                (message) =>
                  message.sentDateTime instanceof Date &&
                  message.sentDateTime.getTime() > incrementalCutoff.getTime(),
              );

              if (!pageHasNewerMessages) {
                reachedIncrementalCutoff = true;
              }
            }
          }

          nextMessageCursor = page.nextLink || null;
          remainingMessageBudget -= Math.max(normalizedMessages.length, 1);

          await db.updateTeamsArchiveConversation(conversationId, {
            syncCursor: incrementalRefresh ? undefined : nextMessageCursor || undefined,
            lastMessageSyncAt: new Date(),
            lastSyncedAt: new Date(),
            lastMessageAt: latestMessageAt || conversation?.lastMessageAt,
          });

          if (!nextMessageCursor || reachedIncrementalCutoff) {
            break;
          }
        }
      } catch (error) {
        if (!isRecoverableChatMessageError(error)) {
          throw error;
        }

        conversationFailed = true;
        skippedMessageChats += 1;
        logger.warn('[TeamsArchiveService] Failed to list chat messages; continuing sync', {
          userId,
          syncJobId: syncJob._id?.toString?.() || syncJob.id,
          chatId: conversation.graphChatId,
          chatType,
          status: error?.status,
          details: error?.details,
        });
      }

      const messageCountForConversation = await db.countTeamsArchiveMessages({
        user: userId,
        graphChatId: conversation.graphChatId,
      });
      const isConversationComplete = !conversationFailed && (incrementalRefresh || !nextMessageCursor);

      const conversationRecord = await db.updateTeamsArchiveConversation(conversationId, {
        syncStatus: conversationFailed ? 'failed' : isConversationComplete ? 'complete' : 'pending',
        syncCursor:
          conversationFailed || incrementalRefresh ? undefined : nextMessageCursor || undefined,
        syncError: conversationFailed ? 'Message sync skipped due to Graph permissions or missing chat data' : '',
        syncCompletedAt: isConversationComplete ? new Date() : undefined,
        lastMessageSyncAt: new Date(),
        lastSyncedAt: new Date(),
        lastMessageAt: latestMessageAt || conversation?.lastMessageAt,
        messageCount: messageCountForConversation,
      });

      if (isConversationComplete) {
        completedGraphChatIds.add(conversation.graphChatId);
      }

      syncedConversations.push({
        id: conversationRecord?._id?.toString?.() || conversationId,
        graphChatId: conversation.graphChatId,
        topic: conversationRecord?.topic || conversation.topic || '',
        chatType: conversationRecord?.chatType || conversation.chatType || '',
        messageCount: conversationRecord?.messageCount || messageCountForConversation || 0,
        lastMessageAt: conversationRecord?.lastMessageAt || latestMessageAt || conversation.lastMessageAt,
        syncStatus:
          conversationRecord?.syncStatus ||
          (conversationFailed ? 'failed' : isConversationComplete ? 'complete' : 'pending'),
      });

      processedChatTypeSummary[chatType] = (processedChatTypeSummary[chatType] || 0) + 1;
      processedChats += 1;

      const shouldHeartbeat =
        processedChats % HEARTBEAT_CHAT_INTERVAL === 0 ||
        Date.now() - lastHeartbeatAt >= HEARTBEAT_MIN_INTERVAL_MS;

      if (shouldHeartbeat) {
        await heartbeatSyncExecution(
          syncJob._id?.toString?.() || syncJob.id,
          {
            phase: 'syncing_messages',
            conversationCount: processedChats,
            messageCount: persistedMessages,
            processedChatCount: processedChats,
            skippedChatCount: skippedMessageChats,
            checkpoint: buildSyncJobCheckpoint({
              phase: 'syncing_messages',
              nextChatPageLink,
              discoveryComplete,
              pageNumber,
              discoveredThisRun,
              processedChats,
              persistedMessages,
            }),
            stats: {
              graphChatTypeSummary,
              processedChatTypeSummary,
            },
          },
          { userLeaseKey, slotLeaseKey, ownerToken },
        );
        lastHeartbeatAt = Date.now();
        await refreshBackfillStateSnapshot(userId, {
          nextChatPageLink,
          discoveryComplete,
          lastSyncJobId: syncJob._id?.toString?.() || syncJob.id,
        }, { includeMessageCount: false, hasActiveSync: true });
      }
    }

    const backfillSnapshot = await refreshBackfillStateSnapshot(userId, {
      nextChatPageLink,
      discoveryComplete,
      lastSyncJobId: syncJob._id?.toString?.() || syncJob.id,
      lastCompletedAt: discoveryComplete ? new Date() : undefined,
      errorMessage: '',
    }, { hasActiveSync: false });

    logger.info('[TeamsArchiveService] Sync completed', {
      userId,
      syncJobId: syncJob._id?.toString?.() || syncJob.id,
      chatLimit,
      messagesPerChat,
      discoveredThisRun,
      processedChats,
      persistedMessages,
      skippedMessageChats,
      discoveryComplete,
      nextChatPageLinkPresent: Boolean(nextChatPageLink),
      graphChatTypeSummary,
      discoveryDecisionSummary,
      processedChatTypeSummary,
      memberLookup: memberLookupController.stats,
    });

    const updatedJob = await db.updateTeamsArchiveSyncJob(syncJob._id?.toString?.() || syncJob.id, {
      status: 'success',
      phase: 'complete',
      checkpoint: buildSyncJobCheckpoint({
        phase: 'complete',
        nextChatPageLink,
        discoveryComplete,
        pageNumber,
        discoveredThisRun,
        processedChats,
        persistedMessages,
      }),
      stats: {
        graphChatTypeSummary,
        discoveryDecisionSummary,
        processedChatTypeSummary,
        memberLookup: memberLookupController.stats,
        backfillState: {
          discoveredChatCount: backfillSnapshot?.discoveredChatCount || 0,
          completedChatCount: backfillSnapshot?.completedChatCount || 0,
          pendingChatCount: backfillSnapshot?.pendingChatCount || 0,
          runningChatCount: backfillSnapshot?.runningChatCount || 0,
          failedChatCount: backfillSnapshot?.failedChatCount || 0,
          totalMessageCount: backfillSnapshot?.totalMessageCount || 0,
        },
      },
      discoveredChatCount: discoveredThisRun,
      processedChatCount: processedChats,
      skippedChatCount: skippedMessageChats,
      conversationCount: processedChats,
      messageCount: persistedMessages,
      completedAt: new Date(),
    });

    const memoryProjection =
      completedGraphChatIds.size > 0
        ? queueTeamsProjection({
            userId,
            tenantId: user?.tenantId,
            syncJobId: updatedJob?._id?.toString?.() || syncJob._id?.toString?.() || syncJob.id,
            graphChatIds: [...completedGraphChatIds],
          }) || {
            status: 'skipped',
            reason: 'enterprise_memory_projection_unavailable',
          }
        : {
            status: 'skipped',
            reason: 'no_completed_conversations_in_run',
          };

    return {
      syncJob: updatedJob || syncJob,
      mode,
      conversationCount: processedChats,
      messageCount: backfillSnapshot?.totalMessageCount || persistedMessages,
      discovery: {
        discoveredThisRun,
        discoveryComplete,
        nextChatPageLinkPresent: Boolean(nextChatPageLink),
      },
      skippedMessageChats,
      conversations: syncedConversations,
      memoryProjection,
    };
  } catch (error) {
    if (!syncJob) {
      throw error;
    }

    if (error instanceof TeamsArchiveSyncCancelledError) {
      const cancelledJob =
        (await getSyncJobById(syncJob._id?.toString?.() || syncJob.id)) ||
        (await db.updateTeamsArchiveSyncJob(syncJob._id?.toString?.() || syncJob.id, {
          status: 'cancelled',
          errorMessage: 'Sync cancelled by user',
          completedAt: new Date(),
        }));

      return {
        syncJob: cancelledJob || syncJob,
        mode,
        conversationCount: cancelledJob?.conversationCount || 0,
        messageCount: cancelledJob?.messageCount || 0,
        conversations: [],
        memoryProjection: null,
        cancelled: true,
      };
    }

    await db.updateTeamsArchiveSyncJob(syncJob._id?.toString?.() || syncJob.id, {
      status: 'failure',
      phase: 'failed',
      errorMessage: error?.message || 'Teams archive sync failed',
      completedAt: new Date(),
    });
    await refreshBackfillStateSnapshot(userId, {
      status: 'failed',
      lastSyncJobId: syncJob._id?.toString?.() || syncJob.id,
      errorMessage: error?.message || 'Teams archive sync failed',
    }, { includeMessageCount: false });
    throw error;
  } finally {
    const releaseOperations = [releaseSyncLease(userLeaseKey, ownerToken)];

    if (slotLeaseKey) {
      releaseOperations.push(releaseSyncLease(slotLeaseKey, ownerToken));
    }

    await Promise.allSettled(releaseOperations);
  }
}

async function listConversations(user, options = {}) {
  assertEnabled();
  const userId = user?.id || user?._id?.toString();
  const limit = clampInteger(options.limit, 3, { max: 5 });
  const offset = clampInteger(options.offset, 0, { min: 0, max: 100000 });
  const { normalizedChatType, topic, daysBack, participantClauses, filter: conversationFilter } =
    buildConversationLookupFilter(userId, options);

  const conversations = await db.findTeamsArchiveConversations(
    conversationFilter,
    { limit, offset, sort: { lastMessageAt: -1, updatedAt: -1 } },
  );

  return {
    retrievalMode: 'conversation_list',
    chatType: normalizedChatType,
    topic: topic || undefined,
    daysBack,
    participants: toArray(options.participants).filter(Boolean),
    guidance:
      participantClauses.length > 0 || normalizedChatType !== 'any' || topic
        ? 'These are scoped conversations. For one exact chat, prefer summarize_conversation first, then get_messages or get_messages_window if you need message-level detail.'
        : 'These are archived conversations. Narrow to one exact chat before expanding messages.',
    conversations: conversations.map((conversation) => mapCompactConversation(conversation)),
  };
}

async function getConversationDossier(user, options = {}) {
  assertEnabled();
  const userId = user?.id || user?._id?.toString();
  const chatId = String(options.chatId || '').trim();
  const query = String(options.query || options.topic || '').trim();
  const queryRegex = query ? buildSearchRegex(query) : null;
  const previewLimit = clampInteger(options.limit, 4, { min: 1, max: 6 });
  const scopedDaysBack = options.daysBack
    ? clampInteger(options.daysBack, 30, { min: 1, max: 3650 })
    : undefined;

  let conversation = null;
  let candidates = [];

  if (chatId) {
    const resolvedGraphChatId = await resolveConversationGraphChatId(userId, chatId);
    if (!resolvedGraphChatId) {
      return {
        retrievalMode: 'conversation_dossier',
        resolved: false,
        chatId,
        graphChatId: null,
        guidance: 'No archived conversation was found for the requested chat id.',
        candidates: [],
      };
    }

    const matches = await db.findTeamsArchiveConversations(
      { user: userId, graphChatId: resolvedGraphChatId },
      { limit: 1 },
    );
    conversation = matches[0] || null;
  } else {
    const lookup = await findConversationCandidates(userId, {
      ...options,
      limit: 10,
    });
    candidates = lookup.conversations;

    if (candidates.length === 0) {
      return {
        retrievalMode: 'conversation_dossier',
        resolved: false,
        chatType: lookup.normalizedChatType,
        topic: lookup.topic || undefined,
        daysBack: lookup.daysBack,
        participants: toArray(options.participants).filter(Boolean),
        guidance:
          'No archived conversations matched these constraints. Broaden the participant, topic, or time filters.',
        candidates: [],
      };
    }

    if (candidates.length > 1) {
      return {
        retrievalMode: 'conversation_dossier',
        resolved: false,
        chatType: lookup.normalizedChatType,
        topic: lookup.topic || undefined,
        daysBack: lookup.daysBack,
        participants: toArray(options.participants).filter(Boolean),
        candidateCount: candidates.length,
        guidance:
          'Multiple archived conversations matched. Pick one chatId from the candidates before asking for exhaustive retrieval.',
        candidates: candidates.slice(0, 10).map((candidate) => mapConversationCandidate(candidate)),
      };
    }

    conversation = candidates[0];
  }

  if (!conversation?.graphChatId) {
    return {
      retrievalMode: 'conversation_dossier',
      resolved: false,
      guidance: 'No archived conversation metadata was found for the requested constraints.',
      candidates: [],
    };
  }

  const messageFilter = {
    user: userId,
    graphChatId: conversation.graphChatId,
    ...(scopedDaysBack
      ? { sentDateTime: { $gte: new Date(Date.now() - scopedDaysBack * 24 * 60 * 60 * 1000) } }
      : {}),
  };

  const [totalMessagesInScope, messages] = await Promise.all([
    db.countTeamsArchiveMessages(messageFilter),
    db.findTeamsArchiveMessages(messageFilter, {
      limit: DEFAULT_CONVERSATION_DOSSIER_MAX_MESSAGES,
      sort: { sentDateTime: 1, createdAt: 1 },
    }),
  ]);

  const matchedMessages = queryRegex
    ? messages.filter((message) => messageMatchesRegex(message, queryRegex))
    : messages;
  const matchedPreviewSource = matchedMessages.length > 0 ? matchedMessages : messages;
  const firstMessageAt = messages[0]?.sentDateTime || null;
  const lastMessageAt = messages[messages.length - 1]?.sentDateTime || null;
  const firstMatchedAt = matchedMessages[0]?.sentDateTime || null;
  const lastMatchedAt = matchedMessages[matchedMessages.length - 1]?.sentDateTime || null;
  const monthlyCoverage = summarizeCoverageByMonth(messages, queryRegex);
  const loadedAllMessages = totalMessagesInScope <= messages.length;

  return {
    retrievalMode: 'conversation_dossier',
    resolved: true,
    archiveBacked: true,
    completeness: {
      loadedAllMessages,
      loadedMessages: messages.length,
      totalMessagesInScope,
      truncated: !loadedAllMessages,
      cap: DEFAULT_CONVERSATION_DOSSIER_MAX_MESSAGES,
    },
    guidance:
      'This dossier is built directly from the archived Teams messages for one resolved chat. Use it when completeness matters more than quick previews.',
    chat: mapConversationCandidate(conversation),
    query: query || undefined,
    daysBack: scopedDaysBack,
    firstMessageAt,
    lastMessageAt,
    matchedMessages: matchedMessages.length,
    firstMatchedAt,
    lastMatchedAt,
    topSenders: summarizeSenderCounts(messages),
    matchedTopSenders: summarizeSenderCounts(matchedMessages),
    monthlyCoverage,
    highlights: mapDossierHighlights(matchedPreviewSource.slice(0, previewLimit), conversation, previewLimit),
    oldestMatches: mapDossierHighlights(matchedMessages.slice(0, previewLimit), conversation, previewLimit),
    newestMatches: mapDossierHighlights(
      matchedMessages.slice(Math.max(0, matchedMessages.length - previewLimit)),
      conversation,
      previewLimit,
    ),
  };
}

async function resolveConversationGraphChatId(userId, chatId) {
  const normalizedChatId = String(chatId || '').trim();
  if (!normalizedChatId) {
    return null;
  }

  const directMatches = await db.findTeamsArchiveConversations(
    { user: userId, graphChatId: normalizedChatId },
    { limit: 1 },
  );

  if (directMatches[0]?.graphChatId) {
    return directMatches[0].graphChatId;
  }

  const archivedMatches = await db.findTeamsArchiveConversations(
    { user: userId, _id: normalizedChatId },
    { limit: 1 },
  );

  return archivedMatches[0]?.graphChatId || null;
}

async function resolveMessageRecord(userId, messageId) {
  const normalizedMessageId = String(messageId || '').trim();
  if (!normalizedMessageId) {
    return null;
  }

  const graphMatches = await db.findTeamsArchiveMessages(
    { user: userId, graphMessageId: normalizedMessageId },
    { limit: 1 },
  );

  if (graphMatches[0]) {
    return graphMatches[0];
  }

  const archivedMatches = await db.findTeamsArchiveMessages(
    { user: userId, _id: normalizedMessageId },
    { limit: 1 },
  );

  return archivedMatches[0] || null;
}

async function listConversationMessages(user, chatId, options = {}) {
  assertEnabled();
  const userId = user?.id || user?._id?.toString();
  if (!chatId) {
    throw new TeamsArchiveServiceError('Chat id is required', 400);
  }

  const limit = clampInteger(options.limit, 6, { max: 8 });
  const offset = clampInteger(options.offset, 0, { min: 0, max: 100000 });
  const resolvedGraphChatId = await resolveConversationGraphChatId(userId, chatId);

  if (!resolvedGraphChatId) {
    return {
      chatId,
      graphChatId: null,
      messages: [],
    };
  }

  const messages = await db.findTeamsArchiveMessages(
    { user: userId, graphChatId: resolvedGraphChatId },
    { limit, offset, sort: { sentDateTime: 1, createdAt: 1 } },
  );

  const conversations = await db.findTeamsArchiveConversations(
    { user: userId, graphChatId: resolvedGraphChatId },
    { limit: 1 },
  );
  const conversation = conversations[0] || null;

  return {
    retrievalMode: 'thread_previews',
    chatId,
    graphChatId: resolvedGraphChatId,
    topic: conversation?.topic || '',
    chatType: conversation?.chatType || '',
    participants: mapCompactParticipants(conversation?.participants || []),
    guidance:
      'These are compact thread previews. Use summarize_conversation for a high-level answer or get_messages_window for a bounded slice around the most relevant message.',
    messages: messages.map((message) =>
      mapCompactMessageResult(message, conversation, {
        includeReplyToId: true,
        includeImportance: true,
        excerptLimit: 320,
      }),
    ),
  };
}

async function getMessageBody(user, options = {}) {
  assertEnabled();
  const userId = user?.id || user?._id?.toString();
  const messageId = String(options.messageId || options.aroundMessageId || '').trim();
  const chatId = String(options.chatId || '').trim();

  if (!messageId) {
    throw new TeamsArchiveServiceError('Message id is required', 400);
  }

  const message = await resolveMessageRecord(userId, messageId);

  if (!message) {
    return {
      retrievalMode: 'message_body',
      resolved: false,
      messageId,
      graphMessageId: null,
      guidance: 'No archived message was found for the requested message id.',
    };
  }

  if (chatId) {
    const resolvedGraphChatId = await resolveConversationGraphChatId(userId, chatId);
    if (resolvedGraphChatId && message.graphChatId !== resolvedGraphChatId) {
      return {
        retrievalMode: 'message_body',
        resolved: false,
        messageId,
        graphMessageId: message.graphMessageId,
        graphChatId: message.graphChatId,
        guidance:
          'The archived message was found, but it does not belong to the requested chat id scope.',
      };
    }
  }

  const conversations = await db.findTeamsArchiveConversations(
    { user: userId, graphChatId: message.graphChatId },
    { limit: 1 },
  );
  const conversation = conversations[0] || null;
  const fullText = String(message.bodyText || '');
  const previewText = String(message.bodyPreview || '');

  return {
    retrievalMode: 'message_body',
    resolved: true,
    guidance:
      'This returns the full archived message text for one specific message. Use it when a preview was truncated and exact wording matters.',
    message: {
      id: message._id?.toString?.() || message.id,
      graphMessageId: message.graphMessageId,
      graphChatId: message.graphChatId,
      topic: conversation?.topic || '',
      chatType: conversation?.chatType || '',
      participants: mapCompactParticipants(conversation?.participants || []),
      fromDisplayName: message.fromDisplayName || '',
      fromEmail: message.fromEmail || '',
      subject: message.subject || '',
      summary: message.summary || '',
      bodyContentType: message.bodyContentType || '',
      bodyPreview: previewText,
      bodyText: fullText,
      previewLength: previewText.length,
      bodyTextLength: fullText.length,
      previewWasTruncated: Boolean(fullText && previewText && fullText !== previewText),
      attachments: toArray(message.attachments).map((attachment) => ({
        id: attachment?.id || '',
        name: attachment?.name || '',
        contentType: attachment?.contentType || '',
        contentUrl: attachment?.contentUrl || '',
      })),
      mentions: toArray(message.mentions).map((mention) => ({
        id: mention?.id || '',
        displayName: mention?.displayName || '',
        mentionedUserId: mention?.mentionedUserId || '',
      })),
      sentDateTime: message.sentDateTime,
      webUrl: message.webUrl || '',
    },
  };
}

async function getMessagesWindow(user, options = {}) {
  assertEnabled();
  const userId = user?.id || user?._id?.toString();
  const chatId = String(options.chatId || '').trim();
  if (!chatId) {
    throw new TeamsArchiveServiceError('Chat id is required', 400);
  }

  const resolvedGraphChatId = await resolveConversationGraphChatId(userId, chatId);
  if (!resolvedGraphChatId) {
    return {
      chatId,
      graphChatId: null,
      anchorMessageId: null,
      messages: [],
    };
  }

  const before = clampInteger(options.before, 3, { min: 0, max: 10 });
  const after = clampInteger(options.after, 3, { min: 0, max: 10 });
  const fallbackLimit = clampInteger(options.limit, before + after + 1, { min: 1, max: 100 });
  const query = String(options.query || '').trim();
  const queryRegex = query ? buildSearchRegex(query) : null;

  let anchorMessage = options.aroundMessageId
    ? await resolveMessageRecord(userId, options.aroundMessageId)
    : null;

  if (!anchorMessage && queryRegex) {
    const matchedMessages = await db.findTeamsArchiveMessages(
      {
        user: userId,
        graphChatId: resolvedGraphChatId,
        $or: SEARCHABLE_MESSAGE_FIELDS.map((field) => ({ [field]: queryRegex })),
      },
      { limit: 1, sort: { sentDateTime: -1, createdAt: -1 } },
    );
    anchorMessage = matchedMessages[0] || null;
  }

  const conversations = await db.findTeamsArchiveConversations(
    { user: userId, graphChatId: resolvedGraphChatId },
    { limit: 1 },
  );
  const conversation = conversations[0] || null;

  if (!anchorMessage) {
    const recentMessages = await db.findTeamsArchiveMessages(
      { user: userId, graphChatId: resolvedGraphChatId },
      { limit: fallbackLimit, sort: { sentDateTime: -1, createdAt: -1 } },
    );

    return {
      retrievalMode: 'message_window',
      chatId,
      graphChatId: resolvedGraphChatId,
      topic: conversation?.topic || '',
      chatType: conversation?.chatType || '',
      participants: mapCompactParticipants(conversation?.participants || []),
      anchorMessageId: null,
      anchorGraphMessageId: null,
      messages: recentMessages.reverse().map((message) =>
        mapCompactMessageResult(message, conversation, {
          includeReplyToId: true,
          excerptLimit: 360,
        }),
      ),
    };
  }

  const anchorTime = anchorMessage.sentDateTime || anchorMessage.createdAt || new Date();
  const [beforeMessages, afterMessages] = await Promise.all([
    db.findTeamsArchiveMessages(
      {
        user: userId,
        graphChatId: resolvedGraphChatId,
        sentDateTime: { $lte: anchorTime },
      },
      { limit: before + 1, sort: { sentDateTime: -1, createdAt: -1 } },
    ),
    db.findTeamsArchiveMessages(
      {
        user: userId,
        graphChatId: resolvedGraphChatId,
        sentDateTime: { $gte: anchorTime },
      },
      { limit: after + 1, sort: { sentDateTime: 1, createdAt: 1 } },
    ),
  ]);

  const messageMap = new Map();
  for (const message of [...beforeMessages.reverse(), ...afterMessages]) {
    const key = message._id?.toString?.() || message.id || message.graphMessageId;
    messageMap.set(key, message);
  }

  const messages = [...messageMap.values()]
    .sort((a, b) => {
      const aTime = a.sentDateTime?.getTime?.() ?? 0;
      const bTime = b.sentDateTime?.getTime?.() ?? 0;
      return aTime - bTime;
    })
    .slice(-(before + after + 1));

  return {
    retrievalMode: 'message_window',
    chatId,
    graphChatId: resolvedGraphChatId,
    topic: conversation?.topic || '',
    chatType: conversation?.chatType || '',
    participants: mapCompactParticipants(conversation?.participants || []),
    anchorMessageId: anchorMessage._id?.toString?.() || anchorMessage.id,
    anchorGraphMessageId: anchorMessage.graphMessageId,
    query: query || undefined,
    messages: messages.map((message) =>
      mapCompactMessageResult(message, conversation, {
        includeReplyToId: true,
        excerptLimit: 360,
      }),
    ),
  };
}

async function searchMessages(user, options = {}) {
  assertEnabled();
  const userId = user?.id || user?._id?.toString();
  const query = String(options.query || '').trim();
  if (!query) {
    throw new TeamsArchiveServiceError('Search query is required', 400);
  }

  const limit = clampInteger(options.limit, Math.min(getTeamsArchiveConfig().defaultSearchLimit, 4), {
    max: 6,
  });
  const offset = clampInteger(options.offset, 0, { min: 0, max: 100000 });
  const chatId = options.chatId ? String(options.chatId) : undefined;
  const resolvedGraphChatId = chatId ? await resolveConversationGraphChatId(userId, chatId) : undefined;
  const regex = buildSearchRegex(query);
  const matchedConversations = await db.findTeamsArchiveConversations(
    {
      user: userId,
      ...(resolvedGraphChatId ? { graphChatId: resolvedGraphChatId } : {}),
      topic: regex,
    },
    { limit: 500 },
  );
  const matchedConversationIds = matchedConversations
    .map((conversation) => conversation.graphChatId)
    .filter(Boolean);
  const filter = {
    user: userId,
    ...(resolvedGraphChatId ? { graphChatId: resolvedGraphChatId } : {}),
    $or: [
      { bodyText: regex },
      { bodyPreview: regex },
      { bodyContent: regex },
      { summary: regex },
      { subject: regex },
      { fromDisplayName: regex },
      { fromEmail: regex },
      { 'attachments.name': regex },
      { 'mentions.displayName': regex },
      ...(matchedConversationIds.length > 0 ? [{ graphChatId: { $in: matchedConversationIds } }] : []),
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
    retrievalMode: 'message_previews',
    query,
    chatId: chatId || undefined,
    graphChatId: resolvedGraphChatId || undefined,
    guidance:
      'These are compact message previews. If one chat is clearly relevant, prefer summarize_conversation first, then get_messages_window if you need local context.',
    resultCount: messages.length,
    results: messages.map((message) =>
      mapCompactMessageResult(message, conversationMap.get(message.graphChatId), {
        excerptLimit: 260,
      }),
    ),
  };
}

async function recentMessages(user, options = {}) {
  assertEnabled();
  if (typeof searchTeamsMemoryChunks === 'function') {
    try {
      const memoryResults = await searchTeamsMemoryChunks(user, {
        query: options.query,
        limit: options.limit,
        daysBack: options.daysBack,
        senderScope: 'me',
        sortBy: 'recent',
      });

      if (hasNonEmptyMemoryResults(memoryResults)) {
        return memoryResults;
      }
    } catch (error) {
      logger.warn('[TeamsArchiveService] Enterprise memory recent retrieval failed, falling back', {
        userId: user?.id || user?._id?.toString?.(),
        error: error?.message || error,
      });
    }
  }

  const userId = user?.id || user?._id?.toString();
  const limit = clampInteger(options.limit, 4, { max: 6 });
  const daysBack = clampInteger(options.daysBack, 14, { min: 1, max: 3650 });
  const query = String(options.query || '').trim();
  const sentAfter = new Date(Date.now() - daysBack * 24 * 60 * 60 * 1000);
  const identityFilter = getUserMessageIdentityFilter(user, userId);
  const queryRegex = query ? buildSearchRegex(query) : null;

  const filter = {
    ...identityFilter,
    sentDateTime: { $gte: sentAfter },
    ...(queryRegex
      ? {
          $and: [
            {
              $or: [
                { bodyText: queryRegex },
                { bodyPreview: queryRegex },
                { bodyContent: queryRegex },
                { summary: queryRegex },
                { subject: queryRegex },
                { 'attachments.name': queryRegex },
                { 'mentions.displayName': queryRegex },
              ],
            },
          ],
        }
      : {}),
  };

  const messages = await db.findTeamsArchiveMessages(filter, {
    limit,
    offset: 0,
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
    retrievalMode: 'recent_message_previews',
    daysBack,
    query: query || undefined,
    guidance:
      'These are compact previews of recent messages sent by the signed-in user. Use get_messages_window for local context around one result.',
    resultCount: messages.length,
    results: messages.map((message) =>
      mapCompactMessageResult(message, conversationMap.get(message.graphChatId), {
        excerptLimit: 220,
      }),
    ),
  };
}

async function advancedSearchMessages(user, options = {}) {
  assertEnabled();
  if (typeof searchTeamsMemoryChunks === 'function') {
    try {
      const memoryResults = await searchTeamsMemoryChunks(user, options);
      if (hasNonEmptyMemoryResults(memoryResults)) {
        return memoryResults;
      }
    } catch (error) {
      logger.warn('[TeamsArchiveService] Enterprise memory advanced retrieval failed, falling back', {
        userId: user?.id || user?._id?.toString?.(),
        error: error?.message || error,
      });
    }
  }

  const userId = user?.id || user?._id?.toString();
  const topic = String(options.topic || options.query || '').trim();
  const senderScope = String(options.senderScope || 'any').trim();
  const chatType = String(options.chatType || 'any').trim();
  const sortBy = String(options.sortBy || 'recent').trim();
  const limit = clampInteger(options.limit, Math.min(getTeamsArchiveConfig().defaultSearchLimit, 4), {
    max: 6,
  });
  const offset = clampInteger(options.offset, 0, { min: 0, max: 100000 });
  const daysBack = options.daysBack
    ? clampInteger(options.daysBack, 30, { min: 1, max: 3650 })
    : undefined;

  const validSenderScopes = new Set(['any', 'me', 'others']);
  const validChatTypes = new Set(['any', 'oneOnOne', 'group', 'meeting']);
  const normalizedSenderScope = validSenderScopes.has(senderScope) ? senderScope : 'any';
  const normalizedChatType = validChatTypes.has(chatType) ? chatType : 'any';

  const { phraseRegex, termRegexes, clauses: topicClauses } = buildTopicSearchClauses(topic);
  const participantClauses = buildParticipantConversationClauses(options.participants);
  let matchedConversationIds = [];
  let matchedConversations = [];
  if (normalizedChatType !== 'any' || topic || participantClauses.length > 0) {
    const lookup = await findConversationCandidates(userId, {
      ...options,
      topic,
      query: topic,
      chatType: normalizedChatType,
      candidateLimit: 1000,
    });
    matchedConversations = lookup.conversations;
    matchedConversationIds = matchedConversations
      .map((conversation) => conversation.graphChatId)
      .filter(Boolean);

    if (matchedConversationIds.length === 0) {
      return {
        retrievalMode: 'advanced_message_previews',
        topic: topic || undefined,
        senderScope: normalizedSenderScope,
        chatType: normalizedChatType,
        daysBack,
        participants: toArray(options.participants).filter(Boolean),
        guidance:
          'No matching chats were found. Consider broadening the topic, participants, or timeframe.',
        results: [],
      };
    }
  }

  const senderClauses = getUserSenderClauses(user);
  const senderFilter =
    normalizedSenderScope === 'me'
      ? senderClauses.length > 0
        ? { $or: senderClauses }
        : {}
      : normalizedSenderScope === 'others'
        ? senderClauses.length > 0
          ? { $nor: senderClauses }
          : {}
        : {};

  const messageFilter = {
    user: userId,
    ...(matchedConversationIds.length > 0 ? { graphChatId: { $in: matchedConversationIds } } : {}),
    ...(daysBack ? { sentDateTime: { $gte: new Date(Date.now() - daysBack * 24 * 60 * 60 * 1000) } } : {}),
    ...senderFilter,
    ...(topicClauses.length > 0 ? { $and: topicClauses } : {}),
  };

  const messages = await db.findTeamsArchiveMessages(messageFilter, {
    limit,
    offset,
    sort: sortBy === 'oldest' ? { sentDateTime: 1, createdAt: 1 } : { sentDateTime: -1, createdAt: -1 },
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
  const resolvedConversation =
    matchedConversations.length === 1 ? mapConversationCandidate(matchedConversations[0]) : undefined;

  return {
    retrievalMode: 'advanced_message_previews',
    topic: topic || undefined,
    senderScope: normalizedSenderScope,
    chatType: normalizedChatType,
    daysBack,
    participants: toArray(options.participants).filter(Boolean),
    guidance:
      resolvedConversation
        ? 'These previews are scoped to one resolved conversation. If completeness matters, use conversation_dossier or summarize_conversation before expanding with get_messages_window.'
        : 'These are compact previews optimized for topic discovery. If a single conversation stands out, use summarize_conversation before expanding with get_messages_window.',
    ...(resolvedConversation ? { resolvedConversation } : {}),
    resultCount: messages.length,
    results: messages.map((message) =>
      mapCompactMessageResult(message, conversationMap.get(message.graphChatId), {
        excerptLimit: 260,
      }),
    ),
  };
}

async function summarizeConversation(user, options = {}) {
  assertEnabled();
  const userId = user?.id || user?._id?.toString();
  const chatId = String(options.chatId || '').trim();
  if (!chatId) {
    throw new TeamsArchiveServiceError('Chat id is required', 400);
  }

  const resolvedGraphChatId = await resolveConversationGraphChatId(userId, chatId);
  if (!resolvedGraphChatId) {
    return {
      chatId,
      graphChatId: null,
      summary: 'No archived conversation was found for the requested chat.',
      highlights: [],
    };
  }

  const query = String(options.query || options.topic || '').trim();
  const queryRegex = query ? buildSearchRegex(query) : null;
  const daysBack = options.daysBack
    ? clampInteger(options.daysBack, 30, { min: 1, max: 3650 })
    : undefined;
  const highlightLimit = clampInteger(options.limit, 4, { min: 1, max: 6 });

  const [conversations, messages] = await Promise.all([
    db.findTeamsArchiveConversations({ user: userId, graphChatId: resolvedGraphChatId }, { limit: 1 }),
    db.findTeamsArchiveMessages(
      {
        user: userId,
        graphChatId: resolvedGraphChatId,
        ...(daysBack
          ? { sentDateTime: { $gte: new Date(Date.now() - daysBack * 24 * 60 * 60 * 1000) } }
          : {}),
      },
      { limit: 5000, sort: { sentDateTime: 1, createdAt: 1 } },
    ),
  ]);

  const conversation = conversations[0] || null;
  const scopedMessages = queryRegex
    ? messages.filter((message) => messageMatchesRegex(message, queryRegex))
    : messages;
  const highlightSource = scopedMessages.length > 0 ? scopedMessages : messages;
  const highlights = highlightSource.slice(-highlightLimit);
  const topSenders = summarizeSenderCounts(scopedMessages.length > 0 ? scopedMessages : messages);
  const firstMessageAt = messages[0]?.sentDateTime || null;
  const lastMessageAt = messages[messages.length - 1]?.sentDateTime || null;

  return {
    retrievalMode: 'conversation_summary',
    chatId,
    graphChatId: resolvedGraphChatId,
    topic: conversation?.topic || '',
    chatType: conversation?.chatType || '',
    participants: mapCompactParticipants(conversation?.participants || []),
    daysBack,
    query: query || undefined,
    totalMessages: messages.length,
    matchedMessages: scopedMessages.length,
    firstMessageAt,
    lastMessageAt,
    topSenders,
    summary: summarizeConversationText({
      topic: conversation?.topic,
      chatType: conversation?.chatType,
      totalMessages: messages.length,
      matchedMessages: scopedMessages.length,
      daysBack,
      topSenders,
    }),
    highlights: highlights.map((message) => ({
      id: message._id?.toString?.() || message.id,
      graphMessageId: message.graphMessageId,
      fromDisplayName: message.fromDisplayName || '',
      fromEmail: message.fromEmail || '',
      sentDateTime: message.sentDateTime,
      excerpt: message.bodyPreview || message.summary || message.bodyText?.slice(0, 500) || '',
      webUrl: message.webUrl || '',
    })),
  };
}

module.exports = {
  TeamsArchiveServiceError,
  TeamsArchiveSyncCancelledError,
  getStatus,
  deleteUserArchive,
  cancelRunningSync,
  getSyncStartAvailability,
  syncUserArchive,
  listConversations,
  getConversationDossier,
  listConversationMessages,
  getMessageBody,
  getMessagesWindow,
  searchMessages,
  recentMessages,
  advancedSearchMessages,
  summarizeConversation,
};
