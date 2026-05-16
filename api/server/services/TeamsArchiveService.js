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

async function listChatMembers(user, chatId, chatType) {
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
      chatType: chatType || 'unknown',
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
  await reconcileRunningSyncJob(userId);
  const [conversationCount, messageCount, latestSync, latestProjection, activeSyncs] = await Promise.all([
    userId ? db.countTeamsArchiveConversations({ user: userId }) : 0,
    userId ? db.countTeamsArchiveMessages({ user: userId }) : 0,
    userId ? db.findLatestTeamsArchiveSyncJob({ user: userId }) : null,
    typeof db.findLatestEnterpriseMemoryJob === 'function'
      ? db.findLatestEnterpriseMemoryJob({ user: userId, source: 'teams', jobType: 'projection' })
      : null,
    countActiveSyncSlots(),
  ]);

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
    latestProjection: latestProjection
      ? {
          id: latestProjection._id?.toString?.() || latestProjection.id,
          status: latestProjection.status,
          startedAt: latestProjection.startedAt,
          completedAt: latestProjection.completedAt,
          errorMessage: latestProjection.errorMessage,
          stats: latestProjection.stats || {},
        }
      : null,
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
      errorMessage: 'Sync cancelled by user',
      completedAt: new Date(),
    },
  );

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
  const chatLimit = clampInteger(options.chatLimit, config.defaultChatLimit, { max: 250 });
  const messagesPerChat = clampInteger(options.messagesPerChat, config.defaultMessagesPerChat, {
    max: 1000,
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

    syncJob = await db.createTeamsArchiveSyncJob({
      user: userId,
      status: 'running',
      mode,
      conversationCount: 0,
      messageCount: 0,
      startedAt: new Date(),
    });

    const syncedConversations = [];
    let nextLink = null;
    let processedChats = 0;
    let persistedMessages = 0;
    let skippedMessageChats = 0;
    let pageNumber = 0;
    let lastHeartbeatAt = Date.now();
    const graphChatTypeSummary = { oneOnOne: 0, group: 0, meeting: 0, unknown: 0 };
    const processedChatTypeSummary = { oneOnOne: 0, group: 0, meeting: 0, unknown: 0 };

    while (processedChats < chatLimit) {
      await ensureSyncJobActive(syncJob._id?.toString?.() || syncJob.id);
      pageNumber += 1;
      const response = await listChatsPage(user, {
        top: Math.min(chatLimit - processedChats, 50),
        nextLink,
      });
      const chats = toArray(response?.value).sort((a, b) => {
        const aTime = toDate(a?.lastUpdatedDateTime)?.getTime() ?? 0;
        const bTime = toDate(b?.lastUpdatedDateTime)?.getTime() ?? 0;
        return bTime - aTime;
      });
      if (chats.length === 0) {
        break;
      }

      const pageChatTypeSummary = summarizeChatTypes(chats);
      for (const [chatType, count] of Object.entries(pageChatTypeSummary)) {
        graphChatTypeSummary[chatType] = (graphChatTypeSummary[chatType] || 0) + count;
      }

      logger.info('[TeamsArchiveService] Sync page loaded', {
        userId,
        syncJobId: syncJob._id?.toString?.() || syncJob.id,
        pageNumber,
        chatsReturned: chats.length,
        processedChats,
        chatLimit,
        pageChatTypeSummary,
        hasNextPage: Boolean(response?.['@odata.nextLink']),
      });

      await heartbeatSyncExecution(
        syncJob._id?.toString?.() || syncJob.id,
        {
          conversationCount: processedChats,
          messageCount: persistedMessages,
        },
        { userLeaseKey, slotLeaseKey, ownerToken },
      );
      lastHeartbeatAt = Date.now();

      for (const chat of chats) {
        await ensureSyncJobActive(syncJob._id?.toString?.() || syncJob.id);
        if (processedChats >= chatLimit) {
          break;
        }

        const chatType = String(chat?.chatType || 'unknown');
        const members = await listChatMembers(user, chat.id, chatType);
        const normalizedConversation = normalizeConversation(chat, members);
        let messages = [];

        try {
          messages = await listChatMessages(user, chat.id, { top: messagesPerChat });
        } catch (error) {
          if (!isRecoverableChatMessageError(error)) {
            throw error;
          }

          skippedMessageChats += 1;
          logger.warn('[TeamsArchiveService] Failed to list chat messages; continuing sync', {
            userId,
            syncJobId: syncJob._id?.toString?.() || syncJob.id,
            chatId: chat.id,
            chatType,
            status: error?.status,
            details: error?.details,
          });
        }

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

        processedChatTypeSummary[chatType] = (processedChatTypeSummary[chatType] || 0) + 1;
        processedChats += 1;

        const shouldHeartbeat =
          processedChats % HEARTBEAT_CHAT_INTERVAL === 0 ||
          Date.now() - lastHeartbeatAt >= HEARTBEAT_MIN_INTERVAL_MS;

        if (shouldHeartbeat) {
          await heartbeatSyncExecution(
            syncJob._id?.toString?.() || syncJob.id,
            {
              conversationCount: processedChats,
              messageCount: persistedMessages,
            },
            { userLeaseKey, slotLeaseKey, ownerToken },
          );
          lastHeartbeatAt = Date.now();
        }
      }

      nextLink =
        response?.['@odata.nextLink'] && processedChats < chatLimit ? response['@odata.nextLink'] : null;

      if (!nextLink) {
        break;
      }
    }

    logger.info('[TeamsArchiveService] Sync completed', {
      userId,
      syncJobId: syncJob._id?.toString?.() || syncJob.id,
      chatLimit,
      messagesPerChat,
      processedChats,
      persistedMessages,
      skippedMessageChats,
      graphChatTypeSummary,
      processedChatTypeSummary,
    });

    const updatedJob = await db.updateTeamsArchiveSyncJob(syncJob._id?.toString?.() || syncJob.id, {
      status: 'success',
      conversationCount: processedChats,
      messageCount: persistedMessages,
      completedAt: new Date(),
    });

    let memoryProjection = null;
    if (typeof projectTeamsArchiveSyncToMemory === 'function') {
      try {
        memoryProjection = await projectTeamsArchiveSyncToMemory({
          userId,
          tenantId: user?.tenantId,
          syncJobId: updatedJob?._id?.toString?.() || syncJob._id?.toString?.() || syncJob.id,
          graphChatIds: syncedConversations.map((conversation) => conversation.graphChatId),
        });
      } catch (projectionError) {
        logger.error('[TeamsArchiveService] Teams enterprise memory projection failed', {
          userId,
          syncJobId: updatedJob?._id?.toString?.() || syncJob._id?.toString?.() || syncJob.id,
          error: projectionError?.message || projectionError,
        });

        memoryProjection = {
          status: 'failure',
          errorMessage: projectionError?.message || 'Teams enterprise memory projection failed',
        };
      }
    } else {
      memoryProjection = {
        status: 'skipped',
        reason: 'enterprise_memory_projection_unavailable',
      };
    }

    return {
      syncJob: updatedJob || syncJob,
      mode,
      conversationCount: syncedConversations.length,
      messageCount: persistedMessages,
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
      errorMessage: error?.message || 'Teams archive sync failed',
      completedAt: new Date(),
    });
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

  const limit = clampInteger(options.limit, 100, { max: 500 });
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

  return {
    chatId,
    graphChatId: resolvedGraphChatId,
    messages: messages.map((message) => ({
      ...mapMessageResult(message),
      replyToId: message.replyToId || '',
      importance: message.importance || '',
      bodyContentType: message.bodyContentType || 'html',
      bodyContent: message.bodyContent || '',
      lastModifiedDateTime: message.lastModifiedDateTime,
    })),
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

  const before = clampInteger(options.before, 5, { min: 0, max: 50 });
  const after = clampInteger(options.after, 5, { min: 0, max: 50 });
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
      chatId,
      graphChatId: resolvedGraphChatId,
      topic: conversation?.topic || '',
      chatType: conversation?.chatType || '',
      participants: conversation?.participants || [],
      anchorMessageId: null,
      anchorGraphMessageId: null,
      messages: recentMessages.reverse().map((message) => mapMessageResult(message, conversation)),
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
    chatId,
    graphChatId: resolvedGraphChatId,
    topic: conversation?.topic || '',
    chatType: conversation?.chatType || '',
    participants: conversation?.participants || [],
    anchorMessageId: anchorMessage._id?.toString?.() || anchorMessage.id,
    anchorGraphMessageId: anchorMessage.graphMessageId,
    query: query || undefined,
    messages: messages.map((message) => mapMessageResult(message, conversation)),
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
    query,
    chatId: chatId || undefined,
    graphChatId: resolvedGraphChatId || undefined,
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
        subject: message.subject || '',
        summary: message.summary || '',
        bodyPreview: message.bodyPreview || '',
        bodyText: message.bodyText || '',
        attachments: message.attachments || [],
        sentDateTime: message.sentDateTime,
        webUrl: message.webUrl || '',
      };
    }),
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

      if (memoryResults) {
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
  const limit = clampInteger(options.limit, 20, { max: 100 });
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
    daysBack,
    query: query || undefined,
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
        subject: message.subject || '',
        summary: message.summary || '',
        bodyPreview: message.bodyPreview || '',
        bodyText: message.bodyText || '',
        sentDateTime: message.sentDateTime,
        webUrl: message.webUrl || '',
      };
    }),
  };
}

async function advancedSearchMessages(user, options = {}) {
  assertEnabled();
  if (typeof searchTeamsMemoryChunks === 'function') {
    try {
      const memoryResults = await searchTeamsMemoryChunks(user, options);
      if (memoryResults) {
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
  const limit = clampInteger(options.limit, getTeamsArchiveConfig().defaultSearchLimit, {
    max: 100,
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
  const conversationFilter = {
    user: userId,
    ...(normalizedChatType !== 'any' ? { chatType: normalizedChatType } : {}),
    ...((phraseRegex || termRegexes.length > 0 || participantClauses.length > 0)
      ? {
          $and: [
            ...(topic ? [buildFieldOrClause(['topic'], phraseRegex || termRegexes[0])] : []),
            ...participantClauses,
          ],
        }
      : {}),
  };

  let matchedConversationIds = [];
  if (normalizedChatType !== 'any' || topic || participantClauses.length > 0) {
    const matchedConversations = await db.findTeamsArchiveConversations(conversationFilter, {
      limit: 1000,
    });
    matchedConversationIds = matchedConversations
      .map((conversation) => conversation.graphChatId)
      .filter(Boolean);

    if (matchedConversationIds.length === 0) {
      return {
        topic: topic || undefined,
        senderScope: normalizedSenderScope,
        chatType: normalizedChatType,
        daysBack,
        participants: toArray(options.participants).filter(Boolean),
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

  return {
    topic: topic || undefined,
    senderScope: normalizedSenderScope,
    chatType: normalizedChatType,
    daysBack,
    participants: toArray(options.participants).filter(Boolean),
    results: messages.map((message) => {
      const conversation = conversationMap.get(message.graphChatId);
      return mapMessageResult(message, conversation);
    }),
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
  const highlightLimit = clampInteger(options.limit, 6, { min: 1, max: 12 });

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
    participants: conversation?.participants || [],
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
  cancelRunningSync,
  getSyncStartAvailability,
  syncUserArchive,
  listConversations,
  listConversationMessages,
  getMessagesWindow,
  searchMessages,
  recentMessages,
  advancedSearchMessages,
  summarizeConversation,
};
