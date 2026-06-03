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
const DEFAULT_RECENCY_BACKFILL_MAX_MESSAGES = 100000;
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
  const namedEntities = {
    nbsp: ' ',
    amp: '&',
    lt: '<',
    gt: '>',
    quot: '"',
    apos: "'",
    copy: '©',
    reg: '®',
    trade: '™',
    mdash: '-',
    ndash: '-',
    hellip: '...',
    bull: '*',
    middot: '·',
    laquo: '<<',
    raquo: '>>',
  };

  return String(value || '')
    .replace(/&#(\d+);/g, (_match, numeric) => {
      const codePoint = Number(numeric);
      return Number.isFinite(codePoint) ? String.fromCodePoint(codePoint) : _match;
    })
    .replace(/&#x([0-9a-f]+);/gi, (_match, hex) => {
      const codePoint = Number.parseInt(hex, 16);
      return Number.isFinite(codePoint) ? String.fromCodePoint(codePoint) : _match;
    })
    .replace(/&([a-z][a-z0-9]+);/gi, (match, entity) => namedEntities[entity.toLowerCase()] ?? match);
}

function normalizeHtmlText(content, contentType = 'text') {
  const raw = String(content || '');
  if (String(contentType || '').toLowerCase() !== 'html') {
    return raw.trim();
  }

  const withLinkTargets = raw.replace(
    /<a\b[^>]*href\s*=\s*["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi,
    (_match, href, linkText) => {
      const normalizedLinkText = String(linkText || '').replace(/<[^>]+>/g, ' ').trim();
      if (!normalizedLinkText) {
        return ` ${href} `;
      }
      return ` ${normalizedLinkText} (${href}) `;
    },
  );

  return decodeHtmlEntities(
    withLinkTargets
      .replace(/<style[\s\S]*?<\/style>/gi, ' ')
      .replace(/<script[\s\S]*?<\/script>/gi, ' ')
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<\/?(p|div|section|article|header|footer|aside|blockquote|h[1-6])[^>]*>/gi, '\n')
      .replace(/<li[^>]*>/gi, '\n- ')
      .replace(/<\/li>/gi, '\n')
      .replace(/<tr[^>]*>/gi, '\n')
      .replace(/<\/tr>/gi, '\n')
      .replace(/<(td|th)[^>]*>/gi, ' | ')
      .replace(/<\/(td|th)>/gi, ' ')
      .replace(/<\/(ul|ol|table)>/gi, '\n')
      .replace(/<[^>]+>/g, ' ')
      .replace(/\r/g, '\n')
      .replace(/[ \t]+\|/g, ' |')
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

function looksLikeMongoObjectId(value) {
  return /^[a-f0-9]{24}$/i.test(String(value || '').trim());
}

function buildParticipantIdentityKey(participant = {}) {
  const userId = String(participant?.userId || participant?.mentionedUserId || '')
    .trim()
    .toLowerCase();
  const email = String(participant?.email || '')
    .trim()
    .toLowerCase();
  const displayName = String(participant?.displayName || '')
    .trim()
    .toLowerCase();

  if (userId) {
    return `user:${userId}`;
  }

  if (email) {
    return `email:${email}`;
  }

  if (displayName) {
    return `name:${displayName}`;
  }

  return '';
}

function collapseParticipantSources(sourceSet = new Set()) {
  const normalizedSources = [...sourceSet].filter(Boolean);
  if (normalizedSources.length === 0) {
    return 'unknown';
  }

  if (normalizedSources.length === 1) {
    return normalizedSources[0];
  }

  return 'mixed';
}

function deriveParticipantConfidence({ userId, email, source }) {
  if (source === 'graph' || source === 'mixed' || String(userId || '').trim()) {
    return 'high';
  }

  if (String(email || '').trim()) {
    return 'medium';
  }

  return 'low';
}

function detectSystemLikeMessage(message, { bodyText, fromUser, fromDisplayName } = {}) {
  const messageType = String(message?.messageType || '').trim().toLowerCase();
  const hasStructuredSender = Boolean(String(fromUser?.id || '').trim() || String(fromDisplayName || '').trim());
  const hasMeaningfulText = Boolean(String(bodyText || message?.summary || message?.subject || '').trim());

  if (message?.deletedDateTime) {
    return true;
  }

  if (messageType && messageType !== 'message' && !hasMeaningfulText) {
    return true;
  }

  if (!hasStructuredSender && !hasMeaningfulText) {
    return true;
  }

  return false;
}

function classifyChunkability(message, { bodyText, normalizedTextLength, isSystemLikeMessage }) {
  const candidateText = String(bodyText || message?.summary || message?.subject || '').trim();

  if (message?.deletedDateTime) {
    return { isChunkable: false, skipChunkReason: 'deleted_message' };
  }

  if (!candidateText && isSystemLikeMessage) {
    return { isChunkable: false, skipChunkReason: 'system_like_message' };
  }

  if (!candidateText) {
    return { isChunkable: false, skipChunkReason: 'empty_normalized_text' };
  }

  if (isSystemLikeMessage && normalizedTextLength < 8) {
    return { isChunkable: false, skipChunkReason: 'system_like_message' };
  }

  return { isChunkable: true, skipChunkReason: '' };
}

function upsertParticipantRecord(targetMap, participant = {}, source = 'unknown') {
  const key = buildParticipantIdentityKey(participant);
  if (!key) {
    return;
  }

  const existing = targetMap.get(key) || {
    displayName: '',
    email: '',
    userId: '',
    sourceSet: new Set(),
  };

  existing.displayName = existing.displayName || String(participant?.displayName || '').trim();
  existing.email = existing.email || String(participant?.email || '').trim();
  existing.userId =
    existing.userId ||
    String(participant?.userId || participant?.mentionedUserId || '').trim();
  existing.sourceSet.add(source || 'unknown');
  targetMap.set(key, existing);
}

function finalizeParticipantRecords(targetMap, { memberLookupFailed = false } = {}) {
  const participants = [...targetMap.values()].map((participant) => {
    const source = collapseParticipantSources(participant.sourceSet);
    return {
      displayName: participant.displayName,
      email: participant.email,
      userId: participant.userId,
      source,
      confidence: deriveParticipantConfidence({
        userId: participant.userId,
        email: participant.email,
        source,
      }),
    };
  });

  const stats = participants.reduce(
    (acc, participant) => {
      acc.totalCount += 1;
      if (participant.source === 'graph') {
        acc.graphCount += 1;
      } else if (participant.source === 'inferred_from_messages') {
        acc.inferredFromMessagesCount += 1;
      } else if (participant.source === 'inferred_from_mentions') {
        acc.inferredFromMentionsCount += 1;
      } else if (participant.source === 'mixed') {
        acc.mixedCount += 1;
      } else {
        acc.unknownCount += 1;
      }
      return acc;
    },
    {
      totalCount: 0,
      graphCount: 0,
      inferredFromMessagesCount: 0,
      inferredFromMentionsCount: 0,
      mixedCount: 0,
      unknownCount: 0,
      memberLookupFailed,
    },
  );

  const sourceSet = new Set(participants.map((participant) => participant.source).filter(Boolean));

  return {
    participants,
    participantMetadataSource: collapseParticipantSources(sourceSet),
    participantConfidence: participants.some((participant) => participant.confidence === 'high')
      ? 'high'
      : participants.some((participant) => participant.confidence === 'medium')
        ? 'medium'
        : 'low',
    participantDegraded: Boolean(memberLookupFailed || (stats.graphCount === 0 && stats.totalCount > 0)),
    participantStats: stats,
  };
}

function mergeConversationParticipants({
  graphParticipants = [],
  existingParticipants = [],
  inferredMessageParticipants = [],
  inferredMentionParticipants = [],
  memberLookupFailed = false,
} = {}) {
  const participantMap = new Map();

  for (const participant of toArray(graphParticipants)) {
    upsertParticipantRecord(participantMap, participant, 'graph');
  }

  for (const participant of toArray(existingParticipants)) {
    upsertParticipantRecord(participantMap, participant, participant?.source || 'unknown');
  }

  for (const participant of toArray(inferredMessageParticipants)) {
    upsertParticipantRecord(participantMap, participant, 'inferred_from_messages');
  }

  for (const participant of toArray(inferredMentionParticipants)) {
    upsertParticipantRecord(participantMap, participant, 'inferred_from_mentions');
  }

  return finalizeParticipantRecords(participantMap, { memberLookupFailed });
}

function extractInferredParticipants(messages = []) {
  const messageParticipants = [];
  const mentionParticipants = [];

  for (const message of messages) {
    if (message?.fromDisplayName || message?.fromEmail || message?.fromUserId) {
      messageParticipants.push({
        displayName: message?.fromDisplayName,
        email: message?.fromEmail,
        userId: message?.fromUserId,
      });
    }

    for (const mention of toArray(message?.mentions)) {
      mentionParticipants.push({
        displayName: mention?.displayName,
        userId: mention?.mentionedUserId,
      });
    }
  }

  return {
    inferredMessageParticipants: messageParticipants,
    inferredMentionParticipants: mentionParticipants,
  };
}

function uniqueParticipants(participants) {
  const participantMap = new Map();

  for (const participant of participants) {
    const key = buildParticipantIdentityKey(participant);
    if (!key) {
      continue;
    }

    const existing = participantMap.get(key) || {
      displayName: '',
      email: '',
      userId: '',
      sourceSet: new Set(),
      confidence: 'low',
    };

    existing.displayName = existing.displayName || String(participant?.displayName || '').trim();
    existing.email = existing.email || String(participant?.email || '').trim();
    existing.userId = existing.userId || String(participant?.userId || '').trim();
    existing.sourceSet.add(String(participant?.source || 'unknown').trim() || 'unknown');
    existing.confidence = participant?.confidence || existing.confidence;
    participantMap.set(key, existing);
  }

  return [...participantMap.values()].map((participant) => ({
    displayName: participant.displayName,
    email: participant.email,
    userId: participant.userId,
    source: collapseParticipantSources(participant.sourceSet),
    confidence: participant.confidence,
  }));
}

function normalizeConversation(chat, members = []) {
  const participants = uniqueParticipants(
    members
      .map((member) => ({
        displayName: member?.displayName || member?.email || '',
        email: member?.email || member?.userId || '',
        userId: member?.userId,
        source: 'graph',
        confidence: member?.userId || member?.email ? 'high' : 'medium',
      }))
      .filter((participant) => participant.displayName || participant.email || participant.userId),
  );

  return {
    graphChatId: chat.id,
    chatType: chat.chatType,
    topic: chat.topic || chat.subject || '',
    webUrl: chat.webUrl || '',
    participants,
    participantMetadataSource: participants.length > 0 ? 'graph' : 'unknown',
    participantConfidence: participants.length > 0 ? 'high' : 'low',
    participantDegraded: false,
    participantStats: {
      totalCount: participants.length,
      graphCount: participants.length,
      inferredFromMessagesCount: 0,
      inferredFromMentionsCount: 0,
      mixedCount: 0,
      unknownCount: participants.length === 0 ? 1 : 0,
      memberLookupFailed: false,
    },
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
  const normalizedTextLength = bodyText.length;
  const isSystemLikeMessage = detectSystemLikeMessage(message, {
    bodyText,
    fromUser,
    fromDisplayName: from.displayName,
  });
  const { isChunkable, skipChunkReason } = classifyChunkability(message, {
    bodyText,
    normalizedTextLength,
    isSystemLikeMessage,
  });

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
    normalizedTextLength,
    isSystemLikeMessage,
    isChunkable,
    skipChunkReason,
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
    lastHumanMessageAt: conversation.lastHumanMessageAt || null,
    lastMeaningfulMessageAt: conversation.lastMeaningfulMessageAt || null,
    lastSystemMessageAt: conversation.lastSystemMessageAt || null,
    lastSyncedAt: conversation.lastSyncedAt,
    sourceUpdatedAt: conversation.sourceUpdatedAt,
    messageCount: conversation.messageCount || 0,
    humanMessageCount: conversation.humanMessageCount || 0,
    systemMessageCount: conversation.systemMessageCount || 0,
    emptyMessageCount: conversation.emptyMessageCount || 0,
    meaningfulMessageCount: conversation.meaningfulMessageCount || 0,
    warnings: buildConversationWarningFlags(conversation),
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
    lastHumanMessageAt: conversation.lastHumanMessageAt || null,
    lastMeaningfulMessageAt: conversation.lastMeaningfulMessageAt || null,
    lastSystemMessageAt: conversation.lastSystemMessageAt || null,
    lastSyncedAt: conversation.lastSyncedAt || null,
    messageCount: conversation.messageCount || 0,
    humanMessageCount: conversation.humanMessageCount || 0,
    systemMessageCount: conversation.systemMessageCount || 0,
    emptyMessageCount: conversation.emptyMessageCount || 0,
    meaningfulMessageCount: conversation.meaningfulMessageCount || 0,
    participantMetadataDegraded: Boolean(conversation.participantDegraded),
    warnings: buildConversationWarningFlags(conversation),
    syncStatus: conversation.syncStatus || '',
    webUrl: conversation.webUrl || '',
  };
}

function isMeaningfulArchiveMessage(message = {}) {
  return (
    message.isSystemLikeMessage !== true &&
    message.isChunkable === true &&
    Number(message.normalizedTextLength || 0) > 0
  );
}

function isEmptyArchiveMessage(message = {}) {
  return (
    Number(message.normalizedTextLength || 0) <= 0 ||
    message.skipChunkReason === 'empty_normalized_text' ||
    (message.isChunkable === false && !message.isSystemLikeMessage)
  );
}

function buildConversationWarningFlags(conversation = {}) {
  const lastMeaningful = toDate(conversation.lastMeaningfulMessageAt);
  const lastRaw = toDate(conversation.lastMessageAt);
  const lastSystem = toDate(conversation.lastSystemMessageAt);
  const meaningfulCount = Number(conversation.meaningfulMessageCount || 0);
  const messageCount = Number(conversation.messageCount || 0);

  return {
    systemOnlyRecentActivity:
      Boolean(lastSystem && lastRaw && lastSystem.getTime() === lastRaw.getTime()) &&
      (!lastMeaningful || lastSystem.getTime() > lastMeaningful.getTime()),
    sparseConversation: meaningfulCount > 0 && meaningfulCount <= 3,
    noMeaningfulMessages: meaningfulCount === 0 && messageCount > 0,
    participantMetadataDegraded: Boolean(conversation.participantDegraded),
  };
}

function buildSelectedConversation(conversation, {
  selectionReason = 'graphChatId',
  identityConfidence = 'high',
} = {}) {
  if (!conversation?.graphChatId) {
    return null;
  }

  return {
    archiveConversationId: conversation._id?.toString?.() || conversation.id || '',
    graphChatId: conversation.graphChatId,
    topic: conversation.topic || '',
    chatType: conversation.chatType || '',
    lastMessageAt: conversation.lastMessageAt || null,
    lastMeaningfulMessageAt: conversation.lastMeaningfulMessageAt || null,
    selectionReason,
    identityConfidence,
  };
}

function buildIdentityWarning(candidates = [], {
  topic = '',
  reason = 'multiple_conversations_match_title',
} = {}) {
  if (!Array.isArray(candidates) || candidates.length <= 1) {
    return null;
  }

  return {
    reason,
    topic: topic || undefined,
    candidateCount: candidates.length,
    guidance:
      'Multiple Teams conversations match this title/topic. Treat graphChatId as the stable identity and ask the user to choose before answering as if one chat is selected.',
    candidates: candidates.slice(0, 10).map((candidate) => ({
      archiveConversationId: candidate._id?.toString?.() || candidate.id || '',
      graphChatId: candidate.graphChatId,
      topic: candidate.topic || '',
      chatType: candidate.chatType || '',
      lastMessageAt: candidate.lastMessageAt || null,
      lastMeaningfulMessageAt: candidate.lastMeaningfulMessageAt || null,
    })),
  };
}

function buildIdentityChangeWarning({ priorGraphChatId, priorTopic, selectedConversation }) {
  const normalizedPriorGraphChatId = String(priorGraphChatId || '').trim();
  if (!normalizedPriorGraphChatId || !selectedConversation?.graphChatId) {
    return null;
  }

  if (normalizedPriorGraphChatId === selectedConversation.graphChatId) {
    return null;
  }

  const priorTopicText = String(priorTopic || '').trim().toLowerCase();
  const selectedTopicText = String(selectedConversation.topic || '').trim().toLowerCase();
  const similarTopic =
    priorTopicText &&
    selectedTopicText &&
    (priorTopicText === selectedTopicText ||
      priorTopicText.includes(selectedTopicText) ||
      selectedTopicText.includes(priorTopicText));

  if (!similarTopic) {
    return null;
  }

  return {
    identityChanged: true,
    previousGraphChatId: normalizedPriorGraphChatId,
    currentGraphChatId: selectedConversation.graphChatId,
    previousTopic: priorTopic || '',
    currentTopic: selectedConversation.topic || '',
    guidance:
      'This selection differs from the previous Teams conversation for a similar topic. Ask for clarification unless the user explicitly requested a different chat.',
  };
}

function isSameOrFollowUpTopic({ query, topic, priorTopic }) {
  const text = [query, topic].map((value) => String(value || '').trim()).filter(Boolean).join(' ');
  const normalizedText = text.toLowerCase();
  const normalizedPriorTopic = String(priorTopic || '').trim().toLowerCase();

  if (!normalizedText) {
    return true;
  }

  if (
    /\b(it|that|this|same|selected|previous|new|latest|changed|updates?|action items?|summarize)\b/i.test(
      normalizedText,
    )
  ) {
    return true;
  }

  if (!normalizedPriorTopic) {
    return false;
  }

  return (
    normalizedText.includes(normalizedPriorTopic) ||
    normalizedPriorTopic.includes(normalizedText)
  );
}

function getPriorGraphChatIdForFollowUp(options = {}) {
  const priorGraphChatId = String(options.priorGraphChatId || '').trim();
  if (!priorGraphChatId || String(options.chatId || '').trim()) {
    return '';
  }

  return isSameOrFollowUpTopic({
    query: options.query,
    topic: options.topic,
    priorTopic: options.priorTopic,
  })
    ? priorGraphChatId
    : '';
}

function isCompletenessSensitiveText(value = '') {
  return /\b(all|everything|entire|complete|completeness|exact|exactly|verbatim|wording|full context|new messages?|what changed|changed|latest|action items?|decisions?)\b/i.test(
    String(value || ''),
  );
}

function buildEvidenceBudget(options = {}, evidence = {}) {
  const requestedCompleteness = Boolean(
    options.requestedCompleteness ||
      isCompletenessSensitiveText(
        [options.query, options.topic, options.action].filter(Boolean).join(' '),
      ),
  );
  const fullBodiesReturned = Number(evidence.fullBodiesReturned || 0);
  const humanReadableReturned = Number(evidence.humanReadableReturned || 0);
  const previewOnlyReturned = Number(evidence.previewOnlyReturned || 0);
  const conversationsScoped = Number(evidence.conversationsScoped || 0);
  const totalMessagesConsidered = Number(evidence.totalMessagesConsidered || 0);
  const insufficiencyReasons = [];

  if (requestedCompleteness && conversationsScoped !== 1) {
    insufficiencyReasons.push('conversation_not_uniquely_scoped');
  }
  if (requestedCompleteness && totalMessagesConsidered === 0) {
    insufficiencyReasons.push('no_messages_considered');
  }
  if (requestedCompleteness && fullBodiesReturned === 0 && humanReadableReturned === 0 && previewOnlyReturned > 0) {
    insufficiencyReasons.push('preview_only_evidence');
  }
  if (requestedCompleteness && evidence.identityWarning) {
    insufficiencyReasons.push('ambiguous_conversation_identity');
  }

  const evidenceSufficient =
    !requestedCompleteness ||
    (insufficiencyReasons.length === 0 &&
      conversationsScoped === 1 &&
      totalMessagesConsidered > 0);

  return {
    requestedCompleteness,
    fullBodiesReturned,
    humanReadableReturned,
    previewOnlyReturned,
    conversationsScoped,
    totalMessagesConsidered,
    evidenceSufficient,
    insufficiencyReasons,
    guidance: evidenceSufficient
      ? 'Evidence is sufficient for the requested retrieval scope.'
      : 'Do not provide a definitive summary. Ask the user to select a conversation or run conversation_recent_messages/conversation_dossier for stronger evidence.',
  };
}

function buildHumanReadableMessageFilter({ includeSystem = false } = {}) {
  if (includeSystem) {
    return {};
  }

  return {
    isSystemLikeMessage: { $ne: true },
    isChunkable: true,
    normalizedTextLength: { $gt: 0 },
  };
}

function messageSenderMatches(message = {}, clauses = []) {
  const fromUserId = String(message.fromUserId || '').trim();
  const fromEmail = String(message.fromEmail || '').trim().toLowerCase();
  const fromDisplayName = String(message.fromDisplayName || '').trim().toLowerCase();

  for (const clause of clauses) {
    if (clause.fromUserId && String(clause.fromUserId).trim() === fromUserId) {
      return { matchedBy: 'fromUserId', confidence: 'high' };
    }
    if (clause.fromEmail && String(clause.fromEmail).trim().toLowerCase() === fromEmail) {
      return { matchedBy: 'fromEmail', confidence: fromEmail === 'aaduser' ? 'low' : 'high' };
    }
    if (
      clause.fromDisplayName &&
      String(clause.fromDisplayName).trim().toLowerCase() === fromDisplayName
    ) {
      return { matchedBy: 'fromDisplayName', confidence: 'medium' };
    }
  }

  return { matchedBy: 'unknown', confidence: 'low' };
}

function buildPersonSenderClauses({ senderName, senderEmail, senderUserId } = {}) {
  const clauses = [];
  const normalizedUserId = String(senderUserId || '').trim();
  const normalizedEmail = String(senderEmail || '').trim();
  const normalizedName = String(senderName || '').trim();

  if (normalizedUserId) {
    clauses.push({ fromUserId: normalizedUserId });
  }
  if (normalizedEmail) {
    clauses.push({ fromEmail: normalizedEmail }, { fromEmail: normalizedEmail.toLowerCase() });
  }
  if (normalizedName) {
    clauses.push({ fromDisplayName: normalizedName });
  }

  return clauses;
}

function summarizeMessageSearchability(messages = []) {
  const stats = {
    humanMessageCount: 0,
    meaningfulMessageCount: 0,
    systemMessageCount: 0,
    emptyMessageCount: 0,
    lastHumanMessageAt: null,
    lastMeaningfulMessageAt: null,
    lastSystemMessageAt: null,
  };

  for (const message of messages) {
    const sentAt = toDate(message?.sentDateTime) || toDate(message?.createdAt);
    const isMeaningful = isMeaningfulArchiveMessage(message);
    const isSystem = message?.isSystemLikeMessage === true;
    const isEmpty = isEmptyArchiveMessage(message);

    if (isMeaningful) {
      stats.humanMessageCount += 1;
      stats.meaningfulMessageCount += 1;
      if (!stats.lastHumanMessageAt || sentAt > stats.lastHumanMessageAt) {
        stats.lastHumanMessageAt = sentAt;
      }
      if (!stats.lastMeaningfulMessageAt || sentAt > stats.lastMeaningfulMessageAt) {
        stats.lastMeaningfulMessageAt = sentAt;
      }
    }

    if (isSystem) {
      stats.systemMessageCount += 1;
      if (!stats.lastSystemMessageAt || sentAt > stats.lastSystemMessageAt) {
        stats.lastSystemMessageAt = sentAt;
      }
    }

    if (isEmpty) {
      stats.emptyMessageCount += 1;
    }
  }

  return stats;
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

function buildIntentText(options = {}) {
  return [
    String(options.query || '').trim(),
    String(options.topic || '').trim(),
    ...toArray(options.participants).map((participant) => String(participant || '').trim()),
    String(options.chatType || '').trim(),
  ]
    .filter(Boolean)
    .join(' ')
    .toLowerCase();
}

function classifyRetrievalIntent(options = {}, { mode = 'advanced' } = {}) {
  const text = buildIntentText(options);
  const participantScoped = toArray(options.participants).filter(Boolean).length > 0;
  const oneOnOneQuery =
    String(options.chatType || '').trim() === 'oneOnOne' ||
    /\b(one[- ]on[- ]one|1:1|dm|direct message)\b/i.test(text);
  const exactnessSensitive =
    /\b(all|everything|entire|complete|completeness|exact|exactly|verbatim|wording)\b/i.test(
      text,
    );
  const detailSensitive =
    /\b(full body|full text|full context|details?|decisions?|action items?|next steps?|context|history)\b/i.test(
      text,
    );
  const recentOnly =
    mode === 'recent' ||
    /\b(recent|latest|today|yesterday|this week|last week)\b/i.test(text) ||
    Boolean(options.daysBack);
  const broadTopicDiscovery =
    Boolean(String(options.topic || options.query || '').trim()) &&
    !participantScoped &&
    !oneOnOneQuery &&
    !exactnessSensitive &&
    !detailSensitive;

  return {
    participantScoped,
    oneOnOneQuery,
    exactnessSensitive,
    broadTopicDiscovery,
    recentOnly,
    fullBodyDetailsQuery: detailSensitive,
    completenessSensitive: exactnessSensitive || detailSensitive,
  };
}

function getPreviewTextLength(result) {
  return String(
    result?.excerpt ||
      result?.bodyPreview ||
      result?.summary ||
      result?.text ||
      '',
  )
    .replace(/\s+/g, ' ')
    .trim().length;
}

function arePreviewsLikelyInsufficient(results = []) {
  if (!Array.isArray(results) || results.length === 0) {
    return false;
  }

  return results.every((result) => getPreviewTextLength(result) < 160);
}

function buildArchiveUnionDecision({ intent, memoryResults, requestedLimit }) {
  const reasons = [];
  const memoryResultCount = Array.isArray(memoryResults?.results) ? memoryResults.results.length : 0;
  const normalizedLimit = clampInteger(requestedLimit, 4, { min: 1, max: 12 });

  if (intent.broadTopicDiscovery) {
    reasons.push('broad_topic_query');
  }
  if (intent.exactnessSensitive) {
    reasons.push('exactness_requested');
  }
  if (intent.fullBodyDetailsQuery) {
    reasons.push('details_requested');
  }
  if (intent.participantScoped || intent.oneOnOneQuery) {
    reasons.push('person_or_dm_scope');
  }
  if (memoryResultCount > 0 && memoryResultCount < Math.min(3, normalizedLimit)) {
    reasons.push('low_memory_result_count');
  }
  if (memoryResultCount > 0 && arePreviewsLikelyInsufficient(memoryResults.results)) {
    reasons.push('preview_insufficient');
  }

  return {
    runArchiveUnion: memoryResultCount === 0 || reasons.length > 0,
    reasons,
    memoryResultCount,
  };
}

function dedupeMergedRetrievalResults({ archiveResults = [], memoryResults = [], limit = 6 }) {
  const merged = [];
  const indexByKey = new Map();

  const pushResult = (result, evidenceSource) => {
    const key = String(result?.graphMessageId || result?.sourceRecordId || result?.id || '').trim();
    const normalized = {
      ...result,
      evidenceSource,
    };

    if (!key) {
      merged.push(normalized);
      return;
    }

    const existingIndex = indexByKey.get(key);
    if (existingIndex === undefined) {
      indexByKey.set(key, merged.length);
      merged.push(normalized);
      return;
    }

    if (evidenceSource === 'archive') {
      merged[existingIndex] = {
        ...merged[existingIndex],
        ...normalized,
        evidenceSource: 'archive',
      };
    }
  };

  for (const result of archiveResults) {
    pushResult(result, 'archive');
  }

  for (const result of memoryResults) {
    pushResult(result, 'enterprise_memory');
  }

  return merged.slice(0, limit);
}

function resolveConversationFromSearchResults(results = [], explicitResolvedConversation) {
  if (explicitResolvedConversation?.graphChatId) {
    return explicitResolvedConversation;
  }

  const conversationIds = [...new Set(results.map((result) => result?.graphChatId).filter(Boolean))];
  if (conversationIds.length !== 1) {
    return null;
  }

  const exemplar = results.find((result) => result?.graphChatId === conversationIds[0]);
  return exemplar
    ? {
        graphChatId: exemplar.graphChatId,
        chatType: exemplar.chatType || '',
        topic: exemplar.topic || '',
        participants: exemplar.participants || [],
      }
    : null;
}

async function buildMessageBodyEscalations(user, options = {}, results = []) {
  const uniqueMessageIds = [
    ...new Set(
      results
        .map((result) => result?.graphMessageId || result?.id)
        .filter(Boolean),
    ),
  ].slice(0, 2);

  if (uniqueMessageIds.length === 0) {
    return [];
  }

  const escalations = [];
  for (const messageId of uniqueMessageIds) {
    const fullBody = await getMessageBody(user, {
      chatId: options.chatId,
      messageId,
    });
    if (fullBody?.resolved) {
      escalations.push(fullBody);
    }
  }

  return escalations;
}

async function runArchiveAdvancedMessageSearch(user, options = {}) {
  const userId = user?.id || user?._id?.toString();
  const topic = String(options.topic || options.query || '').trim();
  const priorFollowUpChatId = getPriorGraphChatIdForFollowUp(options);
  const chatId = String(options.chatId || priorFollowUpChatId || '').trim();
  const senderScope = String(options.senderScope || 'any').trim();
  const chatType = String(options.chatType || 'any').trim();
  const sortBy = String(options.sortBy || 'recent').trim();
  const limit = clampInteger(options.limit, Math.min(getTeamsArchiveConfig().defaultSearchLimit, 4), {
    max: 12,
  });
  const offset = clampInteger(options.offset, 0, { min: 0, max: 100000 });
  const daysBack = options.daysBack
    ? clampInteger(options.daysBack, 30, { min: 1, max: 3650 })
    : undefined;

  const validSenderScopes = new Set(['any', 'me', 'others']);
  const validChatTypes = new Set(['any', 'oneOnOne', 'group', 'meeting']);
  const normalizedSenderScope = validSenderScopes.has(senderScope) ? senderScope : 'any';
  const normalizedChatType = validChatTypes.has(chatType) ? chatType : 'any';
  const participantClauses = buildParticipantConversationClauses(options.participants);
  const resolvedGraphChatId = chatId
    ? await resolveConversationGraphChatId(userId, chatId)
    : undefined;
  const explicitConversationFilters =
    !resolvedGraphChatId && (normalizedChatType !== 'any' || participantClauses.length > 0);

  if (chatId && !resolvedGraphChatId) {
    return {
      retrievalMode: 'advanced_message_previews',
      topic: topic || undefined,
      chatId,
      graphChatId: null,
      senderScope: normalizedSenderScope,
      chatType: normalizedChatType,
      daysBack,
      participants: toArray(options.participants).filter(Boolean),
      guidance: 'No archived Teams chat was found for the requested chat id.',
      resolvedConversation: undefined,
      resultCount: 0,
      results: [],
      trace: {
        archiveSearched: true,
        archiveResultCount: 0,
        conversationFiltersApplied: {
          chatId: true,
          chatType: normalizedChatType !== 'any',
          participants: participantClauses.length > 0,
          daysBack: Boolean(daysBack),
        },
        topicGateApplied: false,
        matchedConversationCount: 0,
      },
    };
  }

  let matchedConversations = [];
  let matchedConversationIds = [];
  if (explicitConversationFilters) {
    const lookup = await findConversationCandidates(userId, {
      ...options,
      topic: '',
      query: '',
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
          'No matching chats were found for the requested participant or chat-type constraints. Consider broadening those filters.',
        resolvedConversation: undefined,
        resultCount: 0,
        results: [],
        trace: {
          archiveSearched: true,
          archiveResultCount: 0,
          conversationFiltersApplied: {
            chatType: normalizedChatType !== 'any',
            participants: participantClauses.length > 0,
            daysBack: Boolean(daysBack),
          },
          topicGateApplied: false,
          matchedConversationCount: 0,
        },
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

  const { clauses: topicClauses } = buildTopicSearchClauses(topic);
  const messageFilter = {
    user: userId,
    ...(resolvedGraphChatId
      ? { graphChatId: resolvedGraphChatId }
      : matchedConversationIds.length > 0
        ? { graphChatId: { $in: matchedConversationIds } }
        : {}),
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
    matchedConversations.length === 1
      ? mapConversationCandidate(matchedConversations[0])
      : conversationIds.length === 1
        ? conversationMap.get(conversationIds[0])
          ? mapConversationCandidate(conversationMap.get(conversationIds[0]))
          : undefined
        : undefined;
  const selectedConversation =
    resolvedConversation?.graphChatId
      ? {
          archiveConversationId: resolvedConversation.id || '',
          graphChatId: resolvedConversation.graphChatId,
          topic: resolvedConversation.topic || '',
          chatType: resolvedConversation.chatType || '',
          lastMessageAt: resolvedConversation.lastMessageAt || null,
          lastMeaningfulMessageAt: resolvedConversation.lastMeaningfulMessageAt || null,
          selectionReason: resolvedGraphChatId
            ? priorFollowUpChatId && !options.chatId
              ? 'prior_context'
              : 'chatId'
            : 'single_result',
          identityConfidence: resolvedGraphChatId ? 'high' : 'medium',
        }
      : null;

  return {
    retrievalMode: 'advanced_message_previews',
    topic: topic || undefined,
    senderScope: normalizedSenderScope,
    chatType: normalizedChatType,
    ...(resolvedGraphChatId ? { chatId, graphChatId: resolvedGraphChatId } : {}),
    daysBack,
    participants: toArray(options.participants).filter(Boolean),
    guidance:
      resolvedConversation
        ? 'These previews are scoped to one resolved conversation. If completeness matters, use conversation_dossier or summarize_conversation before expanding with get_messages_window.'
        : 'These are compact previews optimized for topic discovery. If a single conversation stands out, use summarize_conversation before expanding with get_messages_window.',
    ...(resolvedConversation ? { resolvedConversation } : {}),
    ...(selectedConversation ? { selectedConversation } : {}),
    resultCount: messages.length,
    evidenceBudget: buildEvidenceBudget(options, {
      fullBodiesReturned: 0,
      humanReadableReturned: messages.filter((message) => isMeaningfulArchiveMessage(message)).length,
      previewOnlyReturned: messages.length,
      conversationsScoped: resolvedGraphChatId ? 1 : conversationIds.length,
      totalMessagesConsidered: messages.length,
    }),
    results: messages.map((message) =>
      mapCompactMessageResult(message, conversationMap.get(message.graphChatId), {
        excerptLimit: 260,
      }),
    ),
    trace: {
      archiveSearched: true,
      archiveResultCount: messages.length,
      conversationFiltersApplied: {
        chatId: Boolean(resolvedGraphChatId),
        chatType: normalizedChatType !== 'any',
        participants: participantClauses.length > 0,
        daysBack: Boolean(daysBack),
      },
      topicGateApplied: false,
      matchedConversationCount: matchedConversations.length,
    },
  };
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
        ? { lastMeaningfulMessageAt: { $gte: new Date(Date.now() - daysBack * 24 * 60 * 60 * 1000) } }
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
    sort: { lastMeaningfulMessageAt: -1, lastMessageAt: -1, updatedAt: -1 },
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
                lastMeaningfulMessageAt: {
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
            sort: { lastMeaningfulMessageAt: -1, lastMessageAt: -1, updatedAt: -1 },
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
    return {
      participants: [],
      failed: controller?.mode === 'disabled' || controller?.disabledChatTypes?.has(normalizedChatType),
      pageCount: 0,
    };
  }

  try {
    let nextLink = null;
    let pageCount = 0;
    const participants = [];

    do {
      const response = await graphRequest(
        user,
        nextLink || `/chats/${encodeURIComponent(chatId)}/members`,
        nextLink
          ? {
              suppressErrorLog: true,
            }
          : {
              query: { $top: 50 },
              suppressErrorLog: true,
            },
      );

      pageCount += 1;
      participants.push(
        ...toArray(response?.value).map((member) => ({
          displayName: member?.displayName || member?.email || '',
          email: member?.email || '',
          userId: member?.userId || member?.id || '',
          source: 'graph',
          confidence: member?.userId || member?.email ? 'high' : 'medium',
        })),
      );
      nextLink = response?.['@odata.nextLink'] || null;
    } while (nextLink);

    if (controller) {
      controller.stats.successCount += pageCount;
    }

    return {
      participants: uniqueParticipants(participants),
      failed: false,
      pageCount,
    };
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

    return {
      participants: [],
      failed: true,
      pageCount: 0,
    };
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
    zeroChunkReasonCounts: diagnostics.zeroChunkReasonCounts || {},
    searchableConversationCountsByChatType: diagnostics.searchableConversationCountsByChatType || {},
    participantDegradedConversationCount: Number(diagnostics.participantDegradedConversationCount || 0),
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
    searchableOneOnOneConversationCount,
    searchableGroupConversationCount,
    searchableMeetingConversationCount,
    participantDegradedConversationCount,
    staleOrIncompleteConversationCount,
    zeroMeaningfulMessageConversationCount,
    systemOnlyRecentConversationCount,
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
    userId && typeof db.countDistinctEnterpriseMemoryChunkField === 'function'
      ? db.countDistinctEnterpriseMemoryChunkField('sourceParentRecordId', {
          user: userId,
          source: 'teams',
          'metadata.chatType': 'oneOnOne',
          chunkType: { $in: ['message', 'conversation_window'] },
        })
      : 0,
    userId && typeof db.countDistinctEnterpriseMemoryChunkField === 'function'
      ? db.countDistinctEnterpriseMemoryChunkField('sourceParentRecordId', {
          user: userId,
          source: 'teams',
          'metadata.chatType': 'group',
          chunkType: { $in: ['message', 'conversation_window'] },
        })
      : 0,
    userId && typeof db.countDistinctEnterpriseMemoryChunkField === 'function'
      ? db.countDistinctEnterpriseMemoryChunkField('sourceParentRecordId', {
          user: userId,
          source: 'teams',
          'metadata.chatType': 'meeting',
          chunkType: { $in: ['message', 'conversation_window'] },
        })
      : 0,
    userId ? db.countTeamsArchiveConversations({ user: userId, participantDegraded: true }) : 0,
    userId
      ? db.countTeamsArchiveConversations({
          user: userId,
          $or: [
            { syncStatus: { $in: ['failed', 'pending', 'running'] } },
            { syncCursor: { $exists: true, $nin: [null, ''] } },
            { messageCount: 0, lastMessageAt: { $exists: true, $ne: null } },
          ],
        })
      : 0,
    userId
      ? db.countTeamsArchiveConversations({
          user: userId,
          messageCount: { $gt: 0 },
          meaningfulMessageCount: { $in: [0, null] },
        })
      : 0,
    userId
      ? db.countTeamsArchiveConversations({
          user: userId,
          lastSystemMessageAt: { $exists: true, $ne: null },
          $or: [
            { lastMeaningfulMessageAt: { $exists: false } },
            { lastMeaningfulMessageAt: null },
            { meaningfulMessageCount: 0 },
          ],
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
    searchabilityDiagnostics: {
      discoveredConversationCount: backfillState?.discoveredChatCount || conversationCount || 0,
      archivedConversationCount: conversationCount || 0,
      projectedConversationCount: projectionEntityConversationCount || 0,
      searchableConversationCount: projectionConversationCount || 0,
      zeroChunkConversationCount:
        normalizeProjectionDiagnostics(latestProjection?.stats || {}).zeroChunkConversationCount || 0,
      zeroChunkReasonCounts:
        normalizeProjectionDiagnostics(latestProjection?.stats || {}).zeroChunkReasonCounts || {},
      searchableConversationCountsByChatType: {
        oneOnOne: searchableOneOnOneConversationCount || 0,
        group: searchableGroupConversationCount || 0,
        meeting: searchableMeetingConversationCount || 0,
      },
      participantDegradedConversationCount: participantDegradedConversationCount || 0,
      staleOrIncompleteConversations: staleOrIncompleteConversationCount || 0,
      zeroMeaningfulMessageConversations: zeroMeaningfulMessageConversationCount || 0,
      systemOnlyRecentConversations: systemOnlyRecentConversationCount || 0,
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
        const existingConversation = existingConversationMap.get(chat.id);
        const memberLookup = await listChatMembers(user, chat.id, chatType, memberLookupController);
        const normalizedConversation = normalizeConversation(chat, memberLookup.participants);
        const enrichedParticipants = mergeConversationParticipants({
          graphParticipants: normalizedConversation.participants,
          existingParticipants: existingConversation?.participants || [],
          memberLookupFailed: memberLookup.failed,
        });
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
          participants: enrichedParticipants.participants,
          participantMetadataSource: enrichedParticipants.participantMetadataSource,
          participantConfidence: enrichedParticipants.participantConfidence,
          participantDegraded: enrichedParticipants.participantDegraded,
          participantStats: enrichedParticipants.participantStats,
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
      const inferredMessageParticipants = [];
      const inferredMentionParticipants = [];

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
            const inferredParticipants = extractInferredParticipants(normalizedMessages);
            inferredMessageParticipants.push(...inferredParticipants.inferredMessageParticipants);
            inferredMentionParticipants.push(...inferredParticipants.inferredMentionParticipants);
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
      const archivedMessagesForStats = await db.findTeamsArchiveMessages(
        {
          user: userId,
          graphChatId: conversation.graphChatId,
        },
        {
          limit: DEFAULT_CONVERSATION_DOSSIER_MAX_MESSAGES,
          offset: 0,
          sort: { sentDateTime: -1, createdAt: -1 },
        },
      );
      const searchabilityStats = summarizeMessageSearchability(archivedMessagesForStats);
      const isConversationComplete = !conversationFailed && (incrementalRefresh || !nextMessageCursor);
      const enrichedParticipants = mergeConversationParticipants({
        graphParticipants: toArray(conversation?.participants).filter(
          (participant) => participant?.source === 'graph' || participant?.source === 'mixed',
        ),
        existingParticipants: conversation?.participants || [],
        inferredMessageParticipants,
        inferredMentionParticipants,
        memberLookupFailed: Boolean(conversation?.participantDegraded),
      });

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
        lastHumanMessageAt: searchabilityStats.lastHumanMessageAt,
        lastMeaningfulMessageAt: searchabilityStats.lastMeaningfulMessageAt,
        lastSystemMessageAt: searchabilityStats.lastSystemMessageAt,
        humanMessageCount: searchabilityStats.humanMessageCount,
        meaningfulMessageCount: searchabilityStats.meaningfulMessageCount,
        systemMessageCount: searchabilityStats.systemMessageCount,
        emptyMessageCount: searchabilityStats.emptyMessageCount,
        participants: enrichedParticipants.participants,
        participantMetadataSource: enrichedParticipants.participantMetadataSource,
        participantConfidence: enrichedParticipants.participantConfidence,
        participantDegraded: enrichedParticipants.participantDegraded,
        participantStats: enrichedParticipants.participantStats,
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
        lastMeaningfulMessageAt: conversationRecord?.lastMeaningfulMessageAt || searchabilityStats.lastMeaningfulMessageAt,
        meaningfulMessageCount:
          conversationRecord?.meaningfulMessageCount || searchabilityStats.meaningfulMessageCount,
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
    { limit, offset, sort: { lastMeaningfulMessageAt: -1, lastMessageAt: -1, updatedAt: -1 } },
  );
  const selectedConversation =
    conversations.length === 1
      ? buildSelectedConversation(conversations[0], {
          selectionReason: topic || participantClauses.length > 0 ? 'filtered_single_match' : 'single_result',
          identityConfidence: topic ? 'medium' : 'high',
        })
      : null;
  const identityWarning = buildIdentityWarning(conversations, {
    topic,
    reason: topic ? 'multiple_conversations_match_title_or_topic' : 'multiple_conversations_match_filters',
  });

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
    ...(selectedConversation ? { selectedConversation } : {}),
    ...(identityWarning ? { identityWarning } : {}),
    conversations: conversations.map((conversation) => mapCompactConversation(conversation)),
  };
}

async function computeConversationRecencyFromArchive(userId, graphChatId, options = {}) {
  const limit = clampInteger(
    options.limit,
    DEFAULT_RECENCY_BACKFILL_MAX_MESSAGES,
    { min: 1, max: DEFAULT_RECENCY_BACKFILL_MAX_MESSAGES },
  );
  const [totalMessageCount, messages] = await Promise.all([
    db.countTeamsArchiveMessages({ user: userId, graphChatId }),
    db.findTeamsArchiveMessages(
      { user: userId, graphChatId },
      {
        limit,
        offset: 0,
        sort: { sentDateTime: -1, createdAt: -1 },
      },
    ),
  ]);
  const stats = summarizeMessageSearchability(messages);
  const newestRawMessage = messages.find((message) => toDate(message.sentDateTime) || toDate(message.createdAt));

  return {
    ...stats,
    totalMessageCount,
    loadedMessageCount: messages.length,
    truncated: totalMessageCount > messages.length,
    lastMessageAt:
      toDate(newestRawMessage?.sentDateTime) ||
      toDate(newestRawMessage?.createdAt) ||
      null,
  };
}

function buildRecencyBackfillUpdate(stats = {}) {
  return {
    lastHumanMessageAt: stats.lastHumanMessageAt,
    lastMeaningfulMessageAt: stats.lastMeaningfulMessageAt,
    lastSystemMessageAt: stats.lastSystemMessageAt,
    humanMessageCount: stats.humanMessageCount,
    meaningfulMessageCount: stats.meaningfulMessageCount,
    systemMessageCount: stats.systemMessageCount,
    emptyMessageCount: stats.emptyMessageCount,
    messageCount: stats.totalMessageCount,
    ...(stats.lastMessageAt ? { lastMessageAt: stats.lastMessageAt } : {}),
  };
}

function pickRecencyFields(conversation = {}) {
  return {
    lastMessageAt: conversation.lastMessageAt || null,
    lastHumanMessageAt: conversation.lastHumanMessageAt || null,
    lastMeaningfulMessageAt: conversation.lastMeaningfulMessageAt || null,
    lastSystemMessageAt: conversation.lastSystemMessageAt || null,
    humanMessageCount: conversation.humanMessageCount || 0,
    meaningfulMessageCount: conversation.meaningfulMessageCount || 0,
    systemMessageCount: conversation.systemMessageCount || 0,
    emptyMessageCount: conversation.emptyMessageCount || 0,
    messageCount: conversation.messageCount || 0,
  };
}

function recencyFieldsChanged(oldFields = {}, newFields = {}) {
  return Object.keys(newFields).some((key) => {
    const oldValue = oldFields[key];
    const newValue = newFields[key];
    if (oldValue instanceof Date || newValue instanceof Date) {
      return (toDate(oldValue)?.getTime?.() || 0) !== (toDate(newValue)?.getTime?.() || 0);
    }
    return oldValue !== newValue;
  });
}

async function backfillConversationRecency(user, options = {}) {
  assertEnabled();
  const userId = user?.id || user?._id?.toString();
  if (!userId) {
    throw new TeamsArchiveServiceError('User id is required for recency backfill', 400);
  }

  const chatId = String(options.chatId || '').trim();
  const apply = options.apply === true;
  const dryRun = !apply;
  let conversations = [];

  if (chatId) {
    const resolvedGraphChatId = await resolveConversationGraphChatId(userId, chatId);
    if (!resolvedGraphChatId) {
      return {
        retrievalMode: 'conversation_recency_backfill',
        dryRun,
        apply,
        chatId,
        processedConversationCount: 0,
        changedConversationCount: 0,
        conversations: [],
        guidance: 'No archived Teams conversation was found for the requested chat id.',
      };
    }
    conversations = await db.findTeamsArchiveConversations(
      { user: userId, graphChatId: resolvedGraphChatId },
      { limit: 1 },
    );
  } else {
    conversations = await db.findTeamsArchiveConversations(
      { user: userId },
      {
        limit: clampInteger(options.limit, 10000, { min: 1, max: 100000 }),
        offset: clampInteger(options.offset, 0, { min: 0, max: 100000 }),
        sort: { updatedAt: -1 },
      },
    );
  }

  const results = [];
  for (const conversation of conversations) {
    const stats = await computeConversationRecencyFromArchive(userId, conversation.graphChatId);
    const oldRecency = pickRecencyFields(conversation);
    const newRecency = buildRecencyBackfillUpdate(stats);
    const wouldChange = recencyFieldsChanged(oldRecency, newRecency);

    if (apply && wouldChange) {
      await db.updateTeamsArchiveConversation(
        conversation._id?.toString?.() || conversation.id,
        newRecency,
      );
    }

    results.push({
      topic: conversation.topic || '',
      graphChatId: conversation.graphChatId,
      archiveConversationId: conversation._id?.toString?.() || conversation.id || '',
      oldRecency,
      newRecency,
      totalMessageCount: stats.totalMessageCount,
      loadedMessageCount: stats.loadedMessageCount,
      truncated: stats.truncated,
      wouldChange,
      didChange: apply && wouldChange,
    });
  }

  return {
    retrievalMode: 'conversation_recency_backfill',
    dryRun,
    apply,
    processedConversationCount: results.length,
    changedConversationCount: results.filter((result) => result.wouldChange).length,
    updatedConversationCount: results.filter((result) => result.didChange).length,
    conversations: results,
  };
}

async function recentMeetingChats(user, options = {}) {
  assertEnabled();
  const userId = user?.id || user?._id?.toString();
  const limit = clampInteger(options.limit, 5, { min: 1, max: 10 });
  const offset = clampInteger(options.offset, 0, { min: 0, max: 100000 });
  const daysBack = options.daysBack
    ? clampInteger(options.daysBack, 30, { min: 1, max: 3650 })
    : undefined;
  const topic = String(options.topic || options.query || '').trim();
  const topicRegex = topic ? buildSearchRegex(topic) : null;
  const since = daysBack ? new Date(Date.now() - daysBack * 24 * 60 * 60 * 1000) : null;

  const conversations = await db.findTeamsArchiveConversations(
    {
      user: userId,
      chatType: 'meeting',
      ...(since ? { lastMeaningfulMessageAt: { $gte: since } } : {}),
      ...(topicRegex ? { topic: topicRegex } : {}),
    },
    {
      limit,
      offset,
      sort: { lastMeaningfulMessageAt: -1, lastMessageAt: -1, updatedAt: -1 },
    },
  );

  const graphChatIds = conversations.map((conversation) => conversation.graphChatId).filter(Boolean);
  const previewMessages = graphChatIds.length
    ? await db.findTeamsArchiveMessages(
        {
          user: userId,
          graphChatId: { $in: graphChatIds },
          ...buildHumanReadableMessageFilter(),
        },
        {
          limit: Math.max(limit * 3, graphChatIds.length),
          offset: 0,
          sort: { sentDateTime: -1, createdAt: -1 },
        },
      )
    : [];
  const previewByChatId = new Map();
  for (const message of previewMessages) {
    const existing = previewByChatId.get(message.graphChatId) || [];
    if (existing.length < 3) {
      existing.push(message);
      previewByChatId.set(message.graphChatId, existing);
    }
  }

  const selectedConversation =
    conversations.length === 1
      ? buildSelectedConversation(conversations[0], {
          selectionReason: 'recent_meeting_single_result',
          identityConfidence: topic ? 'medium' : 'high',
        })
      : null;
  const identityWarning = buildIdentityWarning(conversations, {
    topic,
    reason: topic ? 'multiple_recent_meetings_match_title_or_topic' : 'multiple_recent_meetings',
  });

  return {
    retrievalMode: 'recent_meeting_chats',
    chatType: 'meeting',
    topic: topic || undefined,
    daysBack,
    guidance:
      'Meeting chats are ranked by lastMeaningfulMessageAt, not raw Teams system activity. Use selectedConversation.graphChatId for follow-up questions, and ask for clarification when identityWarning is present.',
    ...(selectedConversation ? { selectedConversation } : {}),
    ...(identityWarning ? { identityWarning } : {}),
    resultCount: conversations.length,
    conversations: conversations.map((conversation) => ({
      ...mapCompactConversation(conversation),
      newestHumanReadableMessages: (previewByChatId.get(conversation.graphChatId) || []).map((message) =>
        mapCompactMessageResult(message, conversation, {
          includeReplyToId: true,
          excerptLimit: 240,
        }),
      ),
    })),
  };
}

async function getConversationDossier(user, options = {}) {
  assertEnabled();
  const userId = user?.id || user?._id?.toString();
  const priorFollowUpChatId = getPriorGraphChatIdForFollowUp(options);
  const chatId = String(options.chatId || priorFollowUpChatId || '').trim();
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
      const identityWarning = buildIdentityWarning(candidates, {
        topic: lookup.topic || query,
        reason: 'multiple_conversations_match_title_or_filters',
      });
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
        ...(identityWarning ? { identityWarning } : {}),
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
  const selectedConversation = buildSelectedConversation(conversation, {
    selectionReason: chatId
      ? priorFollowUpChatId && !options.chatId
        ? 'prior_context'
        : 'chatId'
      : 'single_candidate',
    identityConfidence: chatId ? 'high' : 'medium',
  });
  const identityChangeWarning = buildIdentityChangeWarning({
    priorGraphChatId: options.priorGraphChatId,
    priorTopic: options.priorTopic,
    selectedConversation,
  });

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
    ...(selectedConversation ? { selectedConversation } : {}),
    ...(identityChangeWarning ? { identityWarning: identityChangeWarning } : {}),
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
    evidenceBudget: buildEvidenceBudget(options, {
      fullBodiesReturned: messages.length,
      humanReadableReturned: matchedMessages.length || messages.length,
      previewOnlyReturned: 0,
      conversationsScoped: 1,
      totalMessagesConsidered: totalMessagesInScope,
    }),
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

  if (!looksLikeMongoObjectId(normalizedChatId)) {
    return null;
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

  if (!looksLikeMongoObjectId(normalizedMessageId)) {
    return null;
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
  const selectedConversation = buildSelectedConversation(conversation, {
    selectionReason: 'chatId',
    identityConfidence: 'high',
  });

  return {
    retrievalMode: 'thread_previews',
    chatId,
    graphChatId: resolvedGraphChatId,
    ...(selectedConversation ? { selectedConversation } : {}),
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

async function conversationRecentMessages(user, options = {}) {
  assertEnabled();
  const userId = user?.id || user?._id?.toString();
  const chatId = String(options.chatId || '').trim();
  if (!chatId) {
    throw new TeamsArchiveServiceError('Chat id is required for conversation_recent_messages', 400);
  }

  const limit = clampInteger(options.limit, 6, { min: 1, max: 20 });
  const offset = clampInteger(options.offset, 0, { min: 0, max: 100000 });
  const includeSystem =
    options.includeSystem === true || String(options.includeSystem || '').toLowerCase() === 'true';
  const daysBack = options.daysBack
    ? clampInteger(options.daysBack, 30, { min: 1, max: 3650 })
    : undefined;
  const resolvedGraphChatId = await resolveConversationGraphChatId(userId, chatId);

  if (!resolvedGraphChatId) {
    return {
      retrievalMode: 'conversation_recent_messages',
      resolved: false,
      chatId,
      graphChatId: null,
      messages: [],
      guidance: 'No archived Teams chat was found for the requested chat id.',
    };
  }

  const since = daysBack ? new Date(Date.now() - daysBack * 24 * 60 * 60 * 1000) : null;
  const baseFilter = {
    user: userId,
    graphChatId: resolvedGraphChatId,
    ...(since ? { sentDateTime: { $gte: since } } : {}),
  };
  const [conversations, messages, systemMessages, emptyMessages, skippedMessages] = await Promise.all([
    db.findTeamsArchiveConversations({ user: userId, graphChatId: resolvedGraphChatId }, { limit: 1 }),
    db.findTeamsArchiveMessages(
      {
        ...baseFilter,
        ...buildHumanReadableMessageFilter({ includeSystem }),
      },
      { limit, offset, sort: { sentDateTime: -1, createdAt: -1 } },
    ),
    db.countTeamsArchiveMessages({ ...baseFilter, isSystemLikeMessage: true }),
    db.countTeamsArchiveMessages({ ...baseFilter, normalizedTextLength: { $lte: 0 } }),
    db.countTeamsArchiveMessages({ ...baseFilter, isChunkable: false }),
  ]);
  const conversation = conversations[0] || null;
  const selectedConversation = buildSelectedConversation(conversation, {
    selectionReason: 'chatId',
    identityConfidence: 'high',
  });

  return {
    retrievalMode: 'conversation_recent_messages',
    resolved: true,
    chatId,
    graphChatId: resolvedGraphChatId,
    ...(selectedConversation ? { selectedConversation } : {}),
    topic: conversation?.topic || '',
    chatType: conversation?.chatType || '',
    daysBack,
    includeSystem,
    guidance:
      'These are newest messages first from one resolved Teams conversation. By default system/empty messages are excluded so “new” reflects human-readable activity.',
    counts: {
      returned: messages.length,
      systemMessages,
      emptyMessages,
      skippedMessages,
    },
    evidenceBudget: buildEvidenceBudget(
      { ...options, action: 'conversation_recent_messages' },
      {
        fullBodiesReturned: 0,
        humanReadableReturned: messages.length,
        previewOnlyReturned: messages.length,
        conversationsScoped: 1,
        totalMessagesConsidered: messages.length,
      },
    ),
    messages: messages.map((message) =>
      mapCompactMessageResult(message, conversation, {
        includeReplyToId: true,
        includeImportance: true,
        excerptLimit: 420,
      }),
    ),
  };
}

async function conversationSenderMessages(user, options = {}) {
  assertEnabled();
  const userId = user?.id || user?._id?.toString();
  const priorFollowUpChatId = getPriorGraphChatIdForFollowUp(options);
  const chatId = String(options.chatId || priorFollowUpChatId || '').trim();
  if (!chatId) {
    throw new TeamsArchiveServiceError('Chat id is required for conversation_sender_messages', 400);
  }

  const senderScope = String(options.senderScope || 'me').trim();
  const normalizedSenderScope = ['me', 'person', 'all'].includes(senderScope) ? senderScope : 'me';
  const includeSystem =
    options.includeSystem === true || String(options.includeSystem || '').toLowerCase() === 'true';
  const limit = clampInteger(options.limit, 8, { min: 1, max: 50 });
  const offset = clampInteger(options.offset, 0, { min: 0, max: 100000 });
  const sortDirection = String(options.sort || options.sortBy || 'newest').toLowerCase() === 'oldest' ? 1 : -1;
  const resolvedGraphChatId = await resolveConversationGraphChatId(userId, chatId);

  if (!resolvedGraphChatId) {
    return {
      retrievalMode: 'conversation_sender_messages',
      resolved: false,
      chatId,
      graphChatId: null,
      messages: [],
      guidance: 'No archived Teams chat was found for the requested chat id.',
    };
  }

  const senderClauses =
    normalizedSenderScope === 'me'
      ? getUserSenderClauses(user)
      : normalizedSenderScope === 'person'
        ? buildPersonSenderClauses(options)
        : [];
  const senderFilter = senderClauses.length > 0 ? { $or: senderClauses } : {};
  const baseFilter = {
    user: userId,
    graphChatId: resolvedGraphChatId,
    ...senderFilter,
  };
  const [conversations, messages, totalMatchingMessages, skippedSystemCount, skippedEmptyCount] =
    await Promise.all([
      db.findTeamsArchiveConversations({ user: userId, graphChatId: resolvedGraphChatId }, { limit: 1 }),
      db.findTeamsArchiveMessages(
        {
          ...baseFilter,
          ...buildHumanReadableMessageFilter({ includeSystem }),
        },
        { limit, offset, sort: { sentDateTime: sortDirection, createdAt: sortDirection } },
      ),
      db.countTeamsArchiveMessages(baseFilter),
      db.countTeamsArchiveMessages({ ...baseFilter, isSystemLikeMessage: true }),
      db.countTeamsArchiveMessages({ ...baseFilter, normalizedTextLength: { $lte: 0 } }),
    ]);
  const conversation = conversations[0] || null;
  const selectedConversation = buildSelectedConversation(conversation, {
    selectionReason: priorFollowUpChatId && !options.chatId ? 'prior_context' : 'chatId',
    identityConfidence: 'high',
  });
  const senderResolution = {
    senderScope: normalizedSenderScope,
    senderName: options.senderName || '',
    senderEmail: options.senderEmail || '',
    senderUserId: options.senderUserId || '',
    clauseCount: senderClauses.length,
    confidence:
      normalizedSenderScope === 'all'
        ? 'high'
        : senderClauses.some((clause) => clause.fromUserId || clause.fromEmail)
          ? 'high'
          : senderClauses.length > 0
            ? 'medium'
            : 'low',
    warnings:
      normalizedSenderScope !== 'all' && senderClauses.length === 0
        ? ['sender_identity_could_not_be_resolved']
        : [],
  };
  const mappedMessages = messages.map((message) => ({
    ...mapCompactMessageResult(message, conversation, {
      includeReplyToId: true,
      includeImportance: true,
      excerptLimit: 420,
    }),
    senderMatch: messageSenderMatches(message, senderClauses),
  }));

  return {
    retrievalMode: 'conversation_sender_messages',
    resolved: true,
    chatId,
    graphChatId: resolvedGraphChatId,
    ...(selectedConversation ? { selectedConversation } : {}),
    senderResolution,
    skippedSystemCount,
    skippedEmptyCount,
    totalMatchingMessages,
    retrievalTrace: {
      senderScope: normalizedSenderScope,
      senderFilterApplied: senderClauses.length > 0,
      includeSystem,
      sort: sortDirection === -1 ? 'newest' : 'oldest',
      returnedMessageCount: mappedMessages.length,
    },
    evidenceBudget: buildEvidenceBudget(
      { ...options, action: 'conversation_sender_messages' },
      {
        fullBodiesReturned: 0,
        previewOnlyReturned: mappedMessages.length,
        conversationsScoped: 1,
        totalMessagesConsidered: totalMatchingMessages,
      },
    ),
    messages: mappedMessages,
  };
}

function buildActivityDiagnosis(conversation = {}) {
  const warnings = buildConversationWarningFlags(conversation);
  const lastMeaningful = toDate(conversation.lastMeaningfulMessageAt);
  const lastRaw = toDate(conversation.lastMessageAt);
  const lastSync = toDate(conversation.lastMessageSyncAt);
  const sourceUpdated = toDate(conversation.sourceUpdatedAt);
  const messageCount = Number(conversation.messageCount || 0);
  const meaningfulCount = Number(conversation.meaningfulMessageCount || 0);
  const possiblyIncompleteSync =
    ['failed', 'pending', 'running'].includes(String(conversation.syncStatus || '')) ||
    Boolean(conversation.syncCursor) ||
    (messageCount === 0 && Boolean(lastRaw)) ||
    Boolean(lastSync && sourceUpdated && lastSync.getTime() + 5 * 60 * 1000 < sourceUpdated.getTime());
  const truncatedHistoryRisk =
    Boolean(conversation.syncCursor) ||
    messageCount >= getTeamsArchiveConfig().defaultMessagesPerChat;
  const zeroChunkRisk =
    (messageCount > 0 && meaningfulCount === 0) ||
    warnings.noMeaningfulMessages;

  return {
    recentBecauseOfHumanActivity: Boolean(lastMeaningful && lastRaw && lastMeaningful.getTime() === lastRaw.getTime()),
    recentBecauseOfSystemActivity: warnings.systemOnlyRecentActivity,
    sparseConversation: warnings.sparseConversation,
    noMeaningfulMessages: warnings.noMeaningfulMessages,
    participantMetadataDegraded: warnings.participantMetadataDegraded,
    possiblyIncompleteSync,
    zeroChunkRisk,
    truncatedHistoryRisk,
  };
}

function buildActivityExplanation(conversation = {}, diagnosis = {}) {
  const topic = conversation.topic || 'this Teams chat';
  if (diagnosis.recentBecauseOfSystemActivity) {
    return `${topic} appears recent because raw Teams activity is newer than the latest meaningful human-readable message. Treat recency as system-driven unless the human message previews support otherwise.`;
  }
  if (diagnosis.recentBecauseOfHumanActivity) {
    return `${topic} appears recent because the latest raw Teams activity matches the latest meaningful human-readable message.`;
  }
  if (diagnosis.noMeaningfulMessages) {
    return `${topic} has archived records but no meaningful human-readable messages. Do not summarize it as an active discussion without more evidence.`;
  }
  if (diagnosis.possiblyIncompleteSync) {
    return `${topic} may have incomplete archive coverage due to sync status, paging cursor, or stale message sync timestamps.`;
  }
  return `${topic} has archived Teams metadata and message diagnostics available. Use the timestamps and flags before describing it as recent or complete.`;
}

async function conversationActivityDiagnostics(user, options = {}) {
  assertEnabled();
  const userId = user?.id || user?._id?.toString();
  const chatId = String(options.chatId || getPriorGraphChatIdForFollowUp(options) || '').trim();
  if (!chatId) {
    throw new TeamsArchiveServiceError('Chat id is required for conversation_activity_diagnostics', 400);
  }

  const includeRecentMessages =
    options.includeRecentMessages === true ||
    String(options.includeRecentMessages || '').toLowerCase() === 'true';
  const includeSystem =
    options.includeSystem === true || String(options.includeSystem || '').toLowerCase() === 'true';
  const limit = clampInteger(options.limit, 5, { min: 1, max: 20 });
  const resolvedGraphChatId = await resolveConversationGraphChatId(userId, chatId);

  if (!resolvedGraphChatId) {
    return {
      retrievalMode: 'conversation_activity_diagnostics',
      resolved: false,
      chatId,
      graphChatId: null,
      diagnosis: null,
      explanation: 'No archived Teams chat was found for the requested chat id.',
    };
  }

  const [conversations, newestHumanMessages, newestSystemMessages] = await Promise.all([
    db.findTeamsArchiveConversations({ user: userId, graphChatId: resolvedGraphChatId }, { limit: 1 }),
    includeRecentMessages
      ? db.findTeamsArchiveMessages(
          {
            user: userId,
            graphChatId: resolvedGraphChatId,
            ...buildHumanReadableMessageFilter(),
          },
          { limit, sort: { sentDateTime: -1, createdAt: -1 } },
        )
      : Promise.resolve([]),
    includeSystem
      ? db.findTeamsArchiveMessages(
          {
            user: userId,
            graphChatId: resolvedGraphChatId,
            isSystemLikeMessage: true,
          },
          { limit, sort: { sentDateTime: -1, createdAt: -1 } },
        )
      : Promise.resolve([]),
  ]);
  const conversation = conversations[0] || null;
  const selectedConversation = buildSelectedConversation(conversation, {
    selectionReason: 'chatId',
    identityConfidence: 'high',
  });
  const diagnosis = buildActivityDiagnosis(conversation || {});

  return {
    retrievalMode: 'conversation_activity_diagnostics',
    resolved: true,
    chatId,
    graphChatId: resolvedGraphChatId,
    ...(selectedConversation ? { selectedConversation } : {}),
    raw: {
      lastMessageAt: conversation?.lastMessageAt || null,
      lastMeaningfulMessageAt: conversation?.lastMeaningfulMessageAt || null,
      lastHumanMessageAt: conversation?.lastHumanMessageAt || null,
      lastSystemMessageAt: conversation?.lastSystemMessageAt || null,
      lastSyncedAt: conversation?.lastSyncedAt || null,
      lastMessageSyncAt: conversation?.lastMessageSyncAt || null,
      sourceUpdatedAt: conversation?.sourceUpdatedAt || null,
      messageCount: conversation?.messageCount || 0,
      humanMessageCount: conversation?.humanMessageCount || 0,
      meaningfulMessageCount: conversation?.meaningfulMessageCount || 0,
      systemMessageCount: conversation?.systemMessageCount || 0,
      emptyMessageCount: conversation?.emptyMessageCount || 0,
      participantDegraded: Boolean(conversation?.participantDegraded),
      participantMetadataSource: conversation?.participantMetadataSource || 'unknown',
      participantConfidence: conversation?.participantConfidence || 'low',
      participants: mapCompactParticipants(conversation?.participants || [], 12),
      syncStatus: conversation?.syncStatus || '',
      syncCursorPresent: Boolean(conversation?.syncCursor),
    },
    diagnosis,
    explanation: buildActivityExplanation(conversation || {}, diagnosis),
    newestHumanReadableMessages: newestHumanMessages.map((message) =>
      mapCompactMessageResult(message, conversation, { includeReplyToId: true, excerptLimit: 320 }),
    ),
    newestSystemMessages: newestSystemMessages.map((message) =>
      mapCompactMessageResult(message, conversation, { includeReplyToId: true, excerptLimit: 320 }),
    ),
  };
}

function buildSenderIdentityKey(message = {}) {
  return [
    String(message.fromUserId || '').trim(),
    String(message.fromEmail || '').trim().toLowerCase(),
    String(message.fromDisplayName || '').trim().toLowerCase(),
  ].join('|');
}

function summarizeSenderIdentityMessages(messages = [], senderClauses = []) {
  const groups = new Map();
  for (const message of messages) {
    const key = buildSenderIdentityKey(message);
    const existing = groups.get(key) || {
      fromUserId: message.fromUserId || '',
      fromEmail: message.fromEmail || '',
      fromDisplayName: message.fromDisplayName || '',
      count: 0,
      invalidEmail: false,
      senderMatch: messageSenderMatches(message, senderClauses),
      examples: [],
    };

    existing.count += 1;
    existing.invalidEmail =
      existing.invalidEmail || String(message.fromEmail || '').trim().toLowerCase() === 'aaduser';
    if (existing.examples.length < 3) {
      existing.examples.push({
        id: message._id?.toString?.() || message.id || '',
        graphMessageId: message.graphMessageId || '',
        graphChatId: message.graphChatId || '',
        sentDateTime: message.sentDateTime || null,
        excerpt: truncateText(message.bodyPreview || message.summary || message.bodyText || '', 180),
      });
    }

    groups.set(key, existing);
  }

  return [...groups.values()].sort((a, b) => b.count - a.count);
}

function buildRecommendedSenderFilters(identityCandidates = []) {
  const highConfidence = identityCandidates.find(
    (candidate) => candidate.fromUserId || (candidate.fromEmail && !candidate.invalidEmail),
  );
  const fallback = identityCandidates[0] || {};

  return {
    fromUserId: highConfidence?.fromUserId || '',
    fromEmail: highConfidence?.invalidEmail ? '' : highConfidence?.fromEmail || '',
    fromDisplayName: fallback?.fromDisplayName || '',
    useDisplayNameOnlyAsFallback: Boolean(!highConfidence?.fromUserId && !highConfidence?.fromEmail && fallback?.fromDisplayName),
  };
}

async function senderIdentityReport(user, options = {}) {
  assertEnabled();
  const userId = user?.id || user?._id?.toString();
  const senderScope = String(options.senderScope || 'me').trim();
  const normalizedSenderScope = ['me', 'person'].includes(senderScope) ? senderScope : 'me';
  const chatId = String(options.chatId || getPriorGraphChatIdForFollowUp(options) || '').trim();
  const senderClauses =
    normalizedSenderScope === 'me'
      ? getUserSenderClauses(user)
      : buildPersonSenderClauses({
          senderName: options.personName || options.senderName,
          senderEmail: options.personEmail || options.senderEmail,
          senderUserId: options.senderUserId,
        });
  let resolvedGraphChatId = null;
  if (chatId) {
    resolvedGraphChatId = await resolveConversationGraphChatId(userId, chatId);
  }

  const personRegexes =
    normalizedSenderScope === 'person'
      ? [options.personName || options.senderName, options.personEmail || options.senderEmail]
          .map((value) => buildSearchRegex(String(value || '').trim()))
          .filter(Boolean)
      : [];
  const filter = {
    user: userId,
    ...(resolvedGraphChatId ? { graphChatId: resolvedGraphChatId } : {}),
    ...(senderClauses.length > 0
      ? { $or: senderClauses }
      : personRegexes.length > 0
        ? {
            $or: personRegexes.flatMap((regex) => [
              { fromDisplayName: regex },
              { fromEmail: regex },
            ]),
          }
        : {}),
  };
  const messages = await db.findTeamsArchiveMessages(filter, {
    limit: 5000,
    sort: { sentDateTime: -1, createdAt: -1 },
  });
  const identityCandidates = summarizeSenderIdentityMessages(messages, senderClauses);
  const recommendedSenderFilters = buildRecommendedSenderFilters(identityCandidates);
  const warnings = [
    ...(identityCandidates.some((candidate) => candidate.invalidEmail)
      ? ['invalid_fromEmail_values_detected']
      : []),
    ...(identityCandidates.length === 0 ? ['no_sender_identities_observed_for_scope'] : []),
    ...(recommendedSenderFilters.useDisplayNameOnlyAsFallback
      ? ['display_name_only_matching_is_lower_confidence']
      : []),
  ];

  return {
    retrievalMode: 'sender_identity_report',
    senderScope: normalizedSenderScope,
    currentUserIdentity: {
      id: userId,
      openidId: user?.openidId || '',
      email: user?.email || '',
      name: user?.name || '',
      username: user?.username || '',
    },
    chatId: chatId || undefined,
    graphChatId: resolvedGraphChatId || undefined,
    identityCandidates,
    recommendedSenderFilters,
    confidence:
      recommendedSenderFilters.fromUserId || recommendedSenderFilters.fromEmail
        ? 'high'
        : recommendedSenderFilters.fromDisplayName
          ? 'medium'
          : 'low',
    warnings,
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
  const selectedConversation = buildSelectedConversation(conversation, {
    selectionReason: 'message_graphChatId',
    identityConfidence: 'high',
  });

  return {
    retrievalMode: 'message_body',
    resolved: true,
    guidance:
      'This returns the full archived message text for one specific message. Use it when a preview was truncated and exact wording matters.',
    ...(selectedConversation ? { selectedConversation } : {}),
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
  const selectedConversation = buildSelectedConversation(conversation, {
    selectionReason: 'chatId',
    identityConfidence: 'high',
  });

  if (!anchorMessage) {
    const recentMessages = await db.findTeamsArchiveMessages(
      { user: userId, graphChatId: resolvedGraphChatId },
      { limit: fallbackLimit, sort: { sentDateTime: -1, createdAt: -1 } },
    );

    return {
      retrievalMode: 'message_window',
      chatId,
      graphChatId: resolvedGraphChatId,
      ...(selectedConversation ? { selectedConversation } : {}),
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
    ...(selectedConversation ? { selectedConversation } : {}),
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
    evidenceBudget: buildEvidenceBudget(options, {
      fullBodiesReturned: 0,
      humanReadableReturned: messages.filter((message) => isMeaningfulArchiveMessage(message)).length,
      previewOnlyReturned: messages.length,
      conversationsScoped: resolvedGraphChatId ? 1 : conversationIds.length,
      totalMessagesConsidered: messages.length,
    }),
    results: messages.map((message) =>
      mapCompactMessageResult(message, conversationMap.get(message.graphChatId), {
        excerptLimit: 260,
      }),
    ),
  };
}

async function recentMessages(user, options = {}) {
  assertEnabled();
  const intent = classifyRetrievalIntent({ ...options, senderScope: 'me' }, { mode: 'recent' });
  let memoryResults = null;
  let memorySearchError = null;
  let memorySearched = false;

  if (typeof searchTeamsMemoryChunks === 'function') {
    try {
      memorySearched = true;
      memoryResults = await searchTeamsMemoryChunks(user, {
        query: options.query,
        limit: options.limit,
        daysBack: options.daysBack,
        senderScope: 'me',
        sortBy: 'recent',
      });
    } catch (error) {
      memorySearchError = error;
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

  const archiveResults = messages.map((message) =>
    mapCompactMessageResult(message, conversationMap.get(message.graphChatId), {
      excerptLimit: 220,
    }),
  );
  const unionDecision = buildArchiveUnionDecision({
    intent,
    memoryResults,
    requestedLimit: limit,
  });
  const fallbackReasons = [];
  if (memorySearchError) {
    fallbackReasons.push('memory_error');
  }
  fallbackReasons.push(...unionDecision.reasons);

  const finalResults =
    memoryResults && hasNonEmptyMemoryResults(memoryResults)
      ? dedupeMergedRetrievalResults({
          archiveResults: unionDecision.runArchiveUnion ? archiveResults : [],
          memoryResults: memoryResults.results,
          limit,
        })
      : archiveResults.slice(0, limit);

  return {
    retrievalMode: 'recent_message_previews',
    daysBack,
    query: query || undefined,
    guidance:
      'These are compact previews of recent messages sent by the signed-in user. Use get_messages_window for local context around one result.',
    evidenceBudget: buildEvidenceBudget(
      { ...options, action: 'recent_messages' },
      {
        fullBodiesReturned: 0,
        humanReadableReturned: archiveResults.length,
        previewOnlyReturned: finalResults.length,
        conversationsScoped: conversationIds.length === 1 ? 1 : conversationIds.length,
        totalMessagesConsidered: archiveResults.length,
      },
    ),
    trace: {
      detectedIntent: intent,
      filtersApplied: {
        senderScope: 'me',
        daysBack,
        query: Boolean(query),
      },
      memorySearched,
      memoryResultCount: Array.isArray(memoryResults?.results) ? memoryResults.results.length : 0,
      archiveUnionRan: Boolean(memoryResults && unionDecision.runArchiveUnion),
      archiveResultCount: archiveResults.length,
      fullBodyEscalationRan: false,
      conversationDossierRan: false,
      dedupedFinalResultCount: finalResults.length,
      fallbackReasons,
    },
    resultCount: finalResults.length,
    results: finalResults,
  };
}

async function advancedSearchMessages(user, options = {}) {
  assertEnabled();
  const intent = classifyRetrievalIntent(options, { mode: 'advanced' });
  const topic = String(options.topic || options.query || '').trim();
  const senderScope = String(options.senderScope || 'any').trim();
  const chatType = String(options.chatType || 'any').trim();
  const limit = clampInteger(options.limit, Math.min(getTeamsArchiveConfig().defaultSearchLimit, 4), {
    max: 6,
  });
  const daysBack = options.daysBack
    ? clampInteger(options.daysBack, 30, { min: 1, max: 3650 })
    : undefined;

  let memoryResults = null;
  let memorySearchError = null;
  let memorySearched = false;
  if (typeof searchTeamsMemoryChunks === 'function') {
    try {
      memorySearched = true;
      memoryResults = await searchTeamsMemoryChunks(user, options);
    } catch (error) {
      memorySearchError = error;
      logger.warn('[TeamsArchiveService] Enterprise memory advanced retrieval failed, falling back', {
        userId: user?.id || user?._id?.toString?.(),
        error: error?.message || error,
      });
    }
  }

  const unionDecision = buildArchiveUnionDecision({
    intent,
    memoryResults,
    requestedLimit: limit,
  });
  const archiveResult = await runArchiveAdvancedMessageSearch(user, {
    ...options,
    limit: Math.min(limit * 3, 12),
  });
  const archiveResults = archiveResult.results || [];
  const fallbackReasons = [];
  if (memorySearchError) {
    fallbackReasons.push('memory_error');
  }
  fallbackReasons.push(...unionDecision.reasons);

  const finalResults =
    memoryResults && hasNonEmptyMemoryResults(memoryResults)
      ? dedupeMergedRetrievalResults({
          archiveResults: unionDecision.runArchiveUnion ? archiveResults : [],
          memoryResults: memoryResults.results,
          limit,
        })
      : archiveResults.slice(0, limit);

  const resolvedConversation = resolveConversationFromSearchResults(
    finalResults,
    archiveResult.resolvedConversation,
  );
  const selectedConversation =
    archiveResult.selectedConversation ||
    (resolvedConversation?.graphChatId
      ? {
          archiveConversationId: resolvedConversation.id || '',
          graphChatId: resolvedConversation.graphChatId,
          topic: resolvedConversation.topic || '',
          chatType: resolvedConversation.chatType || '',
          lastMessageAt: resolvedConversation.lastMessageAt || null,
          lastMeaningfulMessageAt: resolvedConversation.lastMeaningfulMessageAt || null,
          selectionReason: options.chatId ? 'chatId' : 'single_result',
          identityConfidence: options.chatId ? 'high' : 'medium',
        }
      : null);
  const identityChangeWarning = buildIdentityChangeWarning({
    priorGraphChatId: options.priorGraphChatId,
    priorTopic: options.priorTopic,
    selectedConversation,
  });

  let conversationDossier = null;
  let conversationDossierRan = false;
  if (
    resolvedConversation?.graphChatId &&
    (intent.completenessSensitive ||
      intent.participantScoped ||
      intent.oneOnOneQuery)
  ) {
    conversationDossier = await getConversationDossier(user, {
      chatId: resolvedConversation.graphChatId,
      query: options.query,
      topic: options.topic,
      daysBack: options.daysBack,
      limit: Math.min(limit, 4),
      participants: options.participants,
      chatType: options.chatType,
    });
    conversationDossierRan = Boolean(conversationDossier?.resolved);
    if (conversationDossierRan) {
      fallbackReasons.push('conversation_dossier_escalation');
    }
  }

  let fullBodies = [];
  const shouldRunFullBodyEscalation =
    intent.exactnessSensitive ||
    intent.fullBodyDetailsQuery ||
    ((intent.participantScoped || intent.oneOnOneQuery) &&
      finalResults.length > 0 &&
      arePreviewsLikelyInsufficient(finalResults));

  if (shouldRunFullBodyEscalation && finalResults.length > 0) {
    fullBodies = await buildMessageBodyEscalations(user, options, finalResults);
    if (fullBodies.length > 0) {
      fallbackReasons.push('full_body_escalation');
    }
  }

  return {
    retrievalMode: 'advanced_message_previews',
    topic: topic || undefined,
    ...(archiveResult?.chatId ? { chatId: archiveResult.chatId } : {}),
    ...(archiveResult?.graphChatId ? { graphChatId: archiveResult.graphChatId } : {}),
    senderScope,
    chatType,
    daysBack,
    participants: toArray(options.participants).filter(Boolean),
    guidance:
      resolvedConversation
        ? 'These results were gathered with recall-safe union retrieval. When completeness matters, rely on the conversation dossier and full-body expansions before answering.'
        : 'These results were gathered with recall-safe union retrieval across enterprise memory and the raw Teams archive. Use the dossier or full-body expansions before answering completeness-sensitive questions.',
    ...(resolvedConversation ? { resolvedConversation } : {}),
    ...(selectedConversation ? { selectedConversation } : {}),
    ...(identityChangeWarning ? { identityWarning: identityChangeWarning } : {}),
    ...(conversationDossierRan ? { conversationDossier } : {}),
    ...(fullBodies.length > 0 ? { fullBodies } : {}),
    evidenceBudget: buildEvidenceBudget(options, {
      fullBodiesReturned: fullBodies.length,
      humanReadableReturned: finalResults.length,
      previewOnlyReturned: finalResults.length,
      conversationsScoped: selectedConversation ? 1 : 0,
      totalMessagesConsidered:
        conversationDossier?.completeness?.totalMessagesInScope || archiveResults.length || finalResults.length,
      identityWarning: Boolean(identityChangeWarning),
    }),
    trace: {
      detectedIntent: intent,
      filtersApplied: {
        senderScope,
        chatType,
        daysBack: Boolean(daysBack),
        participants: toArray(options.participants).filter(Boolean),
        topic: Boolean(topic),
      },
      memorySearched,
      memoryResultCount: Array.isArray(memoryResults?.results) ? memoryResults.results.length : 0,
      archiveUnionRan: Boolean(memoryResults ? unionDecision.runArchiveUnion : true),
      archiveResultCount: archiveResults.length,
      fullBodyEscalationRan: fullBodies.length > 0,
      conversationDossierRan,
      dedupedFinalResultCount: finalResults.length,
      fallbackReasons,
    },
    resultCount: finalResults.length,
    results: finalResults,
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
  const scopedStats = summarizeMessageSearchability(scopedMessages.length > 0 ? scopedMessages : messages);
  const conversationWarnings = buildConversationWarningFlags({
    ...conversation,
    messageCount: messages.length,
    meaningfulMessageCount: scopedStats.meaningfulMessageCount,
    systemMessageCount: scopedStats.systemMessageCount,
    emptyMessageCount: scopedStats.emptyMessageCount,
    lastMeaningfulMessageAt: scopedStats.lastMeaningfulMessageAt,
    lastHumanMessageAt: scopedStats.lastHumanMessageAt,
    lastSystemMessageAt: scopedStats.lastSystemMessageAt,
  });
  const lowEvidence =
    messages.length < 5 ||
    scopedStats.meaningfulMessageCount < 3 ||
    conversationWarnings.noMeaningfulMessages ||
    conversationWarnings.systemOnlyRecentActivity;
  const confidence = lowEvidence ? 'low' : scopedMessages.length > 0 ? 'medium' : 'medium';
  const selectedConversation = buildSelectedConversation(conversation, {
    selectionReason: 'chatId',
    identityConfidence: 'high',
  });

  return {
    retrievalMode: 'conversation_summary',
    chatId,
    graphChatId: resolvedGraphChatId,
    ...(selectedConversation ? { selectedConversation } : {}),
    topic: conversation?.topic || '',
    chatType: conversation?.chatType || '',
    participants: mapCompactParticipants(conversation?.participants || []),
    daysBack,
    query: query || undefined,
    totalMessages: messages.length,
    matchedMessages: scopedMessages.length,
    firstMessageAt,
    lastMessageAt,
    lastMeaningfulMessageAt: scopedStats.lastMeaningfulMessageAt || conversation?.lastMeaningfulMessageAt || null,
    rawLastMessageAt: lastMessageAt,
    topSenders,
    trust: {
      confidence,
      evidence: [
        {
          statement: `Summary is based on ${scopedMessages.length || messages.length} human-readable/message records loaded from this resolved Teams chat.`,
          sourceMessageIds: highlights.map((message) => message.graphMessageId).filter(Boolean),
        },
        {
          statement: `The chat identity is graphChatId=${resolvedGraphChatId}.`,
          sourceMessageIds: [],
        },
      ],
      inferences: [],
      unknowns: [
        ...(lowEvidence
          ? [
              'Conversation purpose, importance, and cadence are not inferred from the title alone.',
              'Low message count, unknown senders, empty messages, or system-heavy activity may limit summary quality.',
            ]
          : ['Meeting purpose and importance are only supported where message evidence exists.']),
      ],
      warnings: conversationWarnings,
    },
    evidenceBudget: buildEvidenceBudget(options, {
      fullBodiesReturned: 0,
      humanReadableReturned: scopedMessages.length || messages.length,
      previewOnlyReturned: highlights.length,
      conversationsScoped: 1,
      totalMessagesConsidered: messages.length,
    }),
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
  backfillConversationRecency,
  listConversations,
  recentMeetingChats,
  getConversationDossier,
  listConversationMessages,
  conversationRecentMessages,
  conversationSenderMessages,
  conversationActivityDiagnostics,
  senderIdentityReport,
  getMessageBody,
  getMessagesWindow,
  searchMessages,
  recentMessages,
  advancedSearchMessages,
  summarizeConversation,
};
