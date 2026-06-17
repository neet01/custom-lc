const { isEnabled } = require('@librechat/api');
const { randomUUID } = require('crypto');
const { logger, runAsSystem } = require('@librechat/data-schemas');
const db = require('~/models');
const { getUserPluginAuthValue } = require('~/server/services/PluginService');

let projectSlackArchiveSyncToMemory = null;

try {
  ({ projectSlackArchiveSyncToMemory } = require('~/server/services/EnterpriseMemory/slackProjection'));
} catch (error) {
  if (error?.code !== 'MODULE_NOT_FOUND') {
    throw error;
  }
}

const DEFAULT_SLACK_API_BASE_URL = 'https://slack-gov.com/api';
const DEFAULT_SYNC_CONVERSATION_LIMIT = 100;
const DEFAULT_SYNC_MESSAGES_PER_CONVERSATION = 500;
const DEFAULT_SEARCH_LIMIT = 25;
const DEFAULT_SYNC_STALE_MINUTES = 45;
const DEFAULT_MAX_CONCURRENT_SYNCS = 1;
const DEFAULT_RETRY_ATTEMPTS = 5;
const DEFAULT_RETRY_BASE_MS = 1000;
const DEFAULT_RETRY_MAX_MS = 60000;
const SLACK_ARCHIVE_PLUGIN_KEY = 'slack_archive';
const USER_ACCESS_TOKEN_FIELD = 'SLACK_ARCHIVE_USER_ACCESS_TOKEN';

class SlackArchiveServiceError extends Error {
  constructor(message, status = 500, details) {
    super(message);
    this.name = 'SlackArchiveServiceError';
    this.status = status;
    this.details = details;
  }
}

class SlackArchiveSyncCancelledError extends Error {
  constructor(message = 'Slack archive sync cancelled by user') {
    super(message);
    this.name = 'SlackArchiveSyncCancelledError';
  }
}

function isSlackArchiveEnabled() {
  return isEnabled(process.env.SLACK_ARCHIVE_ENABLED);
}

function assertEnabled() {
  if (!isSlackArchiveEnabled()) {
    throw new SlackArchiveServiceError('Slack archive is not enabled', 403);
  }
}

function clampPositiveInt(value, fallback, { min = 1, max = 1000 } = {}) {
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) {
    return fallback;
  }

  return Math.min(Math.max(Math.floor(parsed), min), max);
}

function escapeRegex(value) {
  return String(value || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function buildSearchRegex(value) {
  const normalized = String(value || '').trim().replace(/\s+/g, ' ');
  if (!normalized) {
    return null;
  }
  return new RegExp(escapeRegex(normalized).replace(/\\ /g, '\\s+'), 'i');
}

function normalizeSlackApiBaseUrl(baseUrl = DEFAULT_SLACK_API_BASE_URL) {
  return String(baseUrl || DEFAULT_SLACK_API_BASE_URL)
    .trim()
    .replace(/\/+$/, '');
}

function getSlackArchiveConfig() {
  const explicitRedirectUri = String(process.env.SLACK_ARCHIVE_REDIRECT_URI || '').trim();
  const domainServer = String(process.env.DOMAIN_SERVER || '').trim().replace(/\/+$/, '');

  return {
    enabled: isSlackArchiveEnabled(),
    apiBaseUrl: normalizeSlackApiBaseUrl(
      process.env.SLACK_ARCHIVE_API_BASE_URL || DEFAULT_SLACK_API_BASE_URL,
    ),
    clientId: String(process.env.SLACK_ARCHIVE_CLIENT_ID || '').trim(),
    clientSecret: String(process.env.SLACK_ARCHIVE_CLIENT_SECRET || '').trim(),
    redirectUri: explicitRedirectUri || (domainServer ? `${domainServer}/api/slack-archive/oauth/callback` : ''),
    userScopes:
      process.env.SLACK_ARCHIVE_USER_SCOPES ||
      'channels:history,groups:history,im:history,mpim:history,channels:read,groups:read,im:read,mpim:read,users:read,users:read.email',
    botScopes:
      process.env.SLACK_ARCHIVE_BOT_SCOPES ||
      'app_mentions:read,channels:history,groups:history,im:history,mpim:history,chat:write',
    syncConversationLimit: clampPositiveInt(
      process.env.SLACK_ARCHIVE_MAX_SYNC_CONVERSATIONS,
      DEFAULT_SYNC_CONVERSATION_LIMIT,
      { max: 10000 },
    ),
    syncMessagesPerConversation: clampPositiveInt(
      process.env.SLACK_ARCHIVE_MAX_MESSAGES_PER_CONVERSATION,
      DEFAULT_SYNC_MESSAGES_PER_CONVERSATION,
      { max: 5000 },
    ),
    searchLimit: clampPositiveInt(process.env.SLACK_ARCHIVE_SEARCH_LIMIT, DEFAULT_SEARCH_LIMIT, {
      max: 100,
    }),
    syncStaleMinutes: clampPositiveInt(
      process.env.SLACK_ARCHIVE_SYNC_STALE_MINUTES,
      DEFAULT_SYNC_STALE_MINUTES,
      { max: 24 * 60 },
    ),
    maxConcurrentSyncs: clampPositiveInt(
      process.env.SLACK_ARCHIVE_MAX_CONCURRENT_SYNCS,
      DEFAULT_MAX_CONCURRENT_SYNCS,
      { min: 0, max: 10 },
    ),
    retryAttempts: clampPositiveInt(process.env.SLACK_ARCHIVE_RETRY_ATTEMPTS, DEFAULT_RETRY_ATTEMPTS, {
      min: 0,
      max: 10,
    }),
    retryBaseMs: clampPositiveInt(process.env.SLACK_ARCHIVE_RETRY_BASE_MS, DEFAULT_RETRY_BASE_MS, {
      min: 0,
      max: 5 * 60 * 1000,
    }),
    retryMaxMs: clampPositiveInt(process.env.SLACK_ARCHIVE_RETRY_MAX_MS, DEFAULT_RETRY_MAX_MS, {
      min: 0,
      max: 60 * 60 * 1000,
    }),
  };
}

function getUserId(user) {
  return user?.id || user?._id?.toString();
}

function toIso(value) {
  return value instanceof Date ? value.toISOString() : value || null;
}

function formatSyncJob(syncJob) {
  if (!syncJob) {
    return null;
  }

  return {
    id: syncJob._id?.toString?.() || syncJob.id || null,
    status: syncJob.status,
    mode: syncJob.mode,
    phase: syncJob.phase || null,
    checkpoint: syncJob.checkpoint || undefined,
    stats: syncJob.stats || undefined,
    requestedConversationLimit: syncJob.requestedConversationLimit || 0,
    requestedMessagesPerConversation: syncJob.requestedMessagesPerConversation || 0,
    discoveredConversationCount: syncJob.discoveredConversationCount || 0,
    processedConversationCount: syncJob.processedConversationCount || 0,
    skippedConversationCount: syncJob.skippedConversationCount || 0,
    conversationCount: syncJob.conversationCount || 0,
    messageCount: syncJob.messageCount || 0,
    errorMessage: syncJob.errorMessage || undefined,
    startedAt: toIso(syncJob.startedAt),
    completedAt: toIso(syncJob.completedAt),
  };
}

function formatConversation(conversation) {
  return {
    id: conversation._id?.toString?.() || conversation.id,
    slackConversationId: conversation.slackConversationId,
    teamId: conversation.teamId || '',
    enterpriseId: conversation.enterpriseId || '',
    conversationType: conversation.conversationType || '',
    name: conversation.name || '',
    topic: conversation.topic || '',
    purpose: conversation.purpose || '',
    isArchived: Boolean(conversation.isArchived),
    isSlackConnect: Boolean(conversation.isSlackConnect),
    participantCount: Array.isArray(conversation.participants) ? conversation.participants.length : 0,
    messageCount: conversation.messageCount || 0,
    lastMessageAt: toIso(conversation.lastMessageAt),
    lastMeaningfulMessageAt: toIso(conversation.lastMeaningfulMessageAt),
    syncStatus: conversation.syncStatus || 'pending',
  };
}

function formatMessage(message) {
  return {
    id: message._id?.toString?.() || message.id,
    slackConversationId: message.slackConversationId,
    slackMessageTs: message.slackMessageTs,
    threadTs: message.threadTs || null,
    slackUserId: message.slackUserId || null,
    botId: message.botId || null,
    displayName: message.displayName || message.username || '',
    subtype: message.subtype || '',
    text: message.text || '',
    normalizedText: message.normalizedText || '',
    sentAt: toIso(message.sentAt),
    editedAt: toIso(message.editedAt),
    deletedAt: toIso(message.deletedAt),
    replyCount: message.replyCount || 0,
    reactions: Array.isArray(message.reactions) ? message.reactions : [],
    files: Array.isArray(message.files) ? message.files : [],
  };
}

function formatProjectionJob(job) {
  if (!job) {
    return null;
  }

  return {
    id: job._id?.toString?.() || job.id || null,
    status: job.status,
    jobType: job.jobType || null,
    sourceRecordType: job.sourceRecordType || null,
    sourceRecordId: job.sourceRecordId || null,
    stats: job.stats || undefined,
    errorMessage: job.errorMessage || undefined,
    startedAt: toIso(job.startedAt),
    completedAt: toIso(job.completedAt),
  };
}

function getLeaseDurationMs() {
  return getSlackArchiveConfig().syncStaleMinutes * 60 * 1000;
}

function getLeaseExpiryDate() {
  return new Date(Date.now() + getLeaseDurationMs());
}

function getUserLeaseKey(userId) {
  return `user:${userId}`;
}

function getSlotLeaseKey(slotNumber) {
  return `slot:${slotNumber}`;
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function parseRetryAfterMs(response) {
  const retryAfter = response.headers.get('retry-after');
  if (!retryAfter) {
    return null;
  }

  const seconds = Number(retryAfter);
  if (Number.isFinite(seconds) && seconds >= 0) {
    return seconds * 1000;
  }

  const timestamp = Date.parse(retryAfter);
  if (Number.isFinite(timestamp)) {
    return Math.max(0, timestamp - Date.now());
  }

  return null;
}

function calculateBackoffMs(attempt) {
  const config = getSlackArchiveConfig();
  const exponential = config.retryBaseMs * 2 ** Math.max(attempt - 1, 0);
  return Math.min(config.retryMaxMs, exponential);
}

function queueSlackProjection(params) {
  if (typeof projectSlackArchiveSyncToMemory !== 'function') {
    return null;
  }

  const queuedAt = new Date();
  void projectSlackArchiveSyncToMemory(params).catch((error) => {
    logger.error('[SlackArchiveService] Background Slack projection failed', {
      userId: params?.userId,
      syncJobId: params?.syncJobId,
      error: error?.message || error,
    });
  });

  return {
    status: 'queued',
    queuedAt: queuedAt.toISOString(),
    requestedConversationCount: Array.isArray(params?.slackConversationIds)
      ? params.slackConversationIds.length
      : 0,
  };
}

async function getSlackConnectionForUser(user) {
  assertEnabled();
  const userId = getUserId(user);
  if (!userId) {
    throw new SlackArchiveServiceError('Authenticated user context is required.', 401);
  }

  const [userAccessToken, identityLink] = await Promise.all([
    getUserPluginAuthValue(userId, USER_ACCESS_TOKEN_FIELD, false, SLACK_ARCHIVE_PLUGIN_KEY),
    typeof db.findSlackIdentityLink === 'function'
      ? db.findSlackIdentityLink({ user: userId, status: 'linked' })
      : Promise.resolve(null),
  ]);

  const workspaceInstall =
    typeof db.findSlackWorkspaceInstall === 'function'
      ? await db.findSlackWorkspaceInstall(
          identityLink?.teamId || identityLink?.enterpriseId
            ? {
                ...(identityLink?.teamId ? { teamId: identityLink.teamId } : {}),
                ...(identityLink?.enterpriseId ? { enterpriseId: identityLink.enterpriseId } : {}),
                status: 'active',
              }
            : { status: 'active' },
        )
      : null;

  if (!workspaceInstall?.botAccessToken) {
    throw new SlackArchiveServiceError(
      'No active GovSlack workspace install is available for Slack archive sync.',
      403,
    );
  }

  if (!userAccessToken) {
    throw new SlackArchiveServiceError(
      'No Slack user token is available for the current Cortex user. Reconnect GovSlack first.',
      401,
    );
  }

  return {
    userId,
    userAccessToken,
    identityLink,
    workspaceInstall,
    teamId: identityLink?.teamId || workspaceInstall.teamId || '',
    enterpriseId: identityLink?.enterpriseId || workspaceInstall.enterpriseId || '',
  };
}

async function acquireSyncLeases(userId) {
  const ownerToken = randomUUID();
  const userLeaseKey = getUserLeaseKey(userId);
  const leaseRecord = {
    leaseKey: userLeaseKey,
    leaseType: 'user',
    ownerToken,
    user: userId,
    leaseExpiresAt: getLeaseExpiryDate(),
    lastHeartbeatAt: new Date(),
  };

  const userLease = await db.acquireSlackArchiveSyncLease(leaseRecord);
  if (!userLease || userLease.ownerToken !== ownerToken) {
    return null;
  }

  const { maxConcurrentSyncs } = getSlackArchiveConfig();
  if (maxConcurrentSyncs === 0) {
    return { ownerToken, slotNumber: null };
  }

  for (let slotNumber = 0; slotNumber < maxConcurrentSyncs; slotNumber += 1) {
    const slotLease = await db.acquireSlackArchiveSyncLease({
      leaseKey: getSlotLeaseKey(slotNumber),
      leaseType: 'slot',
      ownerToken,
      user: userId,
      leaseExpiresAt: getLeaseExpiryDate(),
      lastHeartbeatAt: new Date(),
    });

    if (slotLease?.ownerToken === ownerToken) {
      return { ownerToken, slotNumber };
    }
  }

  await db.releaseSlackArchiveSyncLease(userLeaseKey, ownerToken);
  return null;
}

async function refreshSyncLeases(userId, leaseState) {
  if (!leaseState?.ownerToken) {
    return;
  }

  const leaseExpiresAt = getLeaseExpiryDate();
  await db.refreshSlackArchiveSyncLease(getUserLeaseKey(userId), leaseState.ownerToken, leaseExpiresAt);
  if (leaseState.slotNumber !== null && leaseState.slotNumber !== undefined) {
    await db.refreshSlackArchiveSyncLease(
      getSlotLeaseKey(leaseState.slotNumber),
      leaseState.ownerToken,
      leaseExpiresAt,
    );
  }
}

async function releaseSyncLeases(userId, leaseState) {
  if (!leaseState?.ownerToken) {
    return;
  }

  await Promise.all([
    db.releaseSlackArchiveSyncLease(getUserLeaseKey(userId), leaseState.ownerToken),
    leaseState.slotNumber !== null && leaseState.slotNumber !== undefined
      ? db.releaseSlackArchiveSyncLease(getSlotLeaseKey(leaseState.slotNumber), leaseState.ownerToken)
      : Promise.resolve(false),
  ]);
}

function buildSlackApiUrl(methodName, params = {}) {
  const baseUrl = `${getSlackArchiveConfig().apiBaseUrl}/`;
  const url = new URL(methodName.replace(/^\//, ''), baseUrl);
  for (const [key, value] of Object.entries(params)) {
    if (value === undefined || value === null || value === '') {
      continue;
    }

    if (Array.isArray(value)) {
      url.searchParams.set(key, value.join(','));
    } else {
      url.searchParams.set(key, String(value));
    }
  }
  return url;
}

async function slackApiRequest(methodName, token, params = {}) {
  const config = getSlackArchiveConfig();
  let lastError = null;

  for (let attempt = 0; attempt <= config.retryAttempts; attempt += 1) {
    const response = await fetch(buildSlackApiUrl(methodName, params), {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
      },
    });

    if (response.status === 429 && attempt < config.retryAttempts) {
      const retryAfterMs = parseRetryAfterMs(response) ?? calculateBackoffMs(attempt + 1);
      logger.warn('[SlackArchiveService] GovSlack rate limit encountered', {
        methodName,
        retryAfterMs,
        attempt: attempt + 1,
      });
      await sleep(retryAfterMs);
      continue;
    }

    let payload;
    try {
      payload = await response.json();
    } catch (error) {
      throw new SlackArchiveServiceError('Slack API returned a non-JSON response.', 502, {
        methodName,
        cause: error instanceof Error ? error.message : String(error),
      });
    }

    if (response.ok && payload?.ok === true) {
      return payload;
    }

    lastError = {
      status: response.status,
      error: payload?.error || payload?.message || response.statusText || 'unknown_error',
    };

    const isRetryable =
      response.status === 429 ||
      payload?.error === 'ratelimited' ||
      response.status === 500 ||
      response.status === 503 ||
      response.status === 504;

    if (isRetryable && attempt < config.retryAttempts) {
      const retryAfterMs = parseRetryAfterMs(response) ?? calculateBackoffMs(attempt + 1);
      await sleep(retryAfterMs);
      continue;
    }

    throw new SlackArchiveServiceError(`GovSlack API request failed for ${methodName}.`, 502, {
      methodName,
      ...lastError,
    });
  }

  throw new SlackArchiveServiceError(`GovSlack API request failed for ${methodName}.`, 502, lastError);
}

function mapConversationType(channel = {}) {
  if (channel.is_im) {
    return 'im';
  }
  if (channel.is_mpim) {
    return 'mpim';
  }
  if (channel.is_private) {
    return 'private_channel';
  }
  return 'public_channel';
}

function slackTsToDate(ts) {
  const parsed = Number.parseFloat(String(ts || ''));
  return Number.isFinite(parsed) ? new Date(parsed * 1000) : null;
}

function uniqueBy(array, keySelector) {
  const seen = new Set();
  const results = [];

  for (const item of array) {
    const key = keySelector(item);
    if (!key || seen.has(key)) {
      continue;
    }
    seen.add(key);
    results.push(item);
  }

  return results;
}

function extractMentions(text = '', userCache = new Map()) {
  const mentionRegex = /<@([A-Z0-9]+)>/gi;
  const mentions = [];
  let match;

  while ((match = mentionRegex.exec(text)) !== null) {
    const slackUserId = String(match[1] || '').trim();
    if (!slackUserId) {
      continue;
    }
    const profile = userCache.get(slackUserId);
    mentions.push({
      slackUserId,
      displayName:
        profile?.displayName || profile?.realName || profile?.username || profile?.email || slackUserId,
    });
  }

  return uniqueBy(mentions, (mention) => mention.slackUserId);
}

function normalizeSlackText(rawMessage = {}) {
  const pieces = [];
  if (rawMessage.text) {
    pieces.push(String(rawMessage.text));
  }

  for (const attachment of rawMessage.attachments || []) {
    if (attachment?.title) {
      pieces.push(String(attachment.title));
    }
    if (attachment?.text) {
      pieces.push(String(attachment.text));
    }
    if (attachment?.fallback) {
      pieces.push(String(attachment.fallback));
    }
  }

  return pieces
    .join('\n')
    .replace(/\s+/g, ' ')
    .trim();
}

function isSystemLikeSlackMessage(rawMessage = {}) {
  const subtype = String(rawMessage.subtype || '').trim();
  return [
    'channel_join',
    'channel_leave',
    'channel_topic',
    'channel_purpose',
    'channel_name',
    'channel_archive',
    'channel_unarchive',
    'message_deleted',
    'group_join',
    'group_leave',
  ].includes(subtype);
}

function materializeSlackMessage(rawMessage = {}) {
  if (rawMessage.subtype === 'message_changed' && rawMessage.message) {
    return {
      ...rawMessage.message,
      subtype: rawMessage.message.subtype || 'message_changed',
      edited: rawMessage.message.edited || rawMessage.edited,
      reply_count: rawMessage.message.reply_count || rawMessage.reply_count,
      reply_users: rawMessage.message.reply_users || rawMessage.reply_users,
      latest_reply: rawMessage.message.latest_reply || rawMessage.latest_reply,
    };
  }

  if (rawMessage.subtype === 'message_deleted') {
    return {
      ...rawMessage.previous_message,
      ts: rawMessage.deleted_ts || rawMessage.previous_message?.ts || rawMessage.ts,
      subtype: 'message_deleted',
      deleted_ts: rawMessage.deleted_ts || rawMessage.ts,
      text: '',
    };
  }

  return rawMessage;
}

function mapSlackFiles(files = []) {
  return (Array.isArray(files) ? files : []).map((file) => ({
    id: file?.id || '',
    name: file?.name || '',
    title: file?.title || '',
    mimetype: file?.mimetype || '',
    filetype: file?.filetype || '',
    prettyType: file?.pretty_type || '',
    urlPrivate: file?.url_private || '',
  }));
}

function mapSlackAttachments(attachments = []) {
  return (Array.isArray(attachments) ? attachments : []).map((attachment) => ({
    id: attachment?.id || '',
    title: attachment?.title || '',
    text: attachment?.text || '',
    fallback: attachment?.fallback || '',
    serviceName: attachment?.service_name || '',
    titleLink: attachment?.title_link || '',
  }));
}

function buildSlackParticipantFromProfile(profile = {}, slackUserId) {
  return {
    slackUserId,
    displayName:
      profile?.displayName || profile?.realName || profile?.username || profile?.email || slackUserId,
    realName: profile?.realName || '',
    username: profile?.username || '',
    email: profile?.email || '',
    isBot: Boolean(profile?.isBot),
    isAppUser: Boolean(profile?.isAppUser),
  };
}

async function getSlackUserProfile(token, slackUserId, userCache) {
  if (!slackUserId) {
    return null;
  }

  if (userCache.has(slackUserId)) {
    return userCache.get(slackUserId);
  }

  try {
    const payload = await slackApiRequest('users.info', token, { user: slackUserId });
    const user = payload?.user || {};
    const profile = {
      slackUserId,
      displayName:
        user?.profile?.display_name_normalized ||
        user?.profile?.display_name ||
        user?.real_name_normalized ||
        user?.real_name ||
        user?.name ||
        '',
      realName: user?.real_name || '',
      username: user?.name || '',
      email: user?.profile?.email || '',
      isBot: Boolean(user?.is_bot),
      isAppUser: Boolean(user?.is_app_user),
    };
    userCache.set(slackUserId, profile);
    return profile;
  } catch (error) {
    logger.warn('[SlackArchiveService] Failed to resolve Slack user profile', {
      slackUserId,
      message: error instanceof Error ? error.message : String(error),
    });
    const fallback = {
      slackUserId,
      displayName: slackUserId,
      realName: '',
      username: '',
      email: '',
      isBot: false,
      isAppUser: false,
    };
    userCache.set(slackUserId, fallback);
    return fallback;
  }
}

async function listSlackConversations(token, conversationLimit) {
  const conversations = [];
  let cursor = '';

  while (conversations.length < conversationLimit) {
    const payload = await slackApiRequest('conversations.list', token, {
      types: 'public_channel,private_channel,im,mpim',
      exclude_archived: false,
      limit: Math.min(200, conversationLimit - conversations.length),
      cursor,
    });

    conversations.push(...(payload?.channels || []));
    cursor = payload?.response_metadata?.next_cursor || '';

    if (!cursor || !(payload?.channels || []).length) {
      break;
    }
  }

  return conversations.slice(0, conversationLimit);
}

async function listConversationMembers(channel, token, userCache) {
  const participants = [];

  if (channel?.is_im && channel?.user) {
    const profile = await getSlackUserProfile(token, channel.user, userCache);
    if (profile) {
      participants.push(buildSlackParticipantFromProfile(profile, channel.user));
    }
    return participants;
  }

  let cursor = '';
  do {
    try {
      const payload = await slackApiRequest('conversations.members', token, {
        channel: channel.id,
        limit: 200,
        cursor,
      });

      for (const slackUserId of payload?.members || []) {
        const profile = await getSlackUserProfile(token, slackUserId, userCache);
        if (profile) {
          participants.push(buildSlackParticipantFromProfile(profile, slackUserId));
        }
      }

      cursor = payload?.response_metadata?.next_cursor || '';
    } catch (error) {
      logger.warn('[SlackArchiveService] Failed to list Slack conversation members', {
        channelId: channel?.id,
        message: error instanceof Error ? error.message : String(error),
      });
      break;
    }
  } while (cursor);

  return uniqueBy(participants, (participant) => participant.slackUserId);
}

async function fetchConversationHistory(channelId, token, messageLimit) {
  const messages = [];
  let cursor = '';

  while (messages.length < messageLimit) {
    const payload = await slackApiRequest('conversations.history', token, {
      channel: channelId,
      limit: Math.min(200, messageLimit - messages.length),
      cursor,
      inclusive: true,
    });

    messages.push(...(payload?.messages || []));
    cursor = payload?.response_metadata?.next_cursor || '';

    if (!cursor || !(payload?.messages || []).length) {
      break;
    }
  }

  return messages;
}

async function fetchThreadReplies(channelId, parentMessage, token, remainingBudget) {
  if (!parentMessage?.reply_count || remainingBudget <= 0 || !parentMessage?.ts) {
    return [];
  }

  const replies = [];
  let cursor = '';

  while (replies.length < remainingBudget) {
    const payload = await slackApiRequest('conversations.replies', token, {
      channel: channelId,
      ts: parentMessage.ts,
      limit: Math.min(200, remainingBudget - replies.length + 1),
      cursor,
      inclusive: true,
    });

    const pageReplies = (payload?.messages || []).filter((message) => message?.ts !== parentMessage.ts);
    replies.push(...pageReplies);
    cursor = payload?.response_metadata?.next_cursor || '';

    if (!cursor || pageReplies.length === 0) {
      break;
    }
  }

  return replies;
}

async function fetchConversationMessages(channel, token, messageLimit) {
  const baseMessages = await fetchConversationHistory(channel.id, token, messageLimit);
  const threadParents = baseMessages.filter((message) => Number(message?.reply_count || 0) > 0);
  const threadedReplies = [];

  for (const parentMessage of threadParents) {
    const remainingBudget = messageLimit - baseMessages.length - threadedReplies.length;
    if (remainingBudget <= 0) {
      break;
    }

    const replies = await fetchThreadReplies(channel.id, parentMessage, token, remainingBudget);
    threadedReplies.push(...replies);
  }

  return uniqueBy([...baseMessages, ...threadedReplies], (message) => message?.ts).slice(0, messageLimit);
}

async function normalizeSlackMessages({
  userId,
  teamId,
  enterpriseId,
  slackConversationId,
  rawMessages,
  token,
  userCache,
}) {
  const normalized = [];

  for (const rawMessage of rawMessages) {
    const materialized = materializeSlackMessage(rawMessage);
    const slackUserId = materialized?.user || '';
    const profile = slackUserId ? await getSlackUserProfile(token, slackUserId, userCache) : null;
    const text = materialized?.text || '';
    const normalizedText = normalizeSlackText(materialized);
    const subtype = String(materialized?.subtype || rawMessage?.subtype || '').trim();
    const deletedAt = materialized?.deleted_ts ? slackTsToDate(materialized.deleted_ts) : null;
    const isSystemLikeMessage = isSystemLikeSlackMessage({ ...materialized, subtype });
    const isChunkable = Boolean(normalizedText) && !isSystemLikeMessage && !deletedAt;

    normalized.push({
      user: userId,
      slackConversationId,
      slackMessageTs: String(materialized?.ts || rawMessage?.ts || '').trim(),
      teamId,
      enterpriseId,
      slackUserId: slackUserId || undefined,
      botId: materialized?.bot_id || undefined,
      username: profile?.username || materialized?.username || undefined,
      displayName:
        profile?.displayName ||
        materialized?.username ||
        materialized?.user_profile?.display_name ||
        materialized?.user_profile?.real_name ||
        undefined,
      subtype: subtype || undefined,
      text: text || '',
      normalizedText,
      threadTs: materialized?.thread_ts || materialized?.ts || undefined,
      parentUserId: materialized?.parent_user_id || undefined,
      replyCount: Number(materialized?.reply_count || 0),
      replyUsers: Array.isArray(materialized?.reply_users) ? materialized.reply_users : undefined,
      latestReplyTs: materialized?.latest_reply || undefined,
      reactions: Array.isArray(materialized?.reactions)
        ? materialized.reactions.map((reaction) => ({
            name: reaction?.name || '',
            count: Number(reaction?.count || 0),
            users: Array.isArray(reaction?.users) ? reaction.users : [],
          }))
        : undefined,
      mentions: extractMentions(text, userCache),
      attachments: mapSlackAttachments(materialized?.attachments),
      files: mapSlackFiles(materialized?.files),
      raw: rawMessage,
      sentAt: slackTsToDate(materialized?.ts),
      editedAt: materialized?.edited?.ts ? slackTsToDate(materialized.edited.ts) : null,
      deletedAt,
      normalizedTextLength: normalizedText.length,
      isSystemLikeMessage,
      isChunkable,
      skipChunkReason: isChunkable ? undefined : deletedAt ? 'deleted' : isSystemLikeMessage ? 'system_like' : 'empty_text',
    });
  }

  return normalized.filter((message) => message.slackMessageTs);
}

function buildConversationStats(messages = []) {
  let humanMessageCount = 0;
  let systemMessageCount = 0;
  let meaningfulMessageCount = 0;
  let lastMessageAt = null;
  let lastHumanMessageAt = null;
  let lastMeaningfulMessageAt = null;
  let lastSystemMessageAt = null;

  for (const message of messages) {
    const sentAt = message.sentAt instanceof Date ? message.sentAt : null;
    if (!lastMessageAt || (sentAt && sentAt > lastMessageAt)) {
      lastMessageAt = sentAt || lastMessageAt;
    }

    if (message.isSystemLikeMessage) {
      systemMessageCount += 1;
      if (!lastSystemMessageAt || (sentAt && sentAt > lastSystemMessageAt)) {
        lastSystemMessageAt = sentAt || lastSystemMessageAt;
      }
    } else {
      humanMessageCount += 1;
      if (!lastHumanMessageAt || (sentAt && sentAt > lastHumanMessageAt)) {
        lastHumanMessageAt = sentAt || lastHumanMessageAt;
      }
    }

    if (message.isChunkable) {
      meaningfulMessageCount += 1;
      if (!lastMeaningfulMessageAt || (sentAt && sentAt > lastMeaningfulMessageAt)) {
        lastMeaningfulMessageAt = sentAt || lastMeaningfulMessageAt;
      }
    }
  }

  return {
    messageCount: messages.length,
    humanMessageCount,
    systemMessageCount,
    meaningfulMessageCount,
    lastMessageAt,
    lastHumanMessageAt,
    lastMeaningfulMessageAt,
    lastSystemMessageAt,
  };
}

async function ensureSyncStillRunning(syncJobId) {
  const syncJob = await db.findLatestSlackArchiveSyncJob({ _id: syncJobId });
  if (!syncJob) {
    throw new SlackArchiveSyncCancelledError('Slack archive sync job no longer exists.');
  }

  if (syncJob.status === 'cancelled') {
    throw new SlackArchiveSyncCancelledError();
  }

  if (syncJob.status !== 'running') {
    throw new SlackArchiveSyncCancelledError(
      `Slack archive sync stopped unexpectedly with status "${syncJob.status}".`,
    );
  }
}

async function syncSingleConversation({
  userId,
  syncJobId,
  connection,
  conversation,
  requestedMessagesPerConversation,
  userCache,
}) {
  await db.updateSlackArchiveConversation(conversation._id?.toString?.() || conversation.id, {
    syncStatus: 'running',
    syncStartedAt: new Date(),
    syncError: '',
    syncCompletedAt: null,
  });

  const members = await listConversationMembers(conversation.channel, connection.userAccessToken, userCache);
  const rawMessages = await fetchConversationMessages(
    conversation.channel,
    connection.userAccessToken,
    requestedMessagesPerConversation,
  );
  const normalizedMessages = await normalizeSlackMessages({
    userId,
    teamId: connection.teamId,
    enterpriseId: connection.enterpriseId,
    slackConversationId: conversation.channel.id,
    rawMessages,
    token: connection.userAccessToken,
    userCache,
  });

  await db.bulkUpsertSlackArchiveMessages(normalizedMessages);
  const stats = buildConversationStats(normalizedMessages);
  const discoveredParticipantMap = new Map(members.map((member) => [member.slackUserId, member]));

  for (const message of normalizedMessages) {
    if (!message.slackUserId || discoveredParticipantMap.has(message.slackUserId)) {
      continue;
    }

    const profile = await getSlackUserProfile(connection.userAccessToken, message.slackUserId, userCache);
    if (profile) {
      discoveredParticipantMap.set(
        message.slackUserId,
        buildSlackParticipantFromProfile(profile, message.slackUserId),
      );
    }
  }

  const updatedConversation = await db.updateSlackArchiveConversation(
    conversation._id?.toString?.() || conversation.id,
    {
      participants: Array.from(discoveredParticipantMap.values()),
      syncStatus: 'complete',
      syncCursor: '',
      syncError: '',
      syncAttemptCount: Number(conversation.syncAttemptCount || 0) + 1,
      syncCompletedAt: new Date(),
      lastMessageSyncAt: new Date(),
      lastSyncedAt: new Date(),
      sourceUpdatedAt: stats.lastMessageAt,
      sourceLastMessageAt: stats.lastMessageAt,
      ...stats,
    },
  );

  await db.updateSlackArchiveSyncJob(syncJobId, {
    phase: 'syncing_messages',
  });

  return {
    conversation: updatedConversation,
    messageCount: normalizedMessages.length,
  };
}

async function performArchiveSync(syncJobId, user, options = {}) {
  const userId = getUserId(user);
  const connection = await getSlackConnectionForUser(user);
  const config = getSlackArchiveConfig();
  const requestedConversationLimit = clampPositiveInt(
    options.conversationLimit,
    config.syncConversationLimit,
    { max: 10000 },
  );
  const requestedMessagesPerConversation = clampPositiveInt(
    options.messagesPerConversation,
    config.syncMessagesPerConversation,
    { max: 5000 },
  );
  const leaseState = await acquireSyncLeases(userId);
  if (!leaseState) {
    throw new SlackArchiveServiceError(
      'Slack archive sync capacity is currently full. Please retry after an active sync completes.',
      429,
    );
  }

  const userCache = new Map();
  let processedConversationCount = 0;
  let skippedConversationCount = 0;
  let archivedConversationCount = 0;
  let archivedMessageCount = 0;
  const completedConversationIds = [];

  try {
    await db.updateSlackArchiveSyncJob(syncJobId, {
      status: 'running',
      phase: 'discovering_conversations',
      startedAt: new Date(),
      completedAt: null,
      errorMessage: '',
    });

    const channels = await listSlackConversations(connection.userAccessToken, requestedConversationLimit);
    const discoveredConversations = [];

    for (const channel of channels) {
      const discoveredAt = new Date();
      const upsertedConversation = await db.upsertSlackArchiveConversation({
        user: userId,
        slackConversationId: channel.id,
        teamId: connection.teamId,
        enterpriseId: connection.enterpriseId,
        conversationType: mapConversationType(channel),
        name: channel.name || '',
        topic: channel.topic?.value || '',
        purpose: channel.purpose?.value || '',
        isArchived: Boolean(channel.is_archived),
        isShared: Boolean(channel.is_shared),
        isExtShared: Boolean(channel.is_ext_shared),
        isOrgShared: Boolean(channel.is_org_shared),
        isSlackConnect: Boolean(channel.is_ext_shared || channel.is_org_shared),
        syncStatus: 'pending',
        sourceDiscoveredAt: discoveredAt,
      });

      discoveredConversations.push({
        channel,
        ...upsertedConversation,
      });
    }

    await db.updateSlackArchiveSyncJob(syncJobId, {
      discoveredConversationCount: discoveredConversations.length,
      phase: 'syncing_messages',
      checkpoint: {
        discoveredConversationIds: discoveredConversations.map((conversation) => conversation.slackConversationId),
      },
    });

    for (const conversation of discoveredConversations) {
      await ensureSyncStillRunning(syncJobId);
      await refreshSyncLeases(userId, leaseState);

      try {
        const result = await syncSingleConversation({
          userId,
          syncJobId,
          connection,
          conversation,
          requestedMessagesPerConversation,
          userCache,
        });

        processedConversationCount += 1;
        archivedConversationCount += 1;
        archivedMessageCount += result.messageCount;
        completedConversationIds.push(conversation.slackConversationId);
      } catch (error) {
        skippedConversationCount += 1;
        await db.updateSlackArchiveConversation(conversation._id?.toString?.() || conversation.id, {
          syncStatus: 'failed',
          syncError: error instanceof Error ? error.message : String(error),
          syncAttemptCount: Number(conversation.syncAttemptCount || 0) + 1,
          syncCompletedAt: new Date(),
        });

        logger.warn('[SlackArchiveService] Slack conversation sync failed', {
          slackConversationId: conversation.slackConversationId,
          message: error instanceof Error ? error.message : String(error),
        });
      }

      await db.updateSlackArchiveSyncJob(syncJobId, {
        processedConversationCount,
        skippedConversationCount,
        conversationCount: archivedConversationCount,
        messageCount: archivedMessageCount,
        stats: {
          archivedConversationCount,
          archivedMessageCount,
          skippedConversationCount,
        },
      });
    }

    const finalStatus = skippedConversationCount > 0 ? 'partial' : 'success';
    const completedSyncJob = await db.updateSlackArchiveSyncJob(syncJobId, {
      status: finalStatus,
      phase: 'complete',
      completedAt: new Date(),
      conversationCount: archivedConversationCount,
      messageCount: archivedMessageCount,
      processedConversationCount,
      skippedConversationCount,
      stats: {
        archivedConversationCount,
        archivedMessageCount,
        skippedConversationCount,
      },
    });
    const memoryProjection =
      completedConversationIds.length > 0
        ? queueSlackProjection({
            userId,
            tenantId: user?.tenantId,
            syncJobId: completedSyncJob?._id?.toString?.() || syncJobId,
            slackConversationIds: completedConversationIds,
            runStatus: finalStatus,
          }) || {
            status: 'skipped',
            reason: 'enterprise_memory_projection_unavailable',
          }
        : {
            status: 'skipped',
            reason: 'no_completed_conversations_in_run',
          };

    return {
      syncJob: formatSyncJob(completedSyncJob),
      mode: 'archive',
      conversationCount: archivedConversationCount,
      messageCount: archivedMessageCount,
      skippedConversationCount,
      processedConversationCount,
      memoryProjection,
      conversations: (
        await db.findSlackArchiveConversations(
          { user: userId },
          {
            limit: archivedConversationCount || requestedConversationLimit,
            sort: { lastMeaningfulMessageAt: -1, lastMessageAt: -1, updatedAt: -1 },
          },
        )
      ).map(formatConversation),
    };
  } catch (error) {
    const status = error instanceof SlackArchiveSyncCancelledError ? 'cancelled' : 'failure';
    const message = error instanceof Error ? error.message : String(error);
    await db.updateSlackArchiveSyncJob(syncJobId, {
      status,
      phase: status === 'cancelled' ? 'cancelled' : 'failed',
      completedAt: new Date(),
      processedConversationCount,
      skippedConversationCount,
      conversationCount: archivedConversationCount,
      messageCount: archivedMessageCount,
      errorMessage: message,
      stats: {
        archivedConversationCount,
        archivedMessageCount,
        skippedConversationCount,
      },
    });
    throw error;
  } finally {
    await releaseSyncLeases(userId, leaseState);
  }
}

async function getStatus(user) {
  assertEnabled();
  const userId = getUserId(user);
  const config = getSlackArchiveConfig();
  const SlackArchiveOAuthService = require('~/server/services/SlackArchiveOAuthService');

  const [
    conversationCount,
    messageCount,
    latestSync,
    latestProjection,
    projectionChunkCount,
    projectionConversationCount,
    projectionEntityCount,
    oauthConnection,
    activeSyncs,
  ] = await Promise.all([
    db.countSlackArchiveConversations({ user: userId }),
    db.countSlackArchiveMessages({ user: userId }),
    db.findLatestSlackArchiveSyncJob({ user: userId }),
    typeof db.findLatestEnterpriseMemoryJob === 'function'
      ? db.findLatestEnterpriseMemoryJob({ user: userId, source: 'slack', jobType: 'projection' })
      : null,
    userId && typeof db.countEnterpriseMemoryChunks === 'function'
      ? db.countEnterpriseMemoryChunks({
          user: userId,
          source: 'slack',
          sourceRecordType: 'slack_message',
        })
      : 0,
    userId && typeof db.countDistinctEnterpriseMemoryChunkField === 'function'
      ? db.countDistinctEnterpriseMemoryChunkField('sourceParentRecordId', {
          user: userId,
          source: 'slack',
          sourceRecordType: 'slack_message',
        })
      : 0,
    userId && typeof db.countEnterpriseMemoryEntities === 'function'
      ? db.countEnterpriseMemoryEntities({
          user: userId,
          source: 'slack',
          entityType: 'conversation',
          sourceRecordType: 'slack_conversation',
        })
      : 0,
    SlackArchiveOAuthService.getConnectionStatusForUser(user),
    config.maxConcurrentSyncs > 0
      ? db.countActiveSlackArchiveSyncLeases({
          leaseType: 'slot',
          leaseExpiresAt: { $gt: new Date() },
        })
      : Promise.resolve(0),
  ]);

  return {
    enabled: true,
    apiBaseUrl: config.apiBaseUrl,
    oauth: {
      installConfigured:
        Boolean(config.clientId) &&
        Boolean(config.clientSecret) &&
        Boolean(config.redirectUri) &&
        Boolean(
          process.env.SLACK_ARCHIVE_STATE_SECRET ||
            process.env.CREDS_KEY ||
            process.env.JWT_SECRET ||
            process.env.SLACK_ARCHIVE_CLIENT_SECRET,
        ),
      redirectUri: config.redirectUri,
      connected: oauthConnection.connected,
      identityLinked: oauthConnection.identityLinked,
      teamId: oauthConnection.teamId,
      enterpriseId: oauthConnection.enterpriseId,
    },
    userScopes: config.userScopes,
    botScopes: config.botScopes,
    maxConcurrentSyncs: config.maxConcurrentSyncs,
    activeSyncs,
    syncModes: ['archive'],
    conversationTypes: ['public_channel', 'private_channel', 'im', 'mpim'],
    threadSupport: true,
    conversationCount,
    messageCount,
    latestSync: formatSyncJob(latestSync),
    latestProjection: formatProjectionJob(latestProjection),
    projectionChunkCount,
    projectionConversationCount,
    projectionEntityCount,
  };
}

async function getSyncStartAvailability(user) {
  assertEnabled();
  const userId = getUserId(user);
  const runningSync = await db.findLatestSlackArchiveSyncJob({
    user: userId,
    status: 'running',
  });

  if (runningSync) {
    return {
      allowed: false,
      reason: 'already_running',
      status: 202,
      message: 'A Slack archive sync is already running.',
      syncJob: formatSyncJob(runningSync),
    };
  }

  return {
    allowed: true,
  };
}

async function syncUserArchive(user, options = {}) {
  assertEnabled();

  const availability = await getSyncStartAvailability(user);
  if (!availability.allowed) {
    throw new SlackArchiveServiceError(availability.message, availability.status || 409, availability);
  }

  const userId = getUserId(user);
  const config = getSlackArchiveConfig();
  const requestedConversationLimit = clampPositiveInt(
    options.conversationLimit,
    config.syncConversationLimit,
    { max: 10000 },
  );
  const requestedMessagesPerConversation = clampPositiveInt(
    options.messagesPerConversation,
    config.syncMessagesPerConversation,
    { max: 5000 },
  );

  const syncJob = await db.createSlackArchiveSyncJob({
    user: userId,
    status: 'running',
    mode: 'archive',
    phase: 'queued',
    requestedConversationLimit,
    requestedMessagesPerConversation,
    discoveredConversationCount: 0,
    processedConversationCount: 0,
    skippedConversationCount: 0,
    conversationCount: 0,
    messageCount: 0,
    startedAt: new Date(),
  });

  if (options.async === true) {
    void performArchiveSync(syncJob._id?.toString?.() || syncJob.id, user, options).catch((error) => {
      logger.error('[SlackArchiveService] Background Slack archive sync failed', {
        userId,
        message: error instanceof Error ? error.message : String(error),
      });
    });

    return {
      accepted: true,
      status: 'running',
      mode: 'archive',
      message: 'Slack archive sync started in the background.',
      syncJob: formatSyncJob(syncJob),
    };
  }

  return performArchiveSync(syncJob._id?.toString?.() || syncJob.id, user, options);
}

async function cancelRunningSync(user) {
  assertEnabled();
  const userId = getUserId(user);
  const runningSync = await db.findLatestSlackArchiveSyncJob({
    user: userId,
    status: 'running',
  });

  if (!runningSync) {
    return {
      cancelled: false,
      status: 'idle',
      syncJob: null,
      message: 'No Slack archive sync is currently running.',
    };
  }

  const updatedSync = await db.updateSlackArchiveSyncJob(runningSync._id?.toString?.() || runningSync.id, {
    status: 'cancelled',
    completedAt: new Date(),
  });

  return {
    cancelled: true,
    status: 'cancelled',
    syncJob: formatSyncJob(updatedSync),
    message: 'Slack archive sync marked as cancelled.',
  };
}

async function deleteUserArchive(user) {
  assertEnabled();
  const userId = getUserId(user);

  const latestRunningJob = await db.findLatestSlackArchiveSyncJob({
    user: userId,
    status: 'running',
  });

  if (latestRunningJob) {
    throw new SlackArchiveServiceError('Cannot delete Slack archive data while a sync is running', 409, {
      reason: 'sync_running',
      syncJobId: latestRunningJob._id?.toString?.() || latestRunningJob.id,
    });
  }

  const deleted = await runAsSystem(async () => {
    const [
      conversations,
      messages,
      syncJobs,
      syncLeases,
      projectionJobs,
      chunks,
      entities,
      relationships,
    ] = await Promise.all([
      db.deleteSlackArchiveConversations({ user: userId }),
      db.deleteSlackArchiveMessages({ user: userId }),
      db.deleteSlackArchiveSyncJobs({ user: userId }),
      db.deleteSlackArchiveSyncLeases({ user: userId }),
      typeof db.deleteEnterpriseMemoryJobs === 'function'
        ? db.deleteEnterpriseMemoryJobs({ user: userId, source: 'slack' })
        : 0,
      typeof db.deleteEnterpriseMemoryChunks === 'function'
        ? db.deleteEnterpriseMemoryChunks({ user: userId, source: 'slack' })
        : 0,
      typeof db.deleteEnterpriseMemoryEntities === 'function'
        ? db.deleteEnterpriseMemoryEntities({ user: userId, source: 'slack' })
        : 0,
      typeof db.deleteEnterpriseMemoryRelationships === 'function'
        ? db.deleteEnterpriseMemoryRelationships({ user: userId, source: 'slack' })
        : 0,
    ]);

    return {
      conversations,
      messages,
      syncJobs,
      syncLeases,
      projectionJobs,
      chunks,
      entities,
      relationships,
    };
  });

  return {
    deleted,
    message: 'Deleted archived Slack data for the current user.',
  };
}

async function listConversations(user, options = {}) {
  assertEnabled();
  const userId = getUserId(user);
  const limit = clampPositiveInt(options.limit, 25, { max: 100 });
  const offset = clampPositiveInt(options.offset, 0, { min: 0, max: 100000 });

  const conversations = await db.findSlackArchiveConversations(
    { user: userId },
    { limit, offset, sort: { lastMeaningfulMessageAt: -1, lastMessageAt: -1, updatedAt: -1 } },
  );

  return {
    limit,
    offset,
    count: conversations.length,
    conversations: conversations.map(formatConversation),
  };
}

async function listConversationMessages(user, conversationId, options = {}) {
  assertEnabled();
  const userId = getUserId(user);
  const normalizedConversationId = String(conversationId || '').trim();
  if (!normalizedConversationId) {
    throw new SlackArchiveServiceError('A Slack conversation id is required.', 400);
  }

  const limit = clampPositiveInt(options.limit, 50, { max: 200 });
  const offset = clampPositiveInt(options.offset, 0, { min: 0, max: 100000 });

  const messages = await db.findSlackArchiveMessages(
    { user: userId, slackConversationId: normalizedConversationId },
    { limit, offset, sort: { sentAt: -1, createdAt: -1 } },
  );

  return {
    conversationId: normalizedConversationId,
    limit,
    offset,
    count: messages.length,
    messages: messages.map(formatMessage),
  };
}

async function searchMessages(user, options = {}) {
  assertEnabled();
  const userId = getUserId(user);
  const query = String(options.query || options.q || '').trim();
  if (!query) {
    throw new SlackArchiveServiceError('A search query is required.', 400);
  }

  const limit = clampPositiveInt(options.limit, getSlackArchiveConfig().searchLimit, { max: 100 });
  const offset = clampPositiveInt(options.offset, 0, { min: 0, max: 100000 });
  const regex = buildSearchRegex(query);
  const slackConversationId = String(options.conversationId || options.channelId || '').trim();
  const slackUserId = String(options.senderUserId || '').trim();

  const filter = {
    user: userId,
    ...(slackConversationId ? { slackConversationId } : {}),
    ...(slackUserId ? { slackUserId } : {}),
    ...(regex
      ? {
          $or: [{ text: regex }, { normalizedText: regex }, { 'attachments.text': regex }],
        }
      : {}),
  };

  const messages = await db.findSlackArchiveMessages(filter, {
    limit,
    offset,
    sort: { sentAt: -1, createdAt: -1 },
  });

  return {
    query,
    limit,
    offset,
    count: messages.length,
    results: messages.map(formatMessage),
  };
}

module.exports = {
  SlackArchiveServiceError,
  getSlackArchiveConfig,
  getUserId,
  getStatus,
  getSyncStartAvailability,
  syncUserArchive,
  cancelRunningSync,
  deleteUserArchive,
  listConversations,
  listConversationMessages,
  searchMessages,
};
