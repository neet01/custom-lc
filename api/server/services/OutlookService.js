const { isEnabled } = require('@librechat/api');
const { logger } = require('@librechat/data-schemas');
const { getGraphApiToken } = require('~/server/services/GraphTokenService');
const OutlookAIService = require('~/server/services/OutlookAIService');

const DEFAULT_GRAPH_BASE_URL = 'https://graph.microsoft.us/v1.0';
const DEFAULT_SCOPES = 'https://graph.microsoft.us/.default';
const USER_PROFILE_SELECT =
  'id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation';
const DEFAULT_MEETING_DURATION_MINUTES = 30;
const DEFAULT_WORKING_DAYS = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday'];
const DEFAULT_WORKDAY_START = '09:00:00';
const DEFAULT_WORKDAY_END = '17:00:00';
const DEFAULT_WORKING_TIME_ZONE = 'Eastern Standard Time';
const TENTATIVE_ATTENDEE_CONFIDENCE_PENALTY = 15;
const TENTATIVE_ORGANIZER_CONFIDENCE_PENALTY = 20;
const MAX_TENTATIVE_ATTENDEE_PENALTY = 45;
const MIN_MEETING_CONFIDENCE = 5;
const WINDOWS_TO_IANA_TIME_ZONE = {
  UTC: 'UTC',
  'Eastern Standard Time': 'America/New_York',
  'Central Standard Time': 'America/Chicago',
  'Mountain Standard Time': 'America/Denver',
  'Pacific Standard Time': 'America/Los_Angeles',
  'Alaskan Standard Time': 'America/Anchorage',
  'Hawaiian Standard Time': 'Pacific/Honolulu',
  'US Eastern Standard Time': 'America/Indianapolis',
  'US Mountain Standard Time': 'America/Phoenix',
};

class OutlookServiceError extends Error {
  constructor(message, status = 500, details) {
    super(message);
    this.name = 'OutlookServiceError';
    this.status = status;
    this.details = details;
  }
}

function normalizeOutlookUsage(usage) {
  if (!usage || typeof usage !== 'object') {
    return null;
  }

  const inputTokens = Number(usage.input_tokens ?? usage.inputTokens ?? 0) || 0;
  const outputTokens = Number(usage.output_tokens ?? usage.outputTokens ?? 0) || 0;
  const totalTokens =
    Number(usage.total_tokens ?? usage.totalTokens ?? inputTokens + outputTokens) ||
    inputTokens + outputTokens;

  if (inputTokens <= 0 && outputTokens <= 0 && totalTokens <= 0) {
    return null;
  }

  return {
    input_tokens: inputTokens,
    output_tokens: outputTokens,
    total_tokens: totalTokens,
    model: usage.model,
    provider: usage.provider,
  };
}

function buildOutlookUsageEntry({ context, usage }) {
  const normalizedUsage = normalizeOutlookUsage(usage);
  if (!normalizedUsage) {
    return null;
  }
  return {
    context,
    usage: normalizedUsage,
  };
}

function isOutlookEnabled() {
  return isEnabled(process.env.OUTLOOK_AI_ENABLED) || isEnabled(process.env.ENABLE_OUTLOOK_AI);
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

function getOutlookConfig() {
  return {
    enabled: isOutlookEnabled(),
    graphBaseUrl: normalizeGraphBaseUrl(
      process.env.OUTLOOK_GRAPH_BASE_URL || DEFAULT_GRAPH_BASE_URL,
    ),
    scopes: process.env.OUTLOOK_GRAPH_SCOPES || DEFAULT_SCOPES,
    includeUserContext: process.env.OUTLOOK_AI_INCLUDE_USER_CONTEXT !== 'false',
    includeDirectoryContext: isEnabled(process.env.OUTLOOK_AI_INCLUDE_DIRECTORY_CONTEXT),
    includeMailboxSettings: isEnabled(process.env.OUTLOOK_AI_INCLUDE_MAILBOX_SETTINGS),
    enableMeetingScheduling: process.env.OUTLOOK_AI_ENABLE_MEETING_SCHEDULING !== 'false',
  };
}

function assertEnabled() {
  if (!isOutlookEnabled()) {
    throw new OutlookServiceError('Outlook AI Inbox is not enabled', 403);
  }
}

function assertDelegatedUser(user) {
  if (!user?.openidId || user?.provider !== 'openid') {
    throw new OutlookServiceError('Outlook access requires Entra ID authentication', 403);
  }

  if (!isEnabled(process.env.OPENID_REUSE_TOKENS)) {
    throw new OutlookServiceError('Outlook access requires OPENID_REUSE_TOKENS=true', 403);
  }

  if (!user?.federatedTokens?.access_token) {
    throw new OutlookServiceError(
      'No delegated OpenID token is available for Microsoft Graph',
      401,
    );
  }
}

async function getDelegatedGraphToken(user, scopes = getOutlookConfig().scopes) {
  assertEnabled();
  assertDelegatedUser(user);
  const tokenResponse = await getGraphApiToken(user, user.federatedTokens.access_token, scopes);
  if (!tokenResponse?.access_token) {
    throw new OutlookServiceError(
      'Microsoft Graph token exchange did not return an access token',
      502,
    );
  }
  return tokenResponse.access_token;
}

function buildGraphUrl(pathname, query) {
  const { graphBaseUrl } = getOutlookConfig();
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
      logger.warn('[OutlookService] Microsoft Graph request failed', {
        status: response.status,
        path: pathname,
        graphMessage,
      });
    }
    throw new OutlookServiceError('Microsoft Graph request failed', response.status, graphMessage);
  }

  if (response.status === 204) {
    return null;
  }

  return response.json();
}

function normalizeEmailAddress(recipient) {
  const address = recipient?.emailAddress;
  return {
    name: address?.name || '',
    address: address?.address || '',
  };
}

function normalizeDirectoryUser(profile, fallback = {}) {
  return {
    id: profile?.id,
    displayName: profile?.displayName || fallback.name || '',
    email: profile?.mail || profile?.userPrincipalName || fallback.address || '',
    userPrincipalName: profile?.userPrincipalName,
    jobTitle: profile?.jobTitle || '',
    department: profile?.department || '',
    officeLocation: profile?.officeLocation || '',
  };
}

function normalizeMessage(message, includeBody = false) {
  const bodyContent = includeBody ? message.body?.content || '' : undefined;
  const bodyContentType = includeBody
    ? String(message.body?.contentType || 'text').toLowerCase()
    : undefined;
  const isHtmlBody = bodyContentType === 'html';
  return {
    id: message.id,
    conversationId: message.conversationId,
    subject: message.subject || '(No subject)',
    from: normalizeEmailAddress(message.from),
    toRecipients: Array.isArray(message.toRecipients)
      ? message.toRecipients.map(normalizeEmailAddress)
      : undefined,
    ccRecipients: Array.isArray(message.ccRecipients)
      ? message.ccRecipients.map(normalizeEmailAddress)
      : undefined,
    receivedDateTime: message.receivedDateTime,
    sentDateTime: message.sentDateTime,
    createdDateTime: message.createdDateTime,
    lastModifiedDateTime: message.lastModifiedDateTime,
    bodyPreview: message.bodyPreview || '',
    importance: message.importance || 'normal',
    inferenceClassification: message.inferenceClassification,
    isRead: Boolean(message.isRead),
    isDraft: Boolean(message.isDraft),
    hasAttachments: Boolean(message.hasAttachments),
    webLink: message.webLink,
    bodyContentType,
    bodyHtml: includeBody && isHtmlBody ? bodyContent : undefined,
    body: includeBody ? normalizeEmailBodyText(bodyContent, bodyContentType) : undefined,
  };
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

function normalizeEmailBodyText(content, contentType = 'text') {
  const raw = String(content || '');
  if (String(contentType || '').toLowerCase() !== 'html') {
    return raw;
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

function getMessageTimestamp(message) {
  return new Date(message.receivedDateTime || message.sentDateTime || 0).getTime();
}

function sortMessagesByDateAscending(messages) {
  return [...messages].sort((a, b) => getMessageTimestamp(a) - getMessageTimestamp(b));
}

function escapeODataString(value) {
  return String(value || '').replace(/'/g, "''");
}

function normalizeCalendarEvent(event) {
  return {
    id: event.id,
    subject: event.subject || '(No subject)',
    start: event.start,
    end: event.end,
    location: event.location?.displayName || '',
    organizer: normalizeEmailAddress(event.organizer),
    showAs: event.showAs,
    isOnlineMeeting: Boolean(event.isOnlineMeeting),
    webLink: event.webLink,
  };
}

function normalizeOnlineMeetingEvent(event) {
  return {
    id: event.id,
    subject: event.subject || '(No subject)',
    start: event.start,
    end: event.end,
    webLink: event.webLink,
    onlineMeeting: event.onlineMeeting
      ? {
          joinUrl: event.onlineMeeting.joinUrl,
          conferenceId: event.onlineMeeting.conferenceId,
        }
      : undefined,
  };
}

function normalizeMailboxSettings(settings) {
  if (!settings) {
    return undefined;
  }

  return {
    timeZone: settings.timeZone,
    workingHours: settings.workingHours
      ? {
          daysOfWeek: settings.workingHours.daysOfWeek,
          startTime: settings.workingHours.startTime,
          endTime: settings.workingHours.endTime,
          timeZone: settings.workingHours.timeZone?.name || settings.workingHours.timeZone,
        }
      : undefined,
  };
}

function getFolderPath(folder = 'inbox') {
  const normalized = String(folder || 'inbox').toLowerCase();
  const folderMap = {
    inbox: '/me/mailFolders/inbox/messages',
    drafts: '/me/mailFolders/drafts/messages',
    sent: '/me/mailFolders/sentitems/messages',
    sentitems: '/me/mailFolders/sentitems/messages',
    all: '/me/messages',
  };
  return folderMap[normalized] || folderMap.inbox;
}

function getMessageSelect(includeBody = false, includeRecipients = includeBody) {
  const fields = [
    'id',
    'conversationId',
    'subject',
    'from',
    'receivedDateTime',
    'sentDateTime',
    'createdDateTime',
    'lastModifiedDateTime',
    'isDraft',
    'bodyPreview',
    'importance',
    'inferenceClassification',
    'isRead',
    'hasAttachments',
    'webLink',
  ];
  if (includeRecipients) {
    fields.push('toRecipients', 'ccRecipients');
  }
  if (includeBody) {
    fields.push('body');
  }
  return fields.join(',');
}

function normalizeInboxView(inboxView) {
  const normalized = String(inboxView || 'focused').toLowerCase();
  return ['focused', 'other', 'all'].includes(normalized) ? normalized : 'focused';
}

function filterMessagesByInboxView(messages, folder, inboxView) {
  const normalizedFolder = String(folder || 'inbox').toLowerCase();
  const normalizedView = normalizeInboxView(inboxView);
  if (normalizedFolder !== 'inbox' || normalizedView === 'all') {
    return messages;
  }
  return messages.filter(
    (message) => String(message.inferenceClassification || '').toLowerCase() === normalizedView,
  );
}

async function listMessages(user, { folder = 'inbox', inboxView = 'focused', limit = 25 } = {}) {
  const top = Math.min(Math.max(Number(limit) || 25, 1), 100);
  const payload = await graphRequest(user, getFolderPath(folder), {
    query: {
      $top: top,
      $select: getMessageSelect(false),
      $orderby: 'receivedDateTime desc',
    },
  });
  const messages = Array.isArray(payload?.value)
    ? payload.value.map((message) => normalizeMessage(message, false))
    : [];

  return {
    messages: filterMessagesByInboxView(messages, folder, inboxView),
  };
}

async function getConversationMessages(user, conversationId, { limit = 25 } = {}) {
  if (!conversationId) {
    return [];
  }

  const top = Math.min(Math.max(Number(limit) || 25, 1), 50);
  const payload = await graphRequest(user, '/me/messages', {
    headers: {
      Prefer: 'outlook.body-content-type="html"',
    },
    query: {
      $top: top,
      $select: getMessageSelect(true),
      $filter: `conversationId eq '${escapeODataString(conversationId)}' and isDraft eq false`,
    },
  });

  const messages = Array.isArray(payload?.value)
    ? payload.value.map((message) => normalizeMessage(message, true))
    : [];

  return sortMessagesByDateAscending(messages);
}

function getDraftTimestamp(message) {
  return new Date(
    message.lastModifiedDateTime ||
      message.createdDateTime ||
      message.receivedDateTime ||
      message.sentDateTime ||
      0,
  ).getTime();
}

async function getConversationDraftReplies(user, conversationId, { limit = 10 } = {}) {
  if (!conversationId) {
    return [];
  }

  const top = Math.min(Math.max(Number(limit) || 10, 1), 25);
  const payload = await graphRequest(user, '/me/messages', {
    headers: {
      Prefer: 'outlook.body-content-type="html"',
    },
    query: {
      $top: top,
      $select: getMessageSelect(false, true),
      $orderby: 'lastModifiedDateTime desc',
      $filter: `conversationId eq '${escapeODataString(conversationId)}' and isDraft eq true`,
    },
  });

  const drafts = Array.isArray(payload?.value)
    ? payload.value.map((message) => normalizeMessage(message, false))
    : [];

  return drafts.sort((a, b) => getDraftTimestamp(b) - getDraftTimestamp(a));
}

async function deleteMessage(user, messageId) {
  if (!messageId) {
    throw new OutlookServiceError('Message id is required', 400);
  }

  await graphRequest(user, `/me/messages/${encodeURIComponent(messageId)}`, {
    method: 'DELETE',
  });

  return {
    messageId,
    message: 'Email moved to Deleted Items.',
  };
}

async function getMessage(user, messageId, { includeThread = true } = {}) {
  if (!messageId) {
    throw new OutlookServiceError('Message id is required', 400);
  }

  const payload = await graphRequest(user, `/me/messages/${encodeURIComponent(messageId)}`, {
    headers: {
      Prefer: 'outlook.body-content-type="html"',
    },
    query: {
      $select: getMessageSelect(true),
    },
  });
  const message = normalizeMessage(payload, true);

  if (!includeThread || !message.conversationId) {
    return message;
  }

  try {
    const [thread, draftReplies] = await Promise.all([
      getConversationMessages(user, message.conversationId),
      getConversationDraftReplies(user, message.conversationId),
    ]);
    return {
      ...message,
      thread,
      threadMessageCount: thread.length,
      draftReplies,
      draftReplyCount: draftReplies.length,
    };
  } catch (error) {
    logger.warn('[OutlookService] Conversation thread unavailable for selected message', {
      status: error?.status,
      message: error?.message,
      conversationId: message.conversationId,
    });
    return {
      ...message,
      thread: [message],
      threadMessageCount: 1,
      draftReplies: [],
      draftReplyCount: 0,
    };
  }
}

function shouldFetchCalendarContext(message) {
  if (!isEnabled(process.env.OUTLOOK_AI_INCLUDE_CALENDAR)) {
    return false;
  }
  const threadSource = Array.isArray(message.thread)
    ? message.thread
        .map(
          (threadMessage) =>
            `${threadMessage.subject || ''}\n${threadMessage.body || threadMessage.bodyPreview || ''}`,
        )
        .join('\n\n')
    : '';
  const source = `${message.subject || ''}\n${message.body || message.bodyPreview || ''}\n${threadSource}`;
  return /\b(meeting|calendar|invite|schedule|availability|available|appointment|call|zoom|teams)\b/i.test(
    source,
  );
}

async function getCalendarContext(user, message) {
  if (!shouldFetchCalendarContext(message)) {
    return [];
  }

  const now = new Date();
  const end = new Date(now.getTime() + 14 * 24 * 60 * 60 * 1000);
  try {
    const payload = await graphRequest(user, '/me/calendarView', {
      query: {
        startDateTime: now.toISOString(),
        endDateTime: end.toISOString(),
        $top: 10,
        $orderby: 'start/dateTime',
        $select: 'id,subject,start,end,location,organizer,showAs,isOnlineMeeting,webLink',
      },
    });
    return Array.isArray(payload?.value) ? payload.value.map(normalizeCalendarEvent) : [];
  } catch (error) {
    logger.warn('[OutlookService] Calendar context unavailable for Outlook AI', {
      status: error?.status,
      message: error?.message,
    });
    return [];
  }
}

function getParticipantKey(participant) {
  return String(participant?.address || '')
    .trim()
    .toLowerCase();
}

function collectThreadParticipants(message) {
  const participants = new Map();
  const threadMessages =
    Array.isArray(message.thread) && message.thread.length > 0 ? message.thread : [message];

  const addParticipant = (participant, role) => {
    const key = getParticipantKey(participant);
    if (!key) {
      return;
    }
    const existing = participants.get(key) || {
      name: participant.name || '',
      address: participant.address || '',
      roles: [],
    };
    if (!existing.roles.includes(role)) {
      existing.roles.push(role);
    }
    participants.set(key, existing);
  };

  for (const threadMessage of threadMessages) {
    addParticipant(threadMessage.from, 'from');
    for (const recipient of threadMessage.toRecipients || []) {
      addParticipant(recipient, 'to');
    }
    for (const recipient of threadMessage.ccRecipients || []) {
      addParticipant(recipient, 'cc');
    }
  }

  return Array.from(participants.values());
}

async function getCurrentUserProfile(user) {
  try {
    const profile = await graphRequest(user, '/me', {
      query: {
        $select: USER_PROFILE_SELECT,
      },
      suppressErrorLog: true,
    });
    return normalizeDirectoryUser(profile, {
      name: user?.name || user?.username,
      address: user?.email,
    });
  } catch (error) {
    logger.warn('[OutlookService] Signed-in user profile context unavailable', {
      status: error?.status,
      message: error?.message,
    });
    return normalizeDirectoryUser(null, {
      name: user?.name || user?.username,
      address: user?.email,
    });
  }
}

async function getCurrentUserManager(user) {
  try {
    const manager = await graphRequest(user, '/me/manager', {
      query: {
        $select: USER_PROFILE_SELECT,
      },
      suppressErrorLog: true,
    });
    return normalizeDirectoryUser(manager);
  } catch (_error) {
    return undefined;
  }
}

async function getMailboxContext(user, { force = false } = {}) {
  if (!force && !getOutlookConfig().includeMailboxSettings) {
    return undefined;
  }

  try {
    const settings = await graphRequest(user, '/me/mailboxSettings', {
      query: {
        $select: 'timeZone,workingHours',
      },
      suppressErrorLog: true,
    });
    return normalizeMailboxSettings(settings);
  } catch (error) {
    logger.warn('[OutlookService] Mailbox settings context unavailable', {
      status: error?.status,
      message: error?.message,
    });
    return undefined;
  }
}

async function getDirectoryProfileByAddress(user, participant) {
  const address = getParticipantKey(participant);
  if (!address) {
    return null;
  }

  try {
    const profile = await graphRequest(user, `/users/${encodeURIComponent(address)}`, {
      query: {
        $select: USER_PROFILE_SELECT,
      },
      suppressErrorLog: true,
    });
    return normalizeDirectoryUser(profile, participant);
  } catch (directLookupError) {
    if (directLookupError?.status !== 404) {
      return null;
    }
  }

  try {
    const escapedAddress = escapeODataString(address);
    const payload = await graphRequest(user, '/users', {
      query: {
        $top: 1,
        $select: USER_PROFILE_SELECT,
        $filter: `mail eq '${escapedAddress}' or userPrincipalName eq '${escapedAddress}'`,
      },
      suppressErrorLog: true,
    });
    const profile = Array.isArray(payload?.value) ? payload.value[0] : null;
    return profile ? normalizeDirectoryUser(profile, participant) : null;
  } catch (_error) {
    return null;
  }
}

function markParticipantRelationship(participant, signedInUser) {
  const participantEmail = getParticipantKey(participant);
  const signedInEmails = [signedInUser?.email, signedInUser?.userPrincipalName]
    .map((value) =>
      String(value || '')
        .trim()
        .toLowerCase(),
    )
    .filter(Boolean);

  return {
    ...participant,
    relationshipToSignedInUser: signedInEmails.includes(participantEmail)
      ? 'signed_in_user'
      : participant.internalUser
        ? 'internal_user'
        : 'external_or_unresolved',
  };
}

async function getParticipantDirectoryContext(user, message, signedInUser) {
  const participants = collectThreadParticipants(message);
  if (!getOutlookConfig().includeDirectoryContext) {
    return participants.map((participant) =>
      markParticipantRelationship(participant, signedInUser),
    );
  }

  const enrichedParticipants = [];
  for (const participant of participants.slice(0, 20)) {
    const profile = await getDirectoryProfileByAddress(user, participant);
    enrichedParticipants.push(
      markParticipantRelationship(
        {
          ...participant,
          internalUser: Boolean(profile),
          profile,
        },
        signedInUser,
      ),
    );
  }

  return enrichedParticipants;
}

async function getOutlookAIContext(user, message) {
  const config = getOutlookConfig();
  if (!config.includeUserContext) {
    return {
      participants: collectThreadParticipants(message),
    };
  }

  const signedInUser = await getCurrentUserProfile(user);
  const [manager, mailboxSettings] = await Promise.all([
    config.includeDirectoryContext ? getCurrentUserManager(user) : Promise.resolve(undefined),
    getMailboxContext(user),
  ]);
  const participants = await getParticipantDirectoryContext(user, message, signedInUser);

  return {
    signedInUser,
    manager,
    mailboxSettings,
    participants,
    rules: [
      'Draft only as signedInUser.',
      'Never sign as the sender, another recipient, or another person in the thread.',
      'If the thread is addressed to multiple people, decide whether signedInUser should respond before drafting.',
      'Use participant title and relationship context when available; do not invent hierarchy.',
    ],
  };
}

function normalizePositiveInteger(value, fallback, { min = 1, max = 120 } = {}) {
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) {
    return fallback;
  }
  return Math.min(Math.max(Math.round(parsed), min), max);
}

function buildMeetingDuration(durationMinutes) {
  const minutes = normalizePositiveInteger(
    durationMinutes,
    normalizePositiveInteger(
      process.env.OUTLOOK_AI_DEFAULT_MEETING_DURATION_MINUTES,
      DEFAULT_MEETING_DURATION_MINUTES,
      { min: 5, max: 240 },
    ),
    { min: 5, max: 240 },
  );
  return {
    minutes,
    isoDuration: `PT${minutes}M`,
  };
}

function normalizeMeetingAttendee(attendee) {
  const address = String(attendee?.address || attendee?.email || '').trim();
  if (!address || !address.includes('@')) {
    return null;
  }
  return {
    name: String(attendee?.name || attendee?.displayName || address).trim(),
    address,
  };
}

function resolveMeetingAttendees(message, outlookContext, requestedAttendees = []) {
  const attendees = new Map();
  const addAttendee = (attendee) => {
    const normalized = normalizeMeetingAttendee(attendee);
    if (!normalized) {
      return;
    }
    attendees.set(normalized.address.toLowerCase(), normalized);
  };

  if (Array.isArray(requestedAttendees) && requestedAttendees.length > 0) {
    requestedAttendees.forEach(addAttendee);
  } else {
    for (const participant of outlookContext?.participants || collectThreadParticipants(message)) {
      if (participant.relationshipToSignedInUser === 'signed_in_user') {
        continue;
      }
      addAttendee(participant);
    }
  }

  return Array.from(attendees.values()).slice(0, 20);
}

function stripReplyPrefix(subject) {
  return String(subject || 'Meeting')
    .replace(/^(\s*(re|fw|fwd)\s*:\s*)+/i, '')
    .trim();
}

function buildMeetingSubject(message, subject) {
  const normalized = String(subject || '').trim();
  if (normalized) {
    return normalized;
  }
  return `Meeting: ${stripReplyPrefix(message.subject) || 'Follow-up'}`;
}

function buildDateTimeRange(days = 14) {
  const now = new Date();
  const start = new Date(now.getTime() + 15 * 60 * 1000);
  const end = new Date(
    now.getTime() + normalizePositiveInteger(days, 14, { min: 1, max: 30 }) * 24 * 60 * 60 * 1000,
  );
  return {
    start: {
      dateTime: start.toISOString(),
      timeZone: 'UTC',
    },
    end: {
      dateTime: end.toISOString(),
      timeZone: 'UTC',
    },
  };
}

function normalizeGraphTime(value, fallback) {
  const match = String(value || '').match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if (!match) {
    return fallback;
  }
  const hours = Math.min(Math.max(Number(match[1]), 0), 23);
  const minutes = Math.min(Math.max(Number(match[2]), 0), 59);
  const seconds = Math.min(Math.max(Number(match[3] || 0), 0), 59);
  return [hours, minutes, seconds].map((part) => String(part).padStart(2, '0')).join(':');
}

function formatGraphDate(date) {
  return date.toISOString().slice(0, 10);
}

function getUtcDayName(date) {
  return ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'][
    date.getUTCDay()
  ];
}

function getWorkingHoursConfig(mailboxSettings) {
  const workingHours = mailboxSettings?.workingHours;
  return {
    daysOfWeek:
      Array.isArray(workingHours?.daysOfWeek) && workingHours.daysOfWeek.length > 0
        ? workingHours.daysOfWeek.map((day) => String(day).toLowerCase())
        : DEFAULT_WORKING_DAYS,
    startTime: normalizeGraphTime(
      process.env.OUTLOOK_AI_WORKDAY_START || workingHours?.startTime,
      DEFAULT_WORKDAY_START,
    ),
    endTime: normalizeGraphTime(
      process.env.OUTLOOK_AI_WORKDAY_END || workingHours?.endTime,
      DEFAULT_WORKDAY_END,
    ),
    timeZone:
      process.env.OUTLOOK_AI_WORKING_HOURS_TIME_ZONE ||
      workingHours?.timeZone ||
      mailboxSettings?.timeZone ||
      DEFAULT_WORKING_TIME_ZONE,
  };
}

function getIanaTimeZone(timeZone) {
  const normalized = String(timeZone || '').trim();
  if (!normalized) {
    return undefined;
  }
  return WINDOWS_TO_IANA_TIME_ZONE[normalized] || normalized;
}

function parseGraphDateTimeParts(value) {
  const match = String(value || '').match(
    /^(\d{4})-(\d{2})-(\d{2})T(\d{1,2}):(\d{2})(?::(\d{2}))?/,
  );
  if (!match) {
    return null;
  }
  return {
    year: Number(match[1]),
    month: Number(match[2]),
    day: Number(match[3]),
    hours: Number(match[4]),
    minutes: Number(match[5]),
    seconds: Number(match[6] || 0),
    date: `${match[1]}-${match[2]}-${match[3]}`,
  };
}

function getDayNameFromDateParts(parts) {
  const date = new Date(Date.UTC(parts.year, parts.month - 1, parts.day));
  return getUtcDayName(date);
}

function getTimeMinutes(parts) {
  return parts.hours * 60 + parts.minutes;
}

function getUtcInstantFromGraphDateTime(value) {
  const parts = parseGraphDateTimeParts(value);
  if (!parts) {
    return null;
  }
  return new Date(
    Date.UTC(parts.year, parts.month - 1, parts.day, parts.hours, parts.minutes, parts.seconds),
  );
}

function getDateTimePartsInTimeZone(date, timeZone) {
  const ianaTimeZone = getIanaTimeZone(timeZone);
  if (!ianaTimeZone || Number.isNaN(date.getTime())) {
    return null;
  }

  try {
    const values = Object.fromEntries(
      new Intl.DateTimeFormat('en-US', {
        timeZone: ianaTimeZone,
        hour12: false,
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
      })
        .formatToParts(date)
        .filter((part) => part.type !== 'literal')
        .map((part) => [part.type, part.value]),
    );

    return {
      year: Number(values.year),
      month: Number(values.month),
      day: Number(values.day),
      hours: Number(values.hour === '24' ? '0' : values.hour),
      minutes: Number(values.minute),
      seconds: Number(values.second),
      date: `${values.year}-${values.month}-${values.day}`,
    };
  } catch {
    return null;
  }
}

function getComparableDateTimeParts(value, targetTimeZone) {
  const sourceTimeZone = String(value?.timeZone || '').trim();
  if (!value?.dateTime) {
    return null;
  }
  if (
    sourceTimeZone &&
    targetTimeZone &&
    sourceTimeZone.toLowerCase() === String(targetTimeZone).toLowerCase()
  ) {
    return parseGraphDateTimeParts(value.dateTime);
  }
  if (sourceTimeZone.toUpperCase() === 'UTC') {
    const date = getUtcInstantFromGraphDateTime(value.dateTime);
    return date ? getDateTimePartsInTimeZone(date, targetTimeZone) : null;
  }
  return parseGraphDateTimeParts(value.dateTime);
}

function isMeetingSlotInsideWorkingHours(slot, workingHours) {
  const start = getComparableDateTimeParts(slot.start, workingHours.timeZone);
  const end = getComparableDateTimeParts(slot.end, workingHours.timeZone);
  if (!start || !end || start.date !== end.date) {
    return false;
  }

  const startMinutes = getTimeMinutes(start);
  const endMinutes = getTimeMinutes(end);
  const workStartMinutes = getTimeMinutes(
    parseGraphDateTimeParts(`2000-01-01T${workingHours.startTime}`),
  );
  const workEndMinutes = getTimeMinutes(
    parseGraphDateTimeParts(`2000-01-01T${workingHours.endTime}`),
  );
  const dayName = getDayNameFromDateParts(start);

  return (
    workingHours.daysOfWeek.includes(dayName) &&
    startMinutes >= workStartMinutes &&
    endMinutes <= workEndMinutes &&
    endMinutes > startMinutes
  );
}

function buildWorkingHourTimeSlots(days = 14, mailboxSettings) {
  const normalizedDays = normalizePositiveInteger(days, 14, { min: 1, max: 30 });
  const workingHours = getWorkingHoursConfig(mailboxSettings);
  const slots = [];
  const startDate = new Date();

  for (let offset = 1; slots.length < normalizedDays && offset <= normalizedDays + 14; offset++) {
    const date = new Date(startDate.getTime() + offset * 24 * 60 * 60 * 1000);
    if (!workingHours.daysOfWeek.includes(getUtcDayName(date))) {
      continue;
    }
    const graphDate = formatGraphDate(date);
    slots.push({
      start: {
        dateTime: `${graphDate}T${workingHours.startTime}`,
        timeZone: workingHours.timeZone,
      },
      end: {
        dateTime: `${graphDate}T${workingHours.endTime}`,
        timeZone: workingHours.timeZone,
      },
    });
  }

  return slots.length > 0 ? slots : [buildDateTimeRange(days)];
}

function normalizeMeetingTimeSlot(slot) {
  if (!slot?.start?.dateTime || !slot?.end?.dateTime) {
    throw new OutlookServiceError('A meeting time slot with start and end is required', 400);
  }
  return {
    start: {
      dateTime: slot.start.dateTime,
      timeZone: slot.start.timeZone || 'UTC',
    },
    end: {
      dateTime: slot.end.dateTime,
      timeZone: slot.end.timeZone || slot.start.timeZone || 'UTC',
    },
  };
}

function normalizeMeetingSuggestion(suggestion, index) {
  const slot = normalizeMeetingTimeSlot(suggestion.meetingTimeSlot);
  const { confidence, confidenceReason } =
    adjustSuggestionConfidenceForTentativeConflicts(suggestion);
  return {
    id: `slot-${index + 1}`,
    confidence,
    confidenceReason,
    organizerAvailability: suggestion.organizerAvailability,
    suggestionReason: suggestion.suggestionReason,
    attendeeAvailability: suggestion.attendeeAvailability,
    start: slot.start,
    end: slot.end,
  };
}

function normalizeAvailability(availability) {
  return String(availability || '')
    .trim()
    .toLowerCase();
}

function isTentativeAvailability(availability) {
  const normalized = normalizeAvailability(availability);
  return (
    normalized === 'tentative' ||
    normalized === 'tentativelybusy' ||
    normalized === 'tentatively_busy' ||
    normalized === 'tentatively-busy'
  );
}

function countTentativeAttendeeConflicts(attendeeAvailability) {
  if (!Array.isArray(attendeeAvailability)) {
    return 0;
  }
  return attendeeAvailability.reduce((count, attendee) => {
    return isTentativeAvailability(attendee?.availability) ? count + 1 : count;
  }, 0);
}

function adjustSuggestionConfidenceForTentativeConflicts(suggestion) {
  const baseConfidence = Number(suggestion?.confidence);
  const hasBaseConfidence = Number.isFinite(baseConfidence);
  const tentativeAttendeeCount = countTentativeAttendeeConflicts(suggestion?.attendeeAvailability);
  const organizerTentative = isTentativeAvailability(suggestion?.organizerAvailability);

  if (!hasBaseConfidence || (!organizerTentative && tentativeAttendeeCount === 0)) {
    return {
      confidence: hasBaseConfidence ? baseConfidence : undefined,
      confidenceReason: undefined,
    };
  }

  const attendeePenalty = Math.min(
    MAX_TENTATIVE_ATTENDEE_PENALTY,
    tentativeAttendeeCount * TENTATIVE_ATTENDEE_CONFIDENCE_PENALTY,
  );
  const organizerPenalty = organizerTentative ? TENTATIVE_ORGANIZER_CONFIDENCE_PENALTY : 0;
  const adjustedConfidence = Math.max(
    MIN_MEETING_CONFIDENCE,
    Math.round(baseConfidence - attendeePenalty - organizerPenalty),
  );

  const reasons = [];
  if (organizerTentative) {
    reasons.push('organizer is tentatively busy');
  }
  if (tentativeAttendeeCount > 0) {
    reasons.push(
      `${tentativeAttendeeCount} attendee${tentativeAttendeeCount === 1 ? '' : 's'} ${
        tentativeAttendeeCount === 1 ? 'has' : 'have'
      } tentative conflicts`,
    );
  }

  return {
    confidence: adjustedConfidence,
    confidenceReason: `Confidence reduced because ${reasons.join(' and ')}.`,
  };
}

function assertMeetingSchedulingEnabled() {
  if (!getOutlookConfig().enableMeetingScheduling) {
    throw new OutlookServiceError('Outlook meeting scheduling is not enabled', 403);
  }
}

async function proposeMeetingSlots(user, messageId, options = {}) {
  assertMeetingSchedulingEnabled();
  const message = await getMessage(user, messageId, { includeThread: true });
  const outlookContext = await getOutlookAIContext(user, message);
  const attendees = resolveMeetingAttendees(message, outlookContext, options.attendees);

  if (attendees.length === 0) {
    throw new OutlookServiceError('No meeting attendees could be resolved from this thread', 400);
  }

  const duration = buildMeetingDuration(options.durationMinutes);
  const mailboxSettings =
    outlookContext.mailboxSettings || (await getMailboxContext(user, { force: true }));
  const workingHours = getWorkingHoursConfig(mailboxSettings);
  const timeSlots = buildWorkingHourTimeSlots(options.days, mailboxSettings);
  const maxCandidates = normalizePositiveInteger(options.maxCandidates, 5, { min: 1, max: 10 });

  const payload = await graphRequest(user, '/me/findMeetingTimes', {
    method: 'POST',
    headers: {
      Prefer: `outlook.timezone="${workingHours.timeZone}"`,
    },
    body: {
      attendees: attendees.map((attendee) => ({
        type: 'required',
        emailAddress: {
          name: attendee.name,
          address: attendee.address,
        },
      })),
      timeConstraint: {
        activityDomain: 'work',
        timeslots: timeSlots,
      },
      meetingDuration: duration.isoDuration,
      maxCandidates: Math.max(maxCandidates, 10),
      returnSuggestionReasons: true,
    },
  });

  const suggestions = Array.isArray(payload?.meetingTimeSuggestions)
    ? payload.meetingTimeSuggestions
        .map(normalizeMeetingSuggestion)
        .filter((suggestion) => isMeetingSlotInsideWorkingHours(suggestion, workingHours))
        .slice(0, maxCandidates)
    : [];

  return {
    messageId,
    subject: buildMeetingSubject(message, options.subject),
    attendees,
    durationMinutes: duration.minutes,
    workingHours,
    emptySuggestionsReason: payload?.emptySuggestionsReason,
    suggestions,
  };
}

function escapeHtml(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function formatDateTimeForDraft(value) {
  if (!value?.dateTime) {
    return '';
  }
  const date = new Date(value.dateTime);
  if (Number.isNaN(date.getTime())) {
    return `${value.dateTime} ${value.timeZone || ''}`.trim();
  }
  return new Intl.DateTimeFormat('en-US', {
    dateStyle: 'medium',
    timeStyle: 'short',
    timeZone: value.timeZone === 'UTC' ? 'UTC' : undefined,
  }).format(date);
}

function formatPlainTextAsHtmlParagraphs(value) {
  const normalized = String(value || '').replace(/\r/g, '').trim();
  if (!normalized) {
    return '';
  }
  return normalized
    .split(/\n{2,}/)
    .map((paragraph) => `<p>${escapeHtml(paragraph).replace(/\n/g, '<br/>')}</p>`)
    .join('');
}

function buildLocalMeetingInviteNote({ message, subject, slot, instructions }) {
  const insights = buildLocalInsights(message);
  const lines = [];
  lines.push(`Objective: ${subject}.`);
  const scheduledStart = formatDateTimeForDraft(slot?.start);
  if (scheduledStart) {
    lines.push(`Scheduled time: ${scheduledStart}.`);
  }
  lines.push(`Context: ${truncateText(insights.summary, 700)}`);

  const suggestedActions = Array.isArray(insights.suggestedActions)
    ? insights.suggestedActions.slice(0, 4)
    : [];
  if (suggestedActions.length > 0) {
    lines.push('Agenda:');
    suggestedActions.forEach((action, index) => {
      lines.push(`${index + 1}. ${truncateText(action, 220)}`);
    });
  }

  const organizerNote = truncateText(String(instructions || '').trim(), 500);
  if (organizerNote) {
    lines.push(`Organizer note: ${organizerNote}`);
  }

  return lines.join('\n');
}

async function generateMeetingInviteNote({
  message,
  subject,
  slot,
  instructions,
  calendarEvents,
  outlookContext,
}) {
  if (
    OutlookAIService.isModelBackedAIEnabled() &&
    typeof OutlookAIService.generateMeetingInviteNote === 'function'
  ) {
    try {
      const generated = await OutlookAIService.generateMeetingInviteNote({
        message,
        subject,
        slot,
        instructions,
        calendarEvents,
        outlookContext,
      });

      const generatedNote =
        typeof generated === 'string'
          ? generated
          : typeof generated?.note === 'string'
            ? generated.note
            : '';

      if (generatedNote?.trim()) {
        return {
          note: truncateText(generatedNote, 1400),
          usage: normalizeOutlookUsage(generated?.usage),
        };
      }
    } catch (error) {
      OutlookAIService.logModelFailure('createTeamsMeeting.meetingNote', error);
    }
  }

  return {
    note: buildLocalMeetingInviteNote({
      message,
      subject,
      slot,
      instructions,
    }),
    usage: null,
  };
}

function buildMeetingBody({ message, meetingNote }) {
  const threadSubject = stripReplyPrefix(message.subject);
  return [
    '<p>Meeting scheduled from the Outlook thread.</p>',
    threadSubject ? `<p><strong>Source thread:</strong> ${escapeHtml(threadSubject)}</p>` : '',
    meetingNote
      ? `<p><strong>Meeting brief:</strong></p>${formatPlainTextAsHtmlParagraphs(meetingNote)}`
      : '',
  ]
    .filter(Boolean)
    .join('');
}

function buildMeetingDraftComment({ event, subject, slot, sentInvites = false }) {
  const joinUrl = event.onlineMeeting?.joinUrl || event.webLink;
  const start = formatDateTimeForDraft(slot.start);
  const action = sentInvites ? 'I scheduled' : 'I prepared a Teams meeting for';
  return [
    `${action} ${subject}${start ? ` for ${start}` : ''}.`,
    joinUrl ? `\n\nTeams link: ${joinUrl}` : '',
    sentInvites
      ? '\n\nSee the calendar invite for details.'
      : '\n\nPlease review the proposed time and Teams link.',
  ].join('');
}

async function createTeamsMeeting(user, messageId, options = {}) {
  assertMeetingSchedulingEnabled();
  const message = await getMessage(user, messageId, { includeThread: true });
  const outlookContext = await getOutlookAIContext(user, message);
  const attendees = resolveMeetingAttendees(message, outlookContext, options.attendees);
  if (attendees.length === 0) {
    throw new OutlookServiceError('No meeting attendees could be resolved from this thread', 400);
  }

  const slot = normalizeMeetingTimeSlot(options.slot);
  const subject = buildMeetingSubject(message, options.subject);
  const calendarEvents = await getCalendarContext(user, message);
  const meetingNoteResult = await generateMeetingInviteNote({
    message,
    subject,
    slot,
    instructions: options.instructions,
    calendarEvents,
    outlookContext,
  });
  const meetingNote = meetingNoteResult.note;
  const sendInvites = options.sendInvites === true;

  const eventPayload = {
    subject,
    body: {
      contentType: 'HTML',
      content: buildMeetingBody({ message, meetingNote }),
    },
    start: slot.start,
    end: slot.end,
    allowNewTimeProposals: true,
    isOnlineMeeting: true,
    onlineMeetingProvider: 'teamsForBusiness',
    attendees: attendees.map((attendee) => ({
      type: 'required',
      emailAddress: {
        name: attendee.name,
        address: attendee.address,
      },
    })),
  };

  const event = normalizeOnlineMeetingEvent(
    await graphRequest(user, '/me/events', {
      method: 'POST',
      body: eventPayload,
    }),
  );

  const shouldCreateReplyDraft = options.createReplyDraft === true;
  let draft;
  if (shouldCreateReplyDraft) {
    const comment = buildMeetingDraftComment({
      event,
      subject,
      slot,
      sentInvites: sendInvites,
    });
    draft = await graphRequest(user, `/me/messages/${encodeURIComponent(messageId)}/createReply`, {
      method: 'POST',
      body: { comment },
    });
    if (draft?.id) {
      await graphRequest(user, `/me/messages/${encodeURIComponent(draft.id)}`, {
        method: 'PATCH',
        body: {
          body: {
            contentType: 'Text',
            content: comment,
          },
        },
      }).catch((error) => {
        logger.warn('[OutlookService] Failed to patch meeting reply draft body', {
          status: error?.status,
          message: error?.message,
        });
      });
    }
  }

  const meetingUsageEntry = buildOutlookUsageEntry({
    context: 'outlook_meeting',
    usage: meetingNoteResult.usage,
  });

  return {
    sourceMessageId: messageId,
    conversationId: message.conversationId || `outlook:${messageId}`,
    event,
    attendees,
    meetingNotePreview: meetingNote,
    meetingDraft: {
      id: event.id,
      subject: event.subject,
      webLink: event.webLink,
    },
    draft: draft
      ? {
          id: draft.id,
          subject: draft.subject,
          webLink: draft.webLink,
        }
      : undefined,
    message:
      sendInvites
        ? 'Teams meeting invite sent to attendees.'
        : shouldCreateReplyDraft
          ? 'Teams meeting draft prepared. Review it in Outlook and optionally send your companion reply draft.'
          : 'Teams meeting draft prepared. Review it in Outlook and send when ready.',
    ...(meetingUsageEntry ? { _usage: [meetingUsageEntry] } : {}),
  };
}

function truncateText(value, maxLength = 1200) {
  const normalized = String(value || '')
    .replace(/\r/g, '')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
  if (normalized.length <= maxLength) {
    return normalized;
  }
  return `${normalized.slice(0, maxLength - 1).trim()}...`;
}

function splitSentences(value) {
  return String(value || '')
    .replace(/\s+/g, ' ')
    .split(/(?<=[.!?])\s+/)
    .map((sentence) => sentence.trim())
    .filter(Boolean);
}

function buildLocalInsights(message) {
  const threadMessages =
    Array.isArray(message.thread) && message.thread.length > 0 ? message.thread : [message];
  const source = threadMessages
    .map((threadMessage) => threadMessage.body || threadMessage.bodyPreview || '')
    .filter(Boolean)
    .join('\n\n');
  const sentences = splitSentences(source);
  const summary =
    sentences.length > 0
      ? truncateText(sentences.slice(0, 4).join(' '), 800)
      : 'No body text was available to summarize.';

  const lower = source.toLowerCase();
  const suggestedActions = [];
  if (source.includes('?')) {
    suggestedActions.push('Review and answer the explicit question(s) in the email.');
  }
  if (/\b(please|can you|could you|need|request|action required)\b/i.test(source)) {
    suggestedActions.push(
      'Identify the requested owner, deliverable, and due date before replying.',
    );
  }
  if (/\b(attach|attached|attachment)\b/i.test(source) || message.hasAttachments) {
    suggestedActions.push('Check related attachments before committing to next steps.');
  }
  if (suggestedActions.length === 0) {
    suggestedActions.push(
      'No obvious action request was detected; consider acknowledging receipt.',
    );
  }

  const riskSignals = [];
  if (/\b(urgent|asap|immediately|escalat|blocked|overdue)\b/i.test(lower)) {
    riskSignals.push('Time-sensitive language detected.');
  }
  if (message.importance === 'high') {
    riskSignals.push('Message is marked high importance.');
  }
  if (riskSignals.length === 0) {
    riskSignals.push('No obvious urgency or escalation signals detected.');
  }

  return {
    mode: 'local-extractive',
    summary,
    suggestedActions,
    riskSignals,
    generatedAt: new Date().toISOString(),
  };
}

async function analyzeMessage(user, messageId) {
  const message = await getMessage(user, messageId, { includeThread: true });
  const [calendarEvents, outlookContext] = await Promise.all([
    getCalendarContext(user, message),
    getOutlookAIContext(user, message),
  ]);
  if (OutlookAIService.isModelBackedAIEnabled()) {
    try {
      const generated = await OutlookAIService.generateAnalysis({
        message,
        calendarEvents,
        outlookContext,
      });
      if (generated) {
        const generatedInsights =
          generated?.insights && typeof generated.insights === 'object'
            ? generated.insights
            : generated;
        const usageEntry = buildOutlookUsageEntry({
          context: 'outlook_analyze',
          usage: generated?.usage,
        });
        return {
          messageId,
          conversationId: message.conversationId || `outlook:${messageId}`,
          insights: generatedInsights,
          ...(usageEntry ? { _usage: [usageEntry] } : {}),
        };
      }
    } catch (error) {
      OutlookAIService.logModelFailure('analyzeMessage', error);
    }
  }
  return {
    messageId,
    conversationId: message.conversationId || `outlook:${messageId}`,
    insights: buildLocalInsights(message),
  };
}

function buildDraftComment(message, { instructions = '', tone = 'professional' } = {}) {
  const trimmedInstructions = truncateText(instructions, 600);
  const opener =
    tone === 'concise'
      ? 'Thanks for the note.'
      : 'Thanks for reaching out. I reviewed your message and wanted to follow up.';

  const nextStep = trimmedInstructions
    ? `\n\nDrafting guidance from the user: ${trimmedInstructions}`
    : '\n\nI will review the details and follow up with any questions or next steps.';

  return `${opener}${nextStep}\n\nBest,`;
}

function normalizeReplyMode(options = {}) {
  const mode = String(options.replyMode || '').trim().toLowerCase();
  if (mode === 'reply' || mode === 'reply_all' || mode === 'smart') {
    return mode;
  }
  if (options.replyAll === true) {
    return 'reply_all';
  }
  if (options.replyAll === false) {
    return 'reply';
  }
  return 'smart';
}

function getSignedInEmailSet(outlookContext = {}) {
  const emails = [
    outlookContext?.signedInUser?.email,
    outlookContext?.signedInUser?.userPrincipalName,
  ]
    .map((value) => String(value || '').trim().toLowerCase())
    .filter(Boolean);
  return new Set(emails);
}

function shouldUseReplyAll(message, outlookContext = {}) {
  const signedInEmails = getSignedInEmailSet(outlookContext);
  const nonSignedToRecipients = (message.toRecipients || []).filter((recipient) => {
    const address = String(recipient?.address || '')
      .trim()
      .toLowerCase();
    return address && !signedInEmails.has(address);
  });
  const hasNonSignedCc = (message.ccRecipients || []).some((recipient) => {
    const address = String(recipient?.address || '')
      .trim()
      .toLowerCase();
    return address && !signedInEmails.has(address);
  });

  return nonSignedToRecipients.length > 1 || hasNonSignedCc;
}

function resolveReplyMode(message, outlookContext, options = {}) {
  const requestedMode = normalizeReplyMode(options);
  if (requestedMode === 'smart') {
    return shouldUseReplyAll(message, outlookContext) ? 'reply_all' : 'reply';
  }
  return requestedMode;
}

function normalizeDraftRecipients(payload) {
  return {
    toRecipients: Array.isArray(payload?.toRecipients)
      ? payload.toRecipients.map(normalizeEmailAddress)
      : [],
    ccRecipients: Array.isArray(payload?.ccRecipients)
      ? payload.ccRecipients.map(normalizeEmailAddress)
      : [],
  };
}

function resolveDraftRecipientName(recipient) {
  const name = String(recipient?.name || '').trim();
  if (name && !name.includes('@')) {
    return name;
  }
  const address = String(recipient?.address || '').trim();
  if (!address.includes('@')) {
    return address || 'there';
  }
  const local = address.split('@')[0] || 'there';
  return local
    .replace(/[._-]+/g, ' ')
    .trim()
    .replace(/\b\w/g, (match) => match.toUpperCase());
}

function buildExpectedSalutation(toRecipients = []) {
  const names = [];
  const seen = new Set();
  for (const recipient of toRecipients) {
    const resolved = resolveDraftRecipientName(recipient);
    const key = resolved.toLowerCase();
    if (!key || seen.has(key)) {
      continue;
    }
    seen.add(key);
    names.push(resolved);
  }

  if (names.length === 0) {
    return null;
  }
  if (names.length === 1) {
    return `Hi ${names[0]},`;
  }
  if (names.length === 2) {
    return `Hi ${names[0]} and ${names[1]},`;
  }
  if (names.length === 3) {
    return `Hi ${names[0]}, ${names[1]}, and ${names[2]},`;
  }
  return 'Hi all,';
}

function alignDraftSalutation(comment, toRecipients = []) {
  const expectedSalutation = buildExpectedSalutation(toRecipients);
  const text = String(comment || '').replace(/\r/g, '').trim();
  if (!expectedSalutation || !text) {
    return text;
  }

  const lines = text.split('\n');
  const greetingPattern =
    /^\s*(hi|hello|hey|good\s+(morning|afternoon|evening))\b[\s\S]*?(,|$)\s*$/i;
  const firstContentIndex = lines.findIndex((line) => line.trim().length > 0);

  if (firstContentIndex === -1) {
    return expectedSalutation;
  }

  if (greetingPattern.test(lines[firstContentIndex])) {
    lines[firstContentIndex] = expectedSalutation;
    return lines.join('\n').trim();
  }

  return `${expectedSalutation}\n\n${text}`;
}

async function createReplyDraft(user, messageId, options = {}) {
  const message = await getMessage(user, messageId, { includeThread: true });
  const [calendarEvents, outlookContext] = await Promise.all([
    getCalendarContext(user, message),
    getOutlookAIContext(user, message),
  ]);
  const replyMode = resolveReplyMode(message, outlookContext, options);
  const replyAction = replyMode === 'reply_all' ? 'createReplyAll' : 'createReply';
  const draftPayload = await graphRequest(
    user,
    `/me/messages/${encodeURIComponent(messageId)}/${replyAction}`,
    {
      method: 'POST',
      body: { comment: '' },
    },
  );
  let draftDetails = draftPayload;
  if (draftPayload?.id) {
    try {
      draftDetails = await graphRequest(user, `/me/messages/${encodeURIComponent(draftPayload.id)}`, {
        query: {
          $select: 'id,subject,webLink,toRecipients,ccRecipients',
        },
        suppressErrorLog: true,
      });
    } catch (error) {
      logger.warn('[OutlookService] Failed to fetch draft recipients after draft creation', {
        status: error?.status,
        message: error?.message,
      });
    }
  }
  const draftRecipients = normalizeDraftRecipients(draftDetails);
  let comment = buildDraftComment(message, options);
  let generatedUsage = null;

  if (OutlookAIService.isModelBackedAIEnabled()) {
    try {
      const generated = await OutlookAIService.generateReplyDraft({
        message,
        instructions: options.instructions,
        tone: options.tone,
        calendarEvents,
        outlookContext,
        draftRecipients,
        replyMode,
      });
      const generatedDraft =
        typeof generated === 'string'
          ? generated
          : typeof generated?.draft === 'string'
            ? generated.draft
            : '';
      generatedUsage = normalizeOutlookUsage(generated?.usage);
      if (generatedDraft?.trim()) {
        comment = generatedDraft.trim();
      }
    } catch (error) {
      OutlookAIService.logModelFailure('createReplyDraft', error);
    }
  }
  comment = alignDraftSalutation(comment, draftRecipients.toRecipients);

  if (draftPayload?.id && comment) {
    await graphRequest(user, `/me/messages/${encodeURIComponent(draftPayload.id)}`, {
      method: 'PATCH',
      body: {
        body: {
          contentType: 'Text',
          content: comment,
        },
      },
    }).catch((error) => {
      logger.warn('[OutlookService] Failed to patch generated draft body', {
        status: error?.status,
        message: error?.message,
      });
    });
  }

  const draftUsageEntry = buildOutlookUsageEntry({
    context: 'outlook_draft',
    usage: generatedUsage,
  });

  return {
    sourceMessageId: messageId,
    conversationId: message.conversationId || `outlook:${messageId}`,
    draftId: draftPayload?.id,
    subject: draftDetails?.subject || draftPayload?.subject,
    bodyPreview: comment,
    webLink: draftDetails?.webLink || draftPayload?.webLink,
    replyMode,
    toRecipients: draftRecipients.toRecipients,
    ccRecipients: draftRecipients.ccRecipients,
    message: 'Draft reply created. Review it in Outlook before sending.',
    ...(draftUsageEntry ? { _usage: [draftUsageEntry] } : {}),
  };
}

function getStatus(user) {
  const config = getOutlookConfig();
  return {
    ...config,
    connected:
      config.enabled &&
      user?.provider === 'openid' &&
      Boolean(user?.openidId) &&
      Boolean(user?.federatedTokens?.access_token) &&
      isEnabled(process.env.OPENID_REUSE_TOKENS),
    requires: {
      openid: true,
      openidReuseTokens: true,
      delegatedGraphScopes: config.scopes,
    },
    calendarContextEnabled: isEnabled(process.env.OUTLOOK_AI_INCLUDE_CALENDAR),
    userContextEnabled: config.includeUserContext,
    directoryContextEnabled: config.includeDirectoryContext,
    mailboxSettingsContextEnabled: config.includeMailboxSettings,
    meetingSchedulingEnabled: config.enableMeetingScheduling,
  };
}

module.exports = {
  OutlookServiceError,
  getOutlookConfig,
  getStatus,
  listMessages,
  getConversationMessages,
  getMessage,
  deleteMessage,
  analyzeMessage,
  createReplyDraft,
  proposeMeetingSlots,
  createTeamsMeeting,
  buildLocalInsights,
  getOutlookAIContext,
};
