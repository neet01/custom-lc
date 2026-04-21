const { isEnabled } = require('@librechat/api');
const { logger } = require('@librechat/data-schemas');
const { getGraphApiToken } = require('~/server/services/GraphTokenService');

const DEFAULT_GRAPH_BASE_URL = 'https://graph.microsoft.us/v1.0';
const DEFAULT_SCOPES = 'https://graph.microsoft.us/.default';

class OutlookServiceError extends Error {
  constructor(message, status = 500, details) {
    super(message);
    this.name = 'OutlookServiceError';
    this.status = status;
    this.details = details;
  }
}

function isOutlookEnabled() {
  return isEnabled(process.env.OUTLOOK_AI_ENABLED) || isEnabled(process.env.ENABLE_OUTLOOK_AI);
}

function normalizeGraphBaseUrl(baseUrl = DEFAULT_GRAPH_BASE_URL) {
  const trimmed = String(baseUrl || DEFAULT_GRAPH_BASE_URL).trim().replace(/\/+$/, '');
  if (/\/(v1\.0|beta)$/i.test(trimmed)) {
    return trimmed;
  }
  return `${trimmed}/v1.0`;
}

function getOutlookConfig() {
  return {
    enabled: isOutlookEnabled(),
    graphBaseUrl: normalizeGraphBaseUrl(process.env.OUTLOOK_GRAPH_BASE_URL || DEFAULT_GRAPH_BASE_URL),
    scopes: process.env.OUTLOOK_GRAPH_SCOPES || DEFAULT_SCOPES,
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
    throw new OutlookServiceError('No delegated OpenID token is available for Microsoft Graph', 401);
  }
}

async function getDelegatedGraphToken(user, scopes = getOutlookConfig().scopes) {
  assertEnabled();
  assertDelegatedUser(user);
  const tokenResponse = await getGraphApiToken(user, user.federatedTokens.access_token, scopes);
  if (!tokenResponse?.access_token) {
    throw new OutlookServiceError('Microsoft Graph token exchange did not return an access token', 502);
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
    logger.warn('[OutlookService] Microsoft Graph request failed', {
      status: response.status,
      path: pathname,
      graphMessage,
    });
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

function normalizeMessage(message, includeBody = false) {
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
    bodyPreview: message.bodyPreview || '',
    body: includeBody ? message.body?.content || '' : undefined,
    importance: message.importance || 'normal',
    isRead: Boolean(message.isRead),
    hasAttachments: Boolean(message.hasAttachments),
    webLink: message.webLink,
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

function getMessageSelect(includeBody = false) {
  const fields = [
    'id',
    'conversationId',
    'subject',
    'from',
    'receivedDateTime',
    'sentDateTime',
    'bodyPreview',
    'importance',
    'isRead',
    'hasAttachments',
    'webLink',
  ];
  if (includeBody) {
    fields.push('body', 'toRecipients', 'ccRecipients');
  }
  return fields.join(',');
}

async function listMessages(user, { folder = 'inbox', limit = 25 } = {}) {
  const top = Math.min(Math.max(Number(limit) || 25, 1), 50);
  const payload = await graphRequest(user, getFolderPath(folder), {
    query: {
      $top: top,
      $select: getMessageSelect(false),
      $orderby: 'receivedDateTime desc',
    },
  });
  return {
    messages: Array.isArray(payload?.value)
      ? payload.value.map((message) => normalizeMessage(message, false))
      : [],
  };
}

async function getMessage(user, messageId) {
  if (!messageId) {
    throw new OutlookServiceError('Message id is required', 400);
  }

  const payload = await graphRequest(user, `/me/messages/${encodeURIComponent(messageId)}`, {
    headers: {
      Prefer: 'outlook.body-content-type="text"',
    },
    query: {
      $select: getMessageSelect(true),
    },
  });
  return normalizeMessage(payload, true);
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
  const source = message.body || message.bodyPreview || '';
  const sentences = splitSentences(source);
  const summary =
    sentences.length > 0
      ? truncateText(sentences.slice(0, 3).join(' '), 700)
      : 'No body text was available to summarize.';

  const lower = source.toLowerCase();
  const suggestedActions = [];
  if (source.includes('?')) {
    suggestedActions.push('Review and answer the explicit question(s) in the email.');
  }
  if (/\b(please|can you|could you|need|request|action required)\b/i.test(source)) {
    suggestedActions.push('Identify the requested owner, deliverable, and due date before replying.');
  }
  if (/\b(attach|attached|attachment)\b/i.test(source) || message.hasAttachments) {
    suggestedActions.push('Check related attachments before committing to next steps.');
  }
  if (suggestedActions.length === 0) {
    suggestedActions.push('No obvious action request was detected; consider acknowledging receipt.');
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
  const message = await getMessage(user, messageId);
  return {
    messageId,
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

async function createReplyDraft(user, messageId, options = {}) {
  const message = await getMessage(user, messageId);
  const comment = buildDraftComment(message, options);
  const payload = await graphRequest(user, `/me/messages/${encodeURIComponent(messageId)}/createReply`, {
    method: 'POST',
    body: { comment },
  });

  return {
    sourceMessageId: messageId,
    draftId: payload?.id,
    subject: payload?.subject,
    bodyPreview: payload?.bodyPreview,
    webLink: payload?.webLink,
    message: 'Draft reply created. Review it in Outlook before sending.',
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
  };
}

module.exports = {
  OutlookServiceError,
  getOutlookConfig,
  getStatus,
  listMessages,
  getMessage,
  analyzeMessage,
  createReplyDraft,
  buildLocalInsights,
};
