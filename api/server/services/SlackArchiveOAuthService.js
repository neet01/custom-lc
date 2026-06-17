const { randomBytes, createHmac, timingSafeEqual } = require('crypto');
const { logger } = require('@librechat/data-schemas');
const db = require('~/models');
const {
  SlackArchiveServiceError,
  getSlackArchiveConfig,
  getUserId,
} = require('~/server/services/SlackArchiveService');
const {
  updateUserPluginAuth,
  getUserPluginAuthValue,
} = require('~/server/services/PluginService');

const SLACK_ARCHIVE_PLUGIN_KEY = 'slack_archive';
const STATE_TTL_MS = 10 * 60 * 1000;

const AUTH_FIELDS = {
  userAccessToken: 'SLACK_ARCHIVE_USER_ACCESS_TOKEN',
  userScope: 'SLACK_ARCHIVE_USER_SCOPE',
};

function base64UrlEncode(value) {
  return Buffer.from(value).toString('base64url');
}

function base64UrlDecode(value) {
  return Buffer.from(value, 'base64url').toString('utf8');
}

function getStateSecret() {
  return (
    process.env.SLACK_ARCHIVE_STATE_SECRET ||
    process.env.CREDS_KEY ||
    process.env.JWT_SECRET ||
    process.env.SLACK_ARCHIVE_CLIENT_SECRET ||
    ''
  );
}

function assertOAuthConfigured() {
  const config = getSlackArchiveConfig();
  const missing = [];

  if (!config.clientId) {
    missing.push('SLACK_ARCHIVE_CLIENT_ID');
  }
  if (!config.clientSecret) {
    missing.push('SLACK_ARCHIVE_CLIENT_SECRET');
  }
  if (!config.redirectUri) {
    missing.push('SLACK_ARCHIVE_REDIRECT_URI or DOMAIN_SERVER');
  }
  if (!getStateSecret()) {
    missing.push('SLACK_ARCHIVE_STATE_SECRET or CREDS_KEY or JWT_SECRET');
  }

  if (missing.length > 0) {
    throw new SlackArchiveServiceError('Slack archive OAuth is not configured.', 501, {
      missing,
    });
  }

  return config;
}

function signStatePayload(encodedPayload) {
  return createHmac('sha256', getStateSecret()).update(encodedPayload).digest('base64url');
}

function createSignedState(payload) {
  const encodedPayload = base64UrlEncode(JSON.stringify(payload));
  const signature = signStatePayload(encodedPayload);
  return `${encodedPayload}.${signature}`;
}

function parseAndVerifyState(state) {
  const normalizedState = String(state || '').trim();
  if (!normalizedState || !normalizedState.includes('.')) {
    throw new SlackArchiveServiceError('Slack OAuth state is missing or malformed.', 400);
  }

  const [encodedPayload, signature] = normalizedState.split('.', 2);
  const expectedSignature = signStatePayload(encodedPayload);

  const providedBuffer = Buffer.from(signature);
  const expectedBuffer = Buffer.from(expectedSignature);
  if (
    providedBuffer.length !== expectedBuffer.length ||
    !timingSafeEqual(providedBuffer, expectedBuffer)
  ) {
    throw new SlackArchiveServiceError('Slack OAuth state signature is invalid.', 400);
  }

  let payload;
  try {
    payload = JSON.parse(base64UrlDecode(encodedPayload));
  } catch (error) {
    throw new SlackArchiveServiceError('Slack OAuth state payload could not be parsed.', 400, {
      cause: error instanceof Error ? error.message : String(error),
    });
  }

  if (!payload?.userId || !payload?.issuedAt || !payload?.nonce) {
    throw new SlackArchiveServiceError('Slack OAuth state payload is incomplete.', 400);
  }

  const issuedAt = Number(payload.issuedAt);
  if (!Number.isFinite(issuedAt) || Date.now() - issuedAt > STATE_TTL_MS) {
    throw new SlackArchiveServiceError('Slack OAuth state has expired.', 400);
  }

  return payload;
}

function getCallbackOriginFallback() {
  const domainServer = String(process.env.DOMAIN_SERVER || '').trim().replace(/\/+$/, '');
  return domainServer || null;
}

function buildRedirectUri() {
  const explicit = String(process.env.SLACK_ARCHIVE_REDIRECT_URI || '').trim();
  if (explicit) {
    return explicit;
  }

  const domainServer = getCallbackOriginFallback();
  if (!domainServer) {
    return '';
  }

  return `${domainServer}/api/slack-archive/oauth/callback`;
}

function normalizeReturnTo(value) {
  const rawValue = String(value || '').trim();
  if (!rawValue) {
    return undefined;
  }

  if (rawValue.startsWith('/')) {
    return rawValue;
  }

  const originFallback = getCallbackOriginFallback();
  if (!originFallback) {
    return undefined;
  }

  try {
    const parsed = new URL(rawValue);
    const expectedOrigin = new URL(originFallback).origin;
    if (parsed.origin !== expectedOrigin) {
      return undefined;
    }
    return parsed.toString();
  } catch (error) {
    return undefined;
  }
}

function getInstallStartStatus() {
  const config = getSlackArchiveConfig();
  const redirectUri = buildRedirectUri();
  return {
    clientIdConfigured: Boolean(config.clientId),
    clientSecretConfigured: Boolean(config.clientSecret),
    signingSecretConfigured: Boolean(process.env.SLACK_ARCHIVE_SIGNING_SECRET),
    redirectUri,
    stateSecretConfigured: Boolean(getStateSecret()),
  };
}

async function getConnectionStatusForUser(user) {
  const userId = getUserId(user);
  if (!userId) {
    return {
      connected: false,
      identityLinked: false,
      teamId: null,
      enterpriseId: null,
      scopesGranted: false,
    };
  }

  const [userAccessToken, userScope, identityLink, workspaceInstall] = await Promise.all([
    getUserPluginAuthValue(userId, AUTH_FIELDS.userAccessToken, false, SLACK_ARCHIVE_PLUGIN_KEY),
    getUserPluginAuthValue(userId, AUTH_FIELDS.userScope, false, SLACK_ARCHIVE_PLUGIN_KEY),
    typeof db.findSlackIdentityLink === 'function'
      ? db.findSlackIdentityLink({ user: userId, status: 'linked' })
      : Promise.resolve(null),
    typeof db.findSlackWorkspaceInstall === 'function'
      ? db.findSlackWorkspaceInstall({ status: 'active' })
      : Promise.resolve(null),
  ]);

  return {
    connected: Boolean(workspaceInstall?.botAccessToken && userAccessToken),
    identityLinked: Boolean(identityLink),
    teamId: identityLink?.teamId || workspaceInstall?.teamId || null,
    enterpriseId: identityLink?.enterpriseId || workspaceInstall?.enterpriseId || null,
    scopesGranted: Boolean(workspaceInstall?.botScopes || userScope),
  };
}

function buildInstallUrl(user, options = {}) {
  const config = assertOAuthConfigured();
  const userId = getUserId(user);
  if (!userId) {
    throw new SlackArchiveServiceError('Authenticated user context is required for Slack OAuth.', 401);
  }

  const state = createSignedState({
    userId,
    issuedAt: Date.now(),
    nonce: randomBytes(12).toString('hex'),
    returnTo: normalizeReturnTo(options.returnTo),
    team: typeof options.team === 'string' ? options.team : undefined,
  });

  const url = new URL('/oauth/v2/authorize', 'https://slack-gov.com');
  url.searchParams.set('client_id', config.clientId);
  url.searchParams.set('scope', config.botScopes);
  url.searchParams.set('user_scope', config.userScopes);
  url.searchParams.set('redirect_uri', config.redirectUri);
  url.searchParams.set('state', state);

  if (options.team) {
    url.searchParams.set('team', String(options.team));
  }

  return {
    installUrl: url.toString(),
    redirectUri: config.redirectUri,
    state,
  };
}

async function exchangeCodeForTokens({ code }) {
  const config = assertOAuthConfigured();
  const body = new URLSearchParams({
    code,
    client_id: config.clientId,
    client_secret: config.clientSecret,
    redirect_uri: config.redirectUri,
  });

  const response = await fetch(`${config.apiBaseUrl}/oauth.v2.access`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    body,
  });

  let payload;
  try {
    payload = await response.json();
  } catch (error) {
    throw new SlackArchiveServiceError('Slack OAuth token exchange returned a non-JSON response.', 502, {
      cause: error instanceof Error ? error.message : String(error),
    });
  }

  if (!response.ok || payload?.ok !== true) {
    throw new SlackArchiveServiceError('Slack OAuth token exchange failed.', 502, {
      status: response.status,
      error: payload?.error || payload?.message || 'unknown_error',
    });
  }

  return payload;
}

async function fetchSlackUserProfile({ userAccessToken, slackUserId, apiBaseUrl }) {
  if (!userAccessToken || !slackUserId) {
    return null;
  }

  const url = new URL('/users.info', apiBaseUrl);
  url.pathname = `${new URL(apiBaseUrl).pathname.replace(/\/+$/, '')}/users.info`;
  url.searchParams.set('user', slackUserId);

  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${userAccessToken}`,
    },
  });

  if (!response.ok) {
    logger.warn('[SlackArchiveOAuth] Failed to fetch Slack user profile', {
      slackUserId,
      status: response.status,
    });
    return null;
  }

  try {
    const payload = await response.json();
    if (payload?.ok !== true) {
      logger.warn('[SlackArchiveOAuth] Slack user profile response was not ok', {
        slackUserId,
        error: payload?.error,
      });
      return null;
    }
    return payload?.user || null;
  } catch (error) {
    logger.warn('[SlackArchiveOAuth] Slack user profile response was not JSON', {
      slackUserId,
      cause: error instanceof Error ? error.message : String(error),
    });
    return null;
  }
}

async function persistOAuthPayload(userId, payload) {
  const userAccessToken = payload?.authed_user?.access_token || '';
  const userScope = payload?.authed_user?.scope || '';
  const slackUserId = payload?.authed_user?.id || '';
  const workspaceInstallRecord = {
    installedByUser: userId,
    teamId: payload?.team?.id || '',
    teamName: payload?.team?.name || '',
    enterpriseId: payload?.enterprise?.id || '',
    enterpriseName: payload?.enterprise?.name || '',
    botUserId: payload.bot_user_id || '',
    botAccessToken: payload.access_token || '',
    botScopes: payload.scope || '',
    userScopes: userScope,
    installPayload: payload,
    status: 'active',
    installedAt: new Date(),
    lastValidatedAt: new Date(),
  };

  if (typeof db.upsertSlackWorkspaceInstall === 'function') {
    await db.upsertSlackWorkspaceInstall(workspaceInstallRecord);
  }

  const slackUserProfile = await fetchSlackUserProfile({
    userAccessToken,
    slackUserId,
    apiBaseUrl: getSlackArchiveConfig().apiBaseUrl,
  });

  if (slackUserId && typeof db.upsertSlackIdentityLink === 'function') {
    await db.upsertSlackIdentityLink({
      user: userId,
      slackUserId,
      teamId: payload?.team?.id || '',
      teamName: payload?.team?.name || '',
      enterpriseId: payload?.enterprise?.id || '',
      enterpriseName: payload?.enterprise?.name || '',
      slackEmail: slackUserProfile?.profile?.email || '',
      slackDisplayName:
        slackUserProfile?.profile?.display_name ||
        slackUserProfile?.real_name ||
        slackUserProfile?.name ||
        '',
      status: 'linked',
      source: 'oauth_install',
      linkedAt: new Date(),
      lastVerifiedAt: new Date(),
    });
  }

  const entries = [
    [AUTH_FIELDS.userAccessToken, userAccessToken],
    [AUTH_FIELDS.userScope, userScope],
  ].filter(([, value]) => value);

  await Promise.all(
    entries.map(([authField, value]) =>
      updateUserPluginAuth(userId, authField, SLACK_ARCHIVE_PLUGIN_KEY, value),
    ),
  );
}

async function handleOAuthCallback({ code, state, error, errorDescription }) {
  if (error) {
    throw new SlackArchiveServiceError('Slack OAuth authorization was denied or failed.', 400, {
      error,
      errorDescription,
    });
  }

  const statePayload = parseAndVerifyState(state);
  const userId = statePayload.userId;
  const tokenPayload = await exchangeCodeForTokens({ code });
  await persistOAuthPayload(userId, tokenPayload);

  logger.info('[SlackArchiveOAuth] Stored GovSlack OAuth tokens for user', {
    userId,
    teamId: tokenPayload?.team?.id,
    enterpriseId: tokenPayload?.enterprise?.id,
  });

  return {
    connected: true,
    userId,
    team: tokenPayload.team || null,
    enterprise: tokenPayload.enterprise || null,
    scopes: {
      bot: tokenPayload.scope || '',
      user: tokenPayload?.authed_user?.scope || '',
    },
    returnTo: statePayload.returnTo || null,
  };
}

module.exports = {
  AUTH_FIELDS,
  SLACK_ARCHIVE_PLUGIN_KEY,
  getInstallStartStatus,
  getConnectionStatusForUser,
  buildInstallUrl,
  handleOAuthCallback,
};
