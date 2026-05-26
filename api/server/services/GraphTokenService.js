const client = require('openid-client');
const { logger } = require('@librechat/data-schemas');
const { CacheKeys } = require('librechat-data-provider');
const getLogStores = require('~/cache/getLogStores');

const GRAPH_TOKEN_EXPIRY_SKEW_SECONDS = 300;
const ASSERTION_REFRESH_LEEWAY_SECONDS = 120;

function resolveOpenIdConfig() {
  /** Resolve lazily to avoid brittle module initialization order / partial export issues. */
  const directStrategyModule = require('~/strategies/openidStrategy');
  if (typeof directStrategyModule?.getOpenIdConfig === 'function') {
    return directStrategyModule.getOpenIdConfig();
  }

  const strategiesModule = require('~/strategies');
  if (typeof strategiesModule?.getOpenIdConfig === 'function') {
    return strategiesModule.getOpenIdConfig();
  }

  throw new Error('OpenID configuration accessor is unavailable');
}

function getTokenRefreshParams() {
  return process.env.OPENID_SCOPE ? { scope: process.env.OPENID_SCOPE } : {};
}

function getAssertionExpiryEpochSeconds(user) {
  const expiresAt = Number(user?.federatedTokens?.expires_at || 0);
  return Number.isFinite(expiresAt) && expiresAt > 0 ? expiresAt : 0;
}

function shouldRefreshAssertion(user, leewaySeconds = ASSERTION_REFRESH_LEEWAY_SECONDS) {
  const expiresAt = getAssertionExpiryEpochSeconds(user);
  if (!expiresAt) {
    return false;
  }

  const nowSeconds = Math.floor(Date.now() / 1000);
  return expiresAt - nowSeconds <= leewaySeconds;
}

function updateUserFederatedTokens(user, tokenset) {
  if (!user?.federatedTokens || !tokenset?.access_token) {
    return;
  }

  user.federatedTokens.access_token = tokenset.access_token;

  if (tokenset.refresh_token) {
    user.federatedTokens.refresh_token = tokenset.refresh_token;
  }

  if (tokenset.id_token) {
    user.federatedTokens.id_token = tokenset.id_token;
  }

  if (Number.isFinite(tokenset.expires_in) && tokenset.expires_in > 0) {
    user.federatedTokens.expires_at = Math.floor(Date.now() / 1000) + tokenset.expires_in;
  }
}

async function refreshOpenIdAccessToken(user) {
  const refreshToken = user?.federatedTokens?.refresh_token;
  if (!refreshToken) {
    throw new Error('No OpenID refresh token is available for delegated Graph access');
  }

  const config = resolveOpenIdConfig();
  const refreshParams = getTokenRefreshParams();

  logger.info('[GraphTokenService] Refreshing delegated OpenID access token for Graph access', {
    userId: user?.id || user?._id?.toString?.() || null,
    openidId: user?.openidId || null,
  });

  const tokenset = await client.refreshTokenGrant(config, refreshToken, refreshParams);
  updateUserFederatedTokens(user, tokenset);

  return user?.federatedTokens?.access_token || tokenset?.access_token;
}

function isAssertionExpiryError(error) {
  const code = String(error?.error || error?.code || '').toLowerCase();
  const description = String(error?.error_description || error?.message || '').toLowerCase();

  return (
    code === 'invalid_grant' &&
    (description.includes('assertion is not within its valid time range') ||
      description.includes('access token is not expired before using it for user assertion') ||
      description.includes('aadsts500133'))
  );
}

async function requestGraphToken(config, accessToken, scopes) {
  const grantResponse = await client.genericGrantRequest(
    config,
    'urn:ietf:params:oauth:grant-type:jwt-bearer',
    {
      scope: scopes,
      assertion: accessToken,
      requested_token_use: 'on_behalf_of',
    },
  );

  return {
    access_token: grantResponse.access_token,
    token_type: 'Bearer',
    expires_in: grantResponse.expires_in || 3600,
    scope: scopes,
  };
}

/**
 * Get Microsoft Graph API token using existing token exchange mechanism
 * @param {Object} user - User object with OpenID information
 * @param {string} accessToken - Federated access token used as OBO assertion
 * @param {string} scopes - Graph API scopes for the token
 * @param {boolean} fromCache - Whether to try getting token from cache first
 * @returns {Promise<Object>} Graph API token response with access_token and expires_in
 */
async function getGraphApiToken(user, accessToken, scopes, fromCache = true) {
  try {
    if (!user.openidId) {
      throw new Error('User must be authenticated via Entra ID to access Microsoft Graph');
    }

    if (!accessToken) {
      throw new Error('Access token is required for token exchange');
    }

    if (!scopes) {
      throw new Error('Graph API scopes are required for token exchange');
    }

    const config = resolveOpenIdConfig();
    if (!config) {
      throw new Error('OpenID configuration not available');
    }

    const cacheKey = `${user.openidId}:${scopes}`;
    const tokensCache = getLogStores(CacheKeys.OPENID_EXCHANGED_TOKENS);

    if (fromCache) {
      const cachedToken = await tokensCache.get(cacheKey);
      if (
        cachedToken?.access_token &&
        (!cachedToken.expires_at || cachedToken.expires_at > Date.now() + GRAPH_TOKEN_EXPIRY_SKEW_SECONDS * 1000)
      ) {
        logger.debug(`[GraphTokenService] Using cached Graph API token for user: ${user.openidId}`);
        return cachedToken;
      }
    }

    if (shouldRefreshAssertion(user)) {
      accessToken = await refreshOpenIdAccessToken(user);
    }

    logger.debug(`[GraphTokenService] Requesting new Graph API token for user: ${user.openidId}`);
    logger.debug(`[GraphTokenService] Requested scopes: ${scopes}`);

    let tokenResponse;
    try {
      tokenResponse = await requestGraphToken(config, accessToken, scopes);
    } catch (error) {
      if (!isAssertionExpiryError(error) || !user?.federatedTokens?.refresh_token) {
        throw error;
      }

      logger.warn('[GraphTokenService] Graph OBO assertion expired, refreshing delegated token and retrying', {
        userId: user?.id || user?._id?.toString?.() || null,
        openidId: user?.openidId || null,
      });
      accessToken = await refreshOpenIdAccessToken(user);
      tokenResponse = await requestGraphToken(config, accessToken, scopes);
    }

    await tokensCache.set(
      cacheKey,
      {
        ...tokenResponse,
        expires_at: Date.now() + tokenResponse.expires_in * 1000,
      },
      Math.max(60, tokenResponse.expires_in - GRAPH_TOKEN_EXPIRY_SKEW_SECONDS) * 1000,
    );

    logger.debug(
      `[GraphTokenService] Successfully obtained and cached Graph API token for user: ${user.openidId}`,
    );
    return tokenResponse;
  } catch (error) {
    logger.error(
      `[GraphTokenService] Failed to acquire Graph API token for user ${user.openidId}:`,
      error,
    );
    const wrappedError = new Error(`Graph token acquisition failed: ${error.message}`);
    wrappedError.code = error?.code || error?.error;
    wrappedError.cause = error;
    wrappedError.error_description = error?.error_description;
    throw wrappedError;
  }
}

module.exports = {
  getGraphApiToken,
};
