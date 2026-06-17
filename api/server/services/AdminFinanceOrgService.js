const { logger } = require('@librechat/data-schemas');
const { getGraphApiToken } = require('~/server/services/GraphTokenService');

const DEFAULT_GRAPH_BASE_URL = 'https://graph.microsoft.us/v1.0';
const DEFAULT_GRAPH_SCOPES = 'https://graph.microsoft.us/User.Read.All';
const DEFAULT_CONCURRENCY = 5;

function isExplicitlyEnabled(value) {
  return ['true', '1', 'on', 'yes'].includes(String(value || '').trim().toLowerCase());
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

function getUserId(user) {
  return user?._id?.toString?.() || user?.id || '';
}

function getGraphConfig() {
  return {
    baseUrl: normalizeGraphBaseUrl(
      process.env.ADMIN_USAGE_GRAPH_BASE_URL ||
        process.env.OUTLOOK_GRAPH_BASE_URL ||
        DEFAULT_GRAPH_BASE_URL,
    ),
    scopes:
      process.env.ADMIN_USAGE_GRAPH_SCOPES ||
      process.env.OUTLOOK_GRAPH_SCOPES ||
      DEFAULT_GRAPH_SCOPES,
    concurrency:
      Number(process.env.ADMIN_USAGE_GRAPH_CONCURRENCY) > 0
        ? Math.min(Number(process.env.ADMIN_USAGE_GRAPH_CONCURRENCY), 10)
        : DEFAULT_CONCURRENCY,
  };
}

function getLookupCandidates(user) {
  return [
    user?.idOnTheSource,
    user?.email,
    user?.openidId,
    user?.username && String(user.username).includes('@') ? user.username : '',
  ]
    .map((value) => String(value || '').trim())
    .filter(Boolean)
    .filter((value, index, values) => values.indexOf(value) === index);
}

function mapGraphUser(payload) {
  return {
    graphUserId: payload?.id || '',
    team: payload?.department || '',
    role: payload?.jobTitle || '',
    company: payload?.companyName || '',
    officeLocation: payload?.officeLocation || '',
  };
}

async function fetchGraphUser({ accessToken, baseUrl, user }) {
  const select = 'id,displayName,mail,userPrincipalName,department,jobTitle,companyName,officeLocation';

  for (const candidate of getLookupCandidates(user)) {
    const url = new URL(`${baseUrl}/users/${encodeURIComponent(candidate)}`);
    url.searchParams.set('$select', select);

    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: 'application/json',
      },
    });

    if (response.ok) {
      return await response.json();
    }

    if (response.status === 404) {
      continue;
    }

    const body = await response.text().catch(() => '');
    const error = new Error(`Graph user lookup failed with status ${response.status}`);
    error.status = response.status;
    error.body = body.slice(0, 500);
    throw error;
  }

  return null;
}

async function mapWithConcurrency(items, concurrency, worker) {
  const results = new Array(items.length);
  let nextIndex = 0;

  async function runWorker() {
    while (nextIndex < items.length) {
      const currentIndex = nextIndex;
      nextIndex += 1;
      results[currentIndex] = await worker(items[currentIndex], currentIndex);
    }
  }

  await Promise.all(
    Array.from({ length: Math.min(concurrency, items.length) }, () => runWorker()),
  );

  return results;
}

async function resolveFinanceUserOrgMetadata(users, requester) {
  const result = new Map();

  if (!isExplicitlyEnabled(process.env.ADMIN_USAGE_GRAPH_ORG_ENRICHMENT_ENABLED)) {
    return result;
  }

  if (!Array.isArray(users) || users.length === 0) {
    return result;
  }

  if (!requester?.openidId || !requester?.federatedTokens?.access_token) {
    logger.debug('[AdminFinanceOrgService] Skipping Graph org enrichment; requester has no delegated token');
    return result;
  }

  const config = getGraphConfig();
  let graphToken;
  try {
    graphToken = await getGraphApiToken(
      requester,
      requester.federatedTokens.access_token,
      config.scopes,
    );
  } catch (error) {
    logger.warn('[AdminFinanceOrgService] Skipping Graph org enrichment; token acquisition failed', {
      error: error?.message,
      code: error?.code,
    });
    return result;
  }

  try {
    await mapWithConcurrency(users, config.concurrency, async (user) => {
      const userId = getUserId(user);
      if (!userId) {
        return;
      }

      try {
        const graphUser = await fetchGraphUser({
          accessToken: graphToken.access_token,
          baseUrl: config.baseUrl,
          user,
        });
        if (graphUser) {
          result.set(userId, mapGraphUser(graphUser));
        }
      } catch (error) {
        if (error?.status === 401 || error?.status === 403) {
          throw error;
        }
        logger.debug('[AdminFinanceOrgService] Graph user lookup skipped for finance export user', {
          userId,
          error: error?.message,
        });
      }
    });
  } catch (error) {
    logger.warn('[AdminFinanceOrgService] Graph org enrichment stopped; permission or Graph error', {
      status: error?.status,
      error: error?.message,
    });
  }

  return result;
}

module.exports = {
  resolveFinanceUserOrgMetadata,
};
