const { getTenantId } = require('@librechat/data-schemas');
const { getAppConfig } = require('~/server/services/Config');

function getUserId(user) {
  return user?.id || user?._id?.toString?.() || '';
}

async function getArchiveFeatureFlags(user) {
  if (!user) {
    return {};
  }

  const appConfig = await getAppConfig({
    role: user?.role,
    userId: getUserId(user),
    tenantId: user?.tenantId || getTenantId(),
  });

  return appConfig?.interfaceConfig?.archiveFeatures || {};
}

async function isArchiveFeatureAllowed(user, featureName) {
  const flags = await getArchiveFeatureFlags(user);
  return flags?.[featureName] === true;
}

module.exports = {
  getArchiveFeatureFlags,
  isArchiveFeatureAllowed,
};
