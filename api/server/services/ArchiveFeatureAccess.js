const { getTenantId, logger } = require('@librechat/data-schemas');

function getUserId(user) {
  return user?.id || user?._id?.toString?.() || '';
}

function resolveGetAppConfig() {
  try {
    const configService = require('~/server/services/Config');
    if (typeof configService?.getAppConfig === 'function') {
      return configService.getAppConfig;
    }
  } catch (error) {
    logger.warn('[ArchiveFeatureAccess] Failed to load aggregated Config service', error);
  }

  try {
    const appConfigService = require('~/server/services/Config/app');
    if (typeof appConfigService?.getAppConfig === 'function') {
      return appConfigService.getAppConfig;
    }
  } catch (error) {
    logger.warn('[ArchiveFeatureAccess] Failed to load app config service', error);
  }

  return null;
}

async function getArchiveFeatureFlags(user) {
  if (!user) {
    return {};
  }

  const getAppConfig = resolveGetAppConfig();
  if (typeof getAppConfig !== 'function') {
    logger.warn(
      '[ArchiveFeatureAccess] getAppConfig unavailable during archive access check; allowing archive feature access by fallback.',
    );
    return null;
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
  if (flags == null) {
    return true;
  }
  return flags?.[featureName] === true;
}

module.exports = {
  getArchiveFeatureFlags,
  isArchiveFeatureAllowed,
};
