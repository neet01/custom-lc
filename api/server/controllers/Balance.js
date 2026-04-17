const { getBalanceConfig } = require('@librechat/api');
const { findBalanceByUser, upsertBalanceFields } = require('~/models');
const { getAppConfig } = require('~/server/services/Config');

async function getOrInitializeBalance(req) {
  const userId = req.user.id;
  let balanceData = await findBalanceByUser(userId);

  if (balanceData) {
    return balanceData;
  }

  const appConfig = await getAppConfig({
    role: req.user?.role,
    tenantId: req.user?.tenantId,
  });
  const balanceConfig = getBalanceConfig(appConfig);

  if (!balanceConfig?.enabled || balanceConfig.startBalance == null) {
    return null;
  }

  const fields = {
    user: userId,
    tokenCredits: balanceConfig.startBalance,
  };

  if (
    balanceConfig.autoRefillEnabled &&
    balanceConfig.refillAmount != null &&
    balanceConfig.refillIntervalUnit != null &&
    balanceConfig.refillIntervalValue != null
  ) {
    fields.autoRefillEnabled = true;
    fields.refillAmount = balanceConfig.refillAmount;
    fields.refillIntervalUnit = balanceConfig.refillIntervalUnit;
    fields.refillIntervalValue = balanceConfig.refillIntervalValue;
    fields.lastRefill = new Date();
  }

  return upsertBalanceFields(userId, fields);
}

async function balanceController(req, res) {
  const balanceData = await getOrInitializeBalance(req);

  if (!balanceData) {
    return res.status(404).json({ error: 'Balance not found' });
  }

  const { _id: _, __v, user, ...result } = balanceData;

  if (!result.autoRefillEnabled) {
    delete result.refillIntervalValue;
    delete result.refillIntervalUnit;
    delete result.lastRefill;
    delete result.refillAmount;
  }

  res.status(200).json(result);
}

module.exports = balanceController;
