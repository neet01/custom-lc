const express = require('express');
const { createAdminUsageHandlers } = require('@librechat/api');
const { SystemCapabilities } = require('@librechat/data-schemas');
const { requireCapability } = require('~/server/middleware/roles/capabilities');
const { requireJwtAuth } = require('~/server/middleware');
const db = require('~/models');

const router = express.Router();

const requireAdminAccess = requireCapability(SystemCapabilities.ACCESS_ADMIN);
const requireReadUsage = requireCapability(SystemCapabilities.READ_USAGE);

const handlers = createAdminUsageHandlers({
  findUsageRecords: db.findUsageRecords,
  countUsageRecords: db.countUsageRecords,
  summarizeUsageByUser: db.summarizeUsageByUser,
  summarizeUsageOverview: db.summarizeUsageOverview,
  findUsers: db.findUsers,
});

router.use(requireJwtAuth, requireAdminAccess);

router.get('/summary', requireReadUsage, handlers.getUsageSummary);
router.get('/', requireReadUsage, handlers.listUsage);

module.exports = router;
