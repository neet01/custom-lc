const express = require('express');
const { createAdminIssuesHandlers } = require('@librechat/api');
const { SystemCapabilities } = require('@librechat/data-schemas');
const { requireCapability } = require('~/server/middleware/roles/capabilities');
const { requireJwtAuth } = require('~/server/middleware');
const db = require('~/models');

const router = express.Router();

const requireAdminAccess = requireCapability(SystemCapabilities.ACCESS_ADMIN);
const requireReadUsage = requireCapability(SystemCapabilities.READ_USAGE);

const handlers = createAdminIssuesHandlers({
  findIssueReports: db.findIssueReports,
  countIssueReports: db.countIssueReports,
  findUsers: db.findUsers,
});

router.use(requireJwtAuth, requireAdminAccess);
router.get('/', requireReadUsage, handlers.listIssues);

module.exports = router;
