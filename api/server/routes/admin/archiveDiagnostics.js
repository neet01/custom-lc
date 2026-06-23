const express = require('express');
const { SystemCapabilities } = require('@librechat/data-schemas');
const { requireCapability } = require('~/server/middleware/roles/capabilities');
const { requireJwtAuth } = require('~/server/middleware');
const { getArchiveDiagnostics } = require('~/server/services/ArchiveDiagnosticsService');

const router = express.Router();

const requireAdminAccess = requireCapability(SystemCapabilities.ACCESS_ADMIN);
const requireReadUsage = requireCapability(SystemCapabilities.READ_USAGE);

router.use(requireJwtAuth, requireAdminAccess);

router.get('/', requireReadUsage, async (req, res, next) => {
  try {
    const result = await getArchiveDiagnostics({
      source: req.query.source,
      userId: req.query.userId,
      q: req.query.q,
      type: req.query.type,
      status: req.query.status,
      limit: req.query.limit,
      offset: req.query.offset,
    });
    res.json(result);
  } catch (error) {
    next(error);
  }
});

module.exports = router;
