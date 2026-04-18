const express = require('express');
const { createIssueHandlers } = require('@librechat/api');
const { requireJwtAuth } = require('~/server/middleware');
const db = require('~/models');

const router = express.Router();

const handlers = createIssueHandlers({
  createIssueReport: db.createIssueReport,
});

router.use(requireJwtAuth);
router.post('/', handlers.reportIssue);

module.exports = router;
