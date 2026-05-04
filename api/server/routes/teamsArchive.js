const express = require('express');
const { logger } = require('@librechat/data-schemas');
const { requireJwtAuth } = require('~/server/middleware');
const TeamsArchiveService = require('~/server/services/TeamsArchiveService');

const router = express.Router();

function handleTeamsArchiveError(res, error) {
  if (error?.name === 'TeamsArchiveServiceError') {
    return res.status(error.status || 500).json({
      message: error.message,
      details: error.details,
    });
  }

  logger.error('[TeamsArchiveRoutes] Unexpected Teams archive route error', error);
  return res.status(500).json({ message: 'Teams archive request failed' });
}

router.use(requireJwtAuth);

router.get('/status', async (req, res) => {
  try {
    const result = await TeamsArchiveService.getStatus(req.user);
    res.json(result);
  } catch (error) {
    handleTeamsArchiveError(res, error);
  }
});

router.post('/sync', async (req, res) => {
  try {
    const result = await TeamsArchiveService.syncUserArchive(req.user, req.body || {});
    res.json(result);
  } catch (error) {
    handleTeamsArchiveError(res, error);
  }
});

router.get('/conversations', async (req, res) => {
  try {
    const result = await TeamsArchiveService.listConversations(req.user, {
      limit: req.query.limit,
      offset: req.query.offset,
    });
    res.json(result);
  } catch (error) {
    handleTeamsArchiveError(res, error);
  }
});

router.get('/conversations/:chatId/messages', async (req, res) => {
  try {
    const result = await TeamsArchiveService.listConversationMessages(req.user, req.params.chatId, {
      limit: req.query.limit,
      offset: req.query.offset,
    });
    res.json(result);
  } catch (error) {
    handleTeamsArchiveError(res, error);
  }
});

router.get('/search', async (req, res) => {
  try {
    const result = await TeamsArchiveService.searchMessages(req.user, {
      query: req.query.q,
      chatId: req.query.chatId,
      limit: req.query.limit,
      offset: req.query.offset,
    });
    res.json(result);
  } catch (error) {
    handleTeamsArchiveError(res, error);
  }
});

module.exports = router;
