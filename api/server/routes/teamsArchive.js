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
    const payload = req.body || {};
    const runAsync = payload.async === true || req.query.async === 'true';

    if (runAsync) {
      const availability = await TeamsArchiveService.getSyncStartAvailability(req.user);

      if (!availability.allowed) {
        if (availability.reason === 'already_running') {
          return res.status(202).json({
            accepted: true,
            status: 'running',
            mode: 'chats',
            alreadyRunning: true,
            syncJob: availability.syncJob,
            message: availability.message,
          });
        }

        return res.status(availability.status || 409).json({
          message: availability.message,
          details: availability.details,
        });
      }

      void TeamsArchiveService.syncUserArchive(req.user, payload).catch((error) => {
        if (error?.name === 'TeamsArchiveSyncCancelledError') {
          logger.info('[TeamsArchiveRoutes] Background Teams archive sync cancelled by user');
          return;
        }
        logger.error('[TeamsArchiveRoutes] Background Teams archive sync failed', error);
      });

      return res.status(202).json({
        accepted: true,
        status: 'running',
        mode: 'chats',
        message: 'Teams archive sync started in the background',
      });
    }

    const result = await TeamsArchiveService.syncUserArchive(req.user, payload);
    res.json(result);
  } catch (error) {
    handleTeamsArchiveError(res, error);
  }
});

router.post('/cancel', async (req, res) => {
  try {
    const result = await TeamsArchiveService.cancelRunningSync(req.user);
    res.json(result);
  } catch (error) {
    handleTeamsArchiveError(res, error);
  }
});

router.post('/reset', async (req, res) => {
  try {
    if (req.body?.confirm !== true) {
      return res.status(400).json({
        message: 'Teams archive reset requires confirm=true.',
      });
    }

    const result = await TeamsArchiveService.deleteUserArchive(req.user);
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
