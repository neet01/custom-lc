const express = require('express');
const { logger } = require('@librechat/data-schemas');
const { requireJwtAuth } = require('~/server/middleware');
const SlackArchiveService = require('~/server/services/SlackArchiveService');
const SlackArchiveOAuthService = require('~/server/services/SlackArchiveOAuthService');

const router = express.Router();

function handleSlackArchiveError(res, error) {
  if (error?.name === 'SlackArchiveServiceError') {
    return res.status(error.status || 500).json({
      message: error.message,
      details: error.details,
    });
  }

  logger.error('[SlackArchiveRoutes] Unexpected Slack archive route error', error);
  return res.status(500).json({ message: 'Slack archive request failed' });
}

router.get('/oauth/callback', async (req, res) => {
  try {
    const result = await SlackArchiveOAuthService.handleOAuthCallback({
      code: req.query.code,
      state: req.query.state,
      error: req.query.error,
      errorDescription: req.query.error_description,
    });

    if (result.returnTo) {
      const callbackOrigin =
        String(process.env.DOMAIN_SERVER || '').trim().replace(/\/+$/, '') ||
        `${req.protocol}://${req.get('host')}`;
      const url = result.returnTo.startsWith('/')
        ? new URL(result.returnTo, callbackOrigin)
        : new URL(result.returnTo);
      url.searchParams.set('slackArchive', 'connected');
      if (result.team?.id) {
        url.searchParams.set('teamId', result.team.id);
      }
      return res.redirect(url.toString());
    }

    res.json(result);
  } catch (error) {
    handleSlackArchiveError(res, error);
  }
});

router.use(requireJwtAuth);

router.get('/status', async (req, res) => {
  try {
    const result = await SlackArchiveService.getStatus(req.user);
    res.json(result);
  } catch (error) {
    handleSlackArchiveError(res, error);
  }
});

router.get('/oauth/start', async (req, res) => {
  try {
    const result = SlackArchiveOAuthService.buildInstallUrl(req.user, {
      team: req.query.team,
      returnTo: req.query.returnTo,
    });

    if (req.query.redirect === 'true') {
      return res.redirect(result.installUrl);
    }

    res.json(result);
  } catch (error) {
    handleSlackArchiveError(res, error);
  }
});

router.post('/sync', async (req, res) => {
  try {
    const payload = req.body || {};
    const result = await SlackArchiveService.syncUserArchive(req.user, payload);
    res.json(result);
  } catch (error) {
    handleSlackArchiveError(res, error);
  }
});

router.post('/cancel', async (req, res) => {
  try {
    const result = await SlackArchiveService.cancelRunningSync(req.user);
    res.json(result);
  } catch (error) {
    handleSlackArchiveError(res, error);
  }
});

router.post('/reset', async (req, res) => {
  try {
    if (req.body?.confirm !== true) {
      return res.status(400).json({
        message: 'Slack archive reset requires confirm=true.',
      });
    }

    const result = await SlackArchiveService.deleteUserArchive(req.user);
    res.json(result);
  } catch (error) {
    handleSlackArchiveError(res, error);
  }
});

router.get('/conversations', async (req, res) => {
  try {
    const result = await SlackArchiveService.listConversations(req.user, {
      limit: req.query.limit,
      offset: req.query.offset,
    });
    res.json(result);
  } catch (error) {
    handleSlackArchiveError(res, error);
  }
});

router.get('/conversations/:conversationId/messages', async (req, res) => {
  try {
    const result = await SlackArchiveService.listConversationMessages(
      req.user,
      req.params.conversationId,
      {
        limit: req.query.limit,
        offset: req.query.offset,
      },
    );
    res.json(result);
  } catch (error) {
    handleSlackArchiveError(res, error);
  }
});

router.get('/search', async (req, res) => {
  try {
    const result = await SlackArchiveService.searchMessages(req.user, {
      query: req.query.q,
      conversationId: req.query.conversationId,
      senderUserId: req.query.senderUserId,
      limit: req.query.limit,
      offset: req.query.offset,
    });
    res.json(result);
  } catch (error) {
    handleSlackArchiveError(res, error);
  }
});

module.exports = router;
