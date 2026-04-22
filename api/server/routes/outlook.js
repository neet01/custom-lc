const express = require('express');
const { logger } = require('@librechat/data-schemas');
const { requireJwtAuth } = require('~/server/middleware');
const OutlookService = require('~/server/services/OutlookService');
const db = require('~/models');

const router = express.Router();

async function recordAudit(req, record) {
  try {
    await db.createOutlookAudit({
      user: req.user.id,
      ...record,
    });
  } catch (error) {
    logger.warn('[OutlookRoutes] Failed to write Outlook audit record', {
      action: record?.action,
      status: record?.status,
      error: error?.message,
    });
  }
}

function getErrorAudit(error) {
  return {
    status: 'failure',
    errorCode: error?.status ? String(error.status) : error?.name || 'OUTLOOK_ERROR',
    errorMessage: error?.message || 'Outlook operation failed',
  };
}

function handleOutlookError(res, error) {
  if (error?.name === 'OutlookServiceError') {
    return res.status(error.status || 500).json({
      message: error.message,
      details: error.details,
    });
  }

  logger.error('[OutlookRoutes] Unexpected Outlook route error', error);
  return res.status(500).json({ message: 'Outlook request failed' });
}

router.use(requireJwtAuth);

router.get('/status', (req, res) => {
  res.json(OutlookService.getStatus(req.user));
});

router.get('/messages', async (req, res) => {
  try {
    const result = await OutlookService.listMessages(req.user, {
      folder: req.query.folder,
      inboxView: req.query.inboxView,
      limit: req.query.limit,
    });
    await recordAudit(req, {
      action: 'mailbox_listed',
      status: 'success',
      metadata: {
        folder: typeof req.query.folder === 'string' ? req.query.folder : 'inbox',
        inboxView: typeof req.query.inboxView === 'string' ? req.query.inboxView : 'focused',
        limit: result.messages.length,
      },
    });
    res.json(result);
  } catch (error) {
    await recordAudit(req, {
      action: 'mailbox_listed',
      ...getErrorAudit(error),
      metadata: {
        folder: typeof req.query.folder === 'string' ? req.query.folder : 'inbox',
        inboxView: typeof req.query.inboxView === 'string' ? req.query.inboxView : 'focused',
      },
    });
    handleOutlookError(res, error);
  }
});

router.get('/messages/:messageId', async (req, res) => {
  try {
    const result = await OutlookService.getMessage(req.user, req.params.messageId);
    await recordAudit(req, {
      action: 'message_viewed',
      status: 'success',
      graphMessageId: result.id,
      graphConversationId: result.conversationId,
      metadata: {
        hasAttachments: result.hasAttachments,
        importance: result.importance,
      },
    });
    res.json(result);
  } catch (error) {
    await recordAudit(req, {
      action: 'message_viewed',
      graphMessageId: req.params.messageId,
      ...getErrorAudit(error),
    });
    handleOutlookError(res, error);
  }
});

router.delete('/messages/:messageId', async (req, res) => {
  try {
    const result = await OutlookService.deleteMessage(req.user, req.params.messageId);
    await recordAudit(req, {
      action: 'message_deleted',
      status: 'success',
      graphMessageId: result.messageId,
    });
    res.json(result);
  } catch (error) {
    await recordAudit(req, {
      action: 'message_deleted',
      graphMessageId: req.params.messageId,
      ...getErrorAudit(error),
    });
    handleOutlookError(res, error);
  }
});

router.post('/messages/:messageId/analyze', async (req, res) => {
  try {
    const result = await OutlookService.analyzeMessage(req.user, req.params.messageId);
    await recordAudit(req, {
      action: 'message_analyzed',
      status: 'success',
      graphMessageId: result.messageId,
      metadata: {
        analysisMode: result.insights?.mode,
      },
    });
    res.json(result);
  } catch (error) {
    await recordAudit(req, {
      action: 'message_analyzed',
      graphMessageId: req.params.messageId,
      ...getErrorAudit(error),
    });
    handleOutlookError(res, error);
  }
});

router.post('/messages/:messageId/drafts', async (req, res) => {
  try {
    const result = await OutlookService.createReplyDraft(req.user, req.params.messageId, req.body);
    await recordAudit(req, {
      action: 'draft_created',
      status: 'success',
      graphMessageId: result.sourceMessageId,
      graphDraftId: result.draftId,
      metadata: {
        tone: req.body?.tone || 'professional',
        hasInstructions: Boolean(req.body?.instructions),
      },
    });
    res.json(result);
  } catch (error) {
    await recordAudit(req, {
      action: 'draft_created',
      graphMessageId: req.params.messageId,
      ...getErrorAudit(error),
      metadata: {
        tone: req.body?.tone || 'professional',
        hasInstructions: Boolean(req.body?.instructions),
      },
    });
    handleOutlookError(res, error);
  }
});

router.post('/messages/:messageId/meeting-slots', async (req, res) => {
  try {
    const result = await OutlookService.proposeMeetingSlots(
      req.user,
      req.params.messageId,
      req.body,
    );
    await recordAudit(req, {
      action: 'meeting_slots_proposed',
      status: 'success',
      graphMessageId: result.messageId,
      metadata: {
        attendeeCount: result.attendees.length,
        suggestionCount: result.suggestions.length,
        durationMinutes: result.durationMinutes,
      },
    });
    res.json(result);
  } catch (error) {
    await recordAudit(req, {
      action: 'meeting_slots_proposed',
      graphMessageId: req.params.messageId,
      ...getErrorAudit(error),
    });
    handleOutlookError(res, error);
  }
});

router.post('/messages/:messageId/meetings', async (req, res) => {
  try {
    const result = await OutlookService.createTeamsMeeting(
      req.user,
      req.params.messageId,
      req.body,
    );
    await recordAudit(req, {
      action: 'meeting_created',
      status: 'success',
      graphMessageId: result.sourceMessageId,
      graphDraftId: result.draft?.id,
      metadata: {
        graphEventId: result.event?.id,
        attendeeCount: result.attendees.length,
        hasTeamsJoinUrl: Boolean(result.event?.onlineMeeting?.joinUrl),
      },
    });
    res.json(result);
  } catch (error) {
    await recordAudit(req, {
      action: 'meeting_created',
      graphMessageId: req.params.messageId,
      ...getErrorAudit(error),
    });
    handleOutlookError(res, error);
  }
});

module.exports = router;
