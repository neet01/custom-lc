const express = require('express');
const { logger } = require('@librechat/data-schemas');
const { recordCollectedUsage, getBalanceConfig, getTransactionsConfig } = require('@librechat/api');
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

function normalizeUsageEntry(entry) {
  const usage = entry?.usage;
  if (!usage || typeof usage !== 'object') {
    return null;
  }

  const input_tokens = Number(usage.input_tokens ?? usage.inputTokens ?? 0) || 0;
  const output_tokens = Number(usage.output_tokens ?? usage.outputTokens ?? 0) || 0;
  const total_tokens =
    Number(usage.total_tokens ?? usage.totalTokens ?? input_tokens + output_tokens) ||
    input_tokens + output_tokens;

  if (input_tokens <= 0 && output_tokens <= 0 && total_tokens <= 0) {
    return null;
  }

  return {
    context: entry?.context || 'outlook_message',
    usage: {
      input_tokens,
      output_tokens,
      total_tokens,
      model: usage.model,
      provider: usage.provider,
    },
  };
}

function extractUsageEntries(result) {
  if (!result || typeof result !== 'object') {
    return [];
  }
  const rawEntries = Array.isArray(result._usage) ? result._usage : [];
  if (Object.prototype.hasOwnProperty.call(result, '_usage')) {
    delete result._usage;
  }
  return rawEntries.map(normalizeUsageEntry).filter(Boolean);
}

async function recordOutlookUsage(req, result, { messageId, latencyMs = 0 } = {}) {
  const usageEntries = extractUsageEntries(result);
  if (usageEntries.length === 0) {
    return;
  }

  const userId = req.user?.id;
  const conversationId =
    result?.conversationId || result?.sourceMessageId || `outlook:${messageId || 'mailbox'}`;
  const balanceConfig = getBalanceConfig(req.config);
  const transactionsConfig = getTransactionsConfig(req.config);

  for (const entry of usageEntries) {
    await recordCollectedUsage(
      {
        spendTokens: db.spendTokens,
        spendStructuredTokens: db.spendStructuredTokens,
        pricing: { getMultiplier: db.getMultiplier, getCacheMultiplier: db.getCacheMultiplier },
        bulkWriteOps: { insertMany: db.bulkInsertTransactions, updateBalance: db.updateBalance },
        usagePersistence: { createUsageRecords: db.createUsageRecords },
      },
      {
        user: userId,
        conversationId: String(conversationId),
        collectedUsage: [entry.usage],
        context: entry.context,
        messageId,
        requestId: `${messageId || 'outlook'}:${entry.context}`,
        sessionId: req.sessionID,
        balance: balanceConfig,
        transactions: transactionsConfig,
        model: entry.usage.model || process.env.OUTLOOK_AI_MODEL_ID,
        provider: entry.usage.provider || process.env.OUTLOOK_AI_PROVIDER || 'bedrock',
        endpoint: req.baseUrl,
        source: 'tool',
        latencyMs,
      },
    ).catch((error) => {
      logger.error('[OutlookRoutes] Failed to persist Outlook model usage', {
        messageId,
        context: entry.context,
        error: error?.message || error,
      });
    });
  }
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
      search: req.query.search,
    });
    await recordAudit(req, {
      action: 'mailbox_listed',
      status: 'success',
      metadata: {
        folder: typeof req.query.folder === 'string' ? req.query.folder : 'inbox',
        inboxView: typeof req.query.inboxView === 'string' ? req.query.inboxView : 'focused',
        searched: typeof req.query.search === 'string' && req.query.search.trim().length > 0,
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
        searched: typeof req.query.search === 'string' && req.query.search.trim().length > 0,
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

router.patch('/messages/:messageId/read', async (req, res) => {
  if (typeof req.body?.isRead !== 'boolean') {
    return res.status(400).json({ message: 'isRead must be a boolean' });
  }

  try {
    const result = await OutlookService.updateMessageReadState(
      req.user,
      req.params.messageId,
      req.body.isRead,
    );
    await recordAudit(req, {
      action: result.isRead ? 'message_marked_read' : 'message_marked_unread',
      status: 'success',
      graphMessageId: result.messageId,
    });
    res.json(result);
  } catch (error) {
    await recordAudit(req, {
      action: req.body.isRead ? 'message_marked_read' : 'message_marked_unread',
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
  const startedAt = Date.now();
  try {
    const result = await OutlookService.analyzeMessage(req.user, req.params.messageId);
    await recordOutlookUsage(req, result, {
      messageId: result.messageId || req.params.messageId,
      latencyMs: Date.now() - startedAt,
    });
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

router.post('/messages/analyze-selection', async (req, res) => {
  const startedAt = Date.now();
  const messageIds = Array.isArray(req.body?.messageIds)
    ? req.body.messageIds.filter((messageId) => typeof messageId === 'string')
    : [];

  if (messageIds.length === 0) {
    return res.status(400).json({ message: 'messageIds must contain at least one message id' });
  }

  try {
    const result = await OutlookService.analyzeSelectedMessages(req.user, messageIds);
    await recordOutlookUsage(req, result, {
      messageId: result.messageIds?.[0] || 'selection',
      latencyMs: Date.now() - startedAt,
    });
    await recordAudit(req, {
      action: 'selection_analyzed',
      status: 'success',
      graphMessageId: result.messageIds?.[0],
      metadata: {
        messageCount: result.messageCount,
      },
    });
    res.json(result);
  } catch (error) {
    await recordAudit(req, {
      action: 'selection_analyzed',
      graphMessageId: messageIds[0],
      ...getErrorAudit(error),
      metadata: {
        messageCount: messageIds.length,
      },
    });
    handleOutlookError(res, error);
  }
});

router.post('/messages/:messageId/drafts', async (req, res) => {
  const startedAt = Date.now();
  try {
    const result = await OutlookService.createReplyDraft(req.user, req.params.messageId, req.body);
    await recordOutlookUsage(req, result, {
      messageId: result.sourceMessageId || req.params.messageId,
      latencyMs: Date.now() - startedAt,
    });
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

router.post('/daily-brief', async (req, res) => {
  const startedAt = Date.now();
  try {
    const result = await OutlookService.generateDailyBrief(req.user, { hours: 24 });
    await recordOutlookUsage(req, result, {
      messageId: result.messageIds?.[0] || 'daily-brief',
      latencyMs: Date.now() - startedAt,
    });
    await recordAudit(req, {
      action: 'daily_brief_generated',
      status: 'success',
      graphMessageId: result.messageIds?.[0],
      metadata: {
        emailCount: result.emailCount,
        meetingCount: result.meetingCount,
        windowStart: result.windowStart,
        windowEnd: result.windowEnd,
      },
    });
    res.json(result);
  } catch (error) {
    await recordAudit(req, {
      action: 'daily_brief_generated',
      ...getErrorAudit(error),
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
  const startedAt = Date.now();
  try {
    const result = await OutlookService.createTeamsMeeting(
      req.user,
      req.params.messageId,
      req.body,
    );
    await recordOutlookUsage(req, result, {
      messageId: result.sourceMessageId || req.params.messageId,
      latencyMs: Date.now() - startedAt,
    });
    await recordAudit(req, {
      action: 'meeting_created',
      status: 'success',
      graphMessageId: result.sourceMessageId,
      graphDraftId: result.meetingDraft?.id || result.draft?.id,
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
