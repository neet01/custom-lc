const express = require('express');
const { isValidObjectIdString, logger, SystemCapabilities } = require('@librechat/data-schemas');
const { requireCapability } = require('~/server/middleware/roles/capabilities');
const { requireJwtAuth } = require('~/server/middleware');
const db = require('~/models');

const router = express.Router();

const requireAdminAccess = requireCapability(SystemCapabilities.ACCESS_ADMIN);
const requireReadUsage = requireCapability(SystemCapabilities.READ_USAGE);

const USER_FIELDS = '_id name username email avatar role provider';
const ACTIONS = new Set([
  'mailbox_listed',
  'calendar_viewed',
  'calendar_event_created',
  'calendar_event_updated',
  'calendar_event_deleted',
  'message_viewed',
  'message_deleted',
  'message_analyzed',
  'draft_created',
  'meeting_slots_proposed',
  'meeting_created',
]);
const STATUSES = new Set(['success', 'failure']);

function parsePagination(query) {
  const limit = Math.min(Math.max(Number(query.limit) || 50, 1), 200);
  const offset = Math.max(Number(query.offset) || 0, 0);
  return { limit, offset };
}

function buildFilter(query) {
  const filter = {};

  if (typeof query.user_id === 'string' && query.user_id) {
    if (!isValidObjectIdString(query.user_id)) {
      return { filter: null, error: 'Invalid user ID format' };
    }
    filter.user = query.user_id;
  }

  if (typeof query.action === 'string' && query.action) {
    if (!ACTIONS.has(query.action)) {
      return { filter: null, error: 'Invalid Outlook audit action' };
    }
    filter.action = query.action;
  }

  if (typeof query.status === 'string' && query.status) {
    if (!STATUSES.has(query.status)) {
      return { filter: null, error: 'Invalid Outlook audit status' };
    }
    filter.status = query.status;
  }

  if (typeof query.message_id === 'string' && query.message_id) {
    filter.graphMessageId = query.message_id;
  }

  return { filter };
}

function mapAudit(record, usersById) {
  const userId = record.user?.toString() ?? '';
  const user = usersById.get(userId);

  return {
    id: record._id?.toString() ?? '',
    userId,
    actorName: user?.name || user?.username || user?.email || '',
    actorEmail: user?.email || '',
    actorAvatar: user?.avatar || '',
    actorRole: user?.role || 'USER',
    action: record.action,
    status: record.status,
    graphMessageId: record.graphMessageId,
    graphConversationId: record.graphConversationId,
    graphDraftId: record.graphDraftId,
    errorCode: record.errorCode,
    errorMessage: record.errorMessage,
    metadata: record.metadata,
    createdAt: record.createdAt?.toISOString(),
    updatedAt: record.updatedAt?.toISOString(),
  };
}

router.use(requireJwtAuth, requireAdminAccess);

router.get('/', requireReadUsage, async (req, res) => {
  try {
    const { limit, offset } = parsePagination(req.query);
    const { filter, error } = buildFilter(req.query);
    if (!filter) {
      return res.status(400).json({ error });
    }

    const [audits, total] = await Promise.all([
      db.findOutlookAudits(filter, { limit, offset, sort: { createdAt: -1 } }),
      db.countOutlookAudits(filter),
    ]);

    const userIds = [...new Set(audits.map((audit) => audit.user?.toString()).filter(Boolean))];
    const users =
      userIds.length > 0
        ? await db.findUsers({ _id: { $in: userIds } }, USER_FIELDS, { limit: userIds.length })
        : [];

    const usersById = new Map(
      users.map((user) => [
        user._id?.toString() ?? '',
        {
          name: user.name ?? '',
          username: user.username ?? '',
          email: user.email ?? '',
          avatar: user.avatar ?? '',
          role: user.role ?? 'USER',
        },
      ]),
    );

    return res.status(200).json({
      audits: audits.map((audit) => mapAudit(audit, usersById)),
      total,
      limit,
      offset,
    });
  } catch (error) {
    logger.error('[adminOutlookAudit] list error:', error);
    return res.status(500).json({ error: 'Failed to list Outlook audit records' });
  }
});

module.exports = router;
