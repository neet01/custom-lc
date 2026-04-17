import { logger, isValidObjectIdString } from '@librechat/data-schemas';
import type {
  IUser,
  IUsage,
  AdminUsageListItem,
  AdminUsageOverview,
  AdminUsageSummaryItem,
  UsageOverviewSummary,
  UsageUserSummary,
} from '@librechat/data-schemas';
import type { FilterQuery } from 'mongoose';
import type { Response } from 'express';
import type { ServerRequest } from '~/types/http';
import { isEnabled } from '~/utils/common';
import { parsePagination } from './pagination';
import { USAGE_TRACKING_ENABLED } from '~/usage';

export interface AdminUsageDeps {
  findUsageRecords: (
    filter?: FilterQuery<IUsage>,
    options?: {
      limit?: number;
      offset?: number;
      sort?: Record<string, 1 | -1>;
    },
  ) => Promise<IUsage[]>;
  countUsageRecords: (filter?: FilterQuery<IUsage>) => Promise<number>;
  summarizeUsageByUser: (
    filter?: FilterQuery<IUsage>,
    options?: {
      limit?: number;
      offset?: number;
    },
  ) => Promise<UsageUserSummary[]>;
  summarizeUsageOverview: (filter?: FilterQuery<IUsage>) => Promise<UsageOverviewSummary>;
  findUsers: (
    searchCriteria: FilterQuery<IUser>,
    fieldsToSelect?: string | string[] | null,
    options?: { limit?: number; offset?: number; sort?: Record<string, 1 | -1> },
  ) => Promise<IUser[]>;
}

const USER_SUMMARY_FIELDS = '_id name username email avatar role provider';
const DEFAULT_SUMMARY_DAYS = 30;
const MAX_SUMMARY_DAYS = 365;

function mapUsageRecord(record: IUsage): AdminUsageListItem {
  return {
    id: record._id?.toString() ?? '',
    userId: record.user?.toString() ?? '',
    conversationId: record.conversationId,
    messageId: record.messageId,
    requestId: record.requestId,
    sessionId: record.sessionId,
    model: record.model,
    provider: record.provider,
    endpoint: record.endpoint,
    context: record.context,
    source: record.source,
    inputTokens: record.inputTokens,
    outputTokens: record.outputTokens,
    totalTokens: record.totalTokens,
    cacheCreationTokens: record.cacheCreationTokens,
    cacheReadTokens: record.cacheReadTokens,
    latencyMs: record.latencyMs,
    createdAt: record.createdAt?.toISOString(),
    updatedAt: record.updatedAt?.toISOString(),
  };
}

function parseDays(value: unknown): number {
  const parsed = typeof value === 'string' ? parseInt(value, 10) : NaN;
  if (Number.isNaN(parsed)) {
    return DEFAULT_SUMMARY_DAYS;
  }

  return Math.min(Math.max(parsed, 1), MAX_SUMMARY_DAYS);
}

function buildUsageFilter(query: ServerRequest['query']) {
  const filter: FilterQuery<IUsage> = {};

  const userId = typeof query.user_id === 'string' ? query.user_id : undefined;
  if (userId) {
    if (!isValidObjectIdString(userId)) {
      return { filter: null, error: 'Invalid user ID format' as const };
    }
    filter.user = userId;
  }

  const conversationId = typeof query.conversation_id === 'string' ? query.conversation_id : undefined;
  if (conversationId) {
    filter.conversationId = conversationId;
  }

  const context = typeof query.context === 'string' ? query.context : undefined;
  if (context) {
    filter.context = context;
  }

  const source = typeof query.source === 'string' ? query.source : undefined;
  if (source) {
    filter.source = source;
  }

  return { filter };
}

function mapOverview(
  summary: UsageOverviewSummary,
  windowStart: Date,
  windowEnd: Date,
): AdminUsageOverview {
  return {
    requestCount: summary.requestCount,
    inputTokens: summary.inputTokens,
    outputTokens: summary.outputTokens,
    totalTokens: summary.totalTokens,
    cacheCreationTokens: summary.cacheCreationTokens,
    cacheReadTokens: summary.cacheReadTokens,
    avgLatencyMs: summary.avgLatencyMs,
    activeUsers: summary.activeUsers,
    firstSeenAt: summary.firstSeenAt?.toISOString(),
    lastSeenAt: summary.lastSeenAt?.toISOString(),
    windowStart: windowStart.toISOString(),
    windowEnd: windowEnd.toISOString(),
  };
}

export function createAdminUsageHandlers(deps: AdminUsageDeps) {
  async function listUsage(req: ServerRequest, res: Response) {
    if (!isEnabled(process.env[USAGE_TRACKING_ENABLED])) {
      return res.status(503).json({ error: 'Usage tracking is disabled' });
    }

    try {
      const { limit, offset } = parsePagination(req.query);
      const { filter, error } = buildUsageFilter(req.query);
      if (!filter) {
        return res.status(400).json({ error });
      }

      const [records, total] = await Promise.all([
        deps.findUsageRecords(filter, { limit, offset, sort: { createdAt: -1 } }),
        deps.countUsageRecords(filter),
      ]);

      return res.status(200).json({
        usage: records.map(mapUsageRecord),
        total,
        limit,
        offset,
      });
    } catch (error) {
      logger.error('[adminUsage] listUsage error:', error);
      return res.status(500).json({ error: 'Failed to list usage records' });
    }
  }

  async function getUsageSummary(req: ServerRequest, res: Response) {
    if (!isEnabled(process.env[USAGE_TRACKING_ENABLED])) {
      return res.status(503).json({ error: 'Usage tracking is disabled' });
    }

    try {
      const { limit, offset } = parsePagination(req.query);
      const { filter, error } = buildUsageFilter(req.query);
      if (!filter) {
        return res.status(400).json({ error });
      }

      const days = parseDays(req.query.days);
      const windowEnd = new Date();
      const windowStart = new Date(windowEnd.getTime() - days * 24 * 60 * 60 * 1000);
      filter.createdAt = { $gte: windowStart, $lte: windowEnd };

      const [overviewSummary, usageByUser] = await Promise.all([
        deps.summarizeUsageOverview(filter),
        deps.summarizeUsageByUser(filter, { limit, offset }),
      ]);

      const userIds = usageByUser.map((item) => item.userId).filter(Boolean);
      const users =
        userIds.length > 0
          ? await deps.findUsers({ _id: { $in: userIds } }, USER_SUMMARY_FIELDS, {
              limit: userIds.length,
            })
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
            provider: user.provider ?? 'local',
          },
        ]),
      );

      const summary: AdminUsageSummaryItem[] = usageByUser.map((item) => {
        const user = usersById.get(item.userId);
        return {
          userId: item.userId,
          name: user?.name ?? '',
          username: user?.username ?? '',
          email: user?.email ?? '',
          avatar: user?.avatar ?? '',
          role: user?.role ?? 'USER',
          provider: user?.provider ?? 'local',
          requestCount: item.requestCount,
          inputTokens: item.inputTokens,
          outputTokens: item.outputTokens,
          totalTokens: item.totalTokens,
          cacheCreationTokens: item.cacheCreationTokens,
          cacheReadTokens: item.cacheReadTokens,
          avgLatencyMs: item.avgLatencyMs,
          firstSeenAt: item.firstSeenAt?.toISOString(),
          lastSeenAt: item.lastSeenAt?.toISOString(),
        };
      });

      return res.status(200).json({
        overview: mapOverview(overviewSummary, windowStart, windowEnd),
        users: summary,
        total: overviewSummary.activeUsers,
        limit,
        offset,
        days,
      });
    } catch (error) {
      logger.error('[adminUsage] getUsageSummary error:', error);
      return res.status(500).json({ error: 'Failed to summarize usage records' });
    }
  }

  return { listUsage, getUsageSummary };
}
