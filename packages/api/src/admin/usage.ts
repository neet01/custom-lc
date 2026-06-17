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
  getValueKey?: (model: string, endpoint?: string) => string | undefined;
  getMultiplier: (params: {
    model?: string;
    valueKey?: string;
    endpoint?: string;
    tokenType?: 'prompt' | 'completion';
    inputTokenCount?: number;
    endpointTokenConfig?: Record<string, Record<string, number>>;
  }) => number;
  getCacheMultiplier: (params: {
    valueKey?: string;
    cacheType?: 'write' | 'read';
    model?: string;
    endpoint?: string;
    endpointTokenConfig?: Record<string, Record<string, number>>;
  }) => number | null;
  resolveFinanceUserOrgMetadata?: (
    users: IUser[],
    requester?: IUser,
  ) => Promise<Map<string, FinanceUserOrgMetadata>>;
}

const USER_SUMMARY_FIELDS = '_id name username email avatar role provider openidId idOnTheSource';
const DEFAULT_SUMMARY_DAYS = 30;
const MAX_SUMMARY_DAYS = 365;
const USD_PER_MILLION_TOKENS = 1_000_000;
const CSV_PRICING_BASIS = 'repo_model_pricing_usd_per_1m_tokens';

export type FinanceUserOrgMetadata = {
  graphUserId?: string;
  team?: string;
  role?: string;
  company?: string;
  officeLocation?: string;
};

type EstimatedUsageCost = {
  inputCostUsd: number;
  outputCostUsd: number;
  cacheWriteCostUsd: number;
  cacheReadCostUsd: number;
  totalCostUsd: number;
  priced: boolean;
};

type FinanceReportRow = {
  userId: string;
  name: string;
  username: string;
  email: string;
  role: string;
  provider: string;
  graphUserId: string;
  orgTeam: string;
  orgRole: string;
  orgCompany: string;
  orgOfficeLocation: string;
  requestCount: number;
  pricedRequestCount: number;
  unpricedRequestCount: number;
  totalTokens: number;
  inputTokens: number;
  outputTokens: number;
  cacheCreationTokens: number;
  cacheReadTokens: number;
  unpricedTotalTokens: number;
  estimatedInputCostUsd: number;
  estimatedOutputCostUsd: number;
  estimatedCacheWriteCostUsd: number;
  estimatedCacheReadCostUsd: number;
  estimatedTotalCostUsd: number;
  modelsUsed: string[];
  firstSeenAt?: string;
  lastSeenAt?: string;
};

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

function formatCost(value: number): string {
  return value.toFixed(6);
}

function csvEscape(value: string | number | null | undefined): string {
  const stringValue = value == null ? '' : String(value);
  if (!/[",\n]/.test(stringValue)) {
    return stringValue;
  }

  return `"${stringValue.replace(/"/g, '""')}"`;
}

function buildWindow(query: ServerRequest['query']) {
  const days = parseDays(query.days);
  const windowEnd = new Date();
  const windowStart = new Date(windowEnd.getTime() - days * 24 * 60 * 60 * 1000);
  return { days, windowStart, windowEnd };
}

function estimateUsageCost(record: IUsage, deps: AdminUsageDeps): EstimatedUsageCost {
  const model = record.model;
  if (!model) {
    return {
      inputCostUsd: 0,
      outputCostUsd: 0,
      cacheWriteCostUsd: 0,
      cacheReadCostUsd: 0,
      totalCostUsd: 0,
      priced: false,
    };
  }

  const valueKey = deps.getValueKey?.(model, record.endpoint);
  if (deps.getValueKey != null && !valueKey) {
    return {
      inputCostUsd: 0,
      outputCostUsd: 0,
      cacheWriteCostUsd: 0,
      cacheReadCostUsd: 0,
      totalCostUsd: 0,
      priced: false,
    };
  }

  const totalPromptTokens =
    (record.inputTokens ?? 0) + (record.cacheCreationTokens ?? 0) + (record.cacheReadTokens ?? 0);
  const promptRate = deps.getMultiplier({
    model,
    valueKey,
    endpoint: record.endpoint,
    tokenType: 'prompt',
    inputTokenCount: totalPromptTokens,
  });
  const completionRate = deps.getMultiplier({
    model,
    valueKey,
    endpoint: record.endpoint,
    tokenType: 'completion',
    inputTokenCount: totalPromptTokens,
  });
  const cacheWriteRate =
    deps.getCacheMultiplier({
      valueKey,
      model,
      endpoint: record.endpoint,
      cacheType: 'write',
    }) ?? promptRate;
  const cacheReadRate =
    deps.getCacheMultiplier({
      valueKey,
      model,
      endpoint: record.endpoint,
      cacheType: 'read',
    }) ?? promptRate;

  const inputCostUsd = ((record.inputTokens ?? 0) * promptRate) / USD_PER_MILLION_TOKENS;
  const outputCostUsd = ((record.outputTokens ?? 0) * completionRate) / USD_PER_MILLION_TOKENS;
  const cacheWriteCostUsd =
    ((record.cacheCreationTokens ?? 0) * cacheWriteRate) / USD_PER_MILLION_TOKENS;
  const cacheReadCostUsd =
    ((record.cacheReadTokens ?? 0) * cacheReadRate) / USD_PER_MILLION_TOKENS;

  return {
    inputCostUsd,
    outputCostUsd,
    cacheWriteCostUsd,
    cacheReadCostUsd,
    totalCostUsd: inputCostUsd + outputCostUsd + cacheWriteCostUsd + cacheReadCostUsd,
    priced: true,
  };
}

function buildFinanceReportCsv(
  rows: FinanceReportRow[],
  totals: FinanceReportRow,
  metadata: {
    generatedAt: string;
    days: number;
    windowStart: string;
    windowEnd: string;
  },
) {
  const headers = [
    'report_generated_at',
    'window_days',
    'window_start',
    'window_end',
    'user_id',
    'name',
    'username',
    'email',
    'role',
    'provider',
    'graph_user_id',
    'org_team',
    'org_role',
    'org_company',
    'org_office_location',
    'request_count',
    'priced_request_count',
    'unpriced_request_count',
    'total_tokens',
    'input_tokens',
    'output_tokens',
    'cache_write_tokens',
    'cache_read_tokens',
    'unpriced_total_tokens',
    'estimated_input_cost_usd',
    'estimated_output_cost_usd',
    'estimated_cache_write_cost_usd',
    'estimated_cache_read_cost_usd',
    'estimated_total_cost_usd',
    'models_used',
    'first_seen_at',
    'last_seen_at',
    'pricing_basis',
  ];

  const serializeRow = (row: FinanceReportRow) =>
    [
      metadata.generatedAt,
      metadata.days,
      metadata.windowStart,
      metadata.windowEnd,
      row.userId,
      row.name,
      row.username,
      row.email,
      row.role,
      row.provider,
      row.graphUserId,
      row.orgTeam,
      row.orgRole,
      row.orgCompany,
      row.orgOfficeLocation,
      row.requestCount,
      row.pricedRequestCount,
      row.unpricedRequestCount,
      row.totalTokens,
      row.inputTokens,
      row.outputTokens,
      row.cacheCreationTokens,
      row.cacheReadTokens,
      row.unpricedTotalTokens,
      formatCost(row.estimatedInputCostUsd),
      formatCost(row.estimatedOutputCostUsd),
      formatCost(row.estimatedCacheWriteCostUsd),
      formatCost(row.estimatedCacheReadCostUsd),
      formatCost(row.estimatedTotalCostUsd),
      row.modelsUsed.join(' | '),
      row.firstSeenAt ?? '',
      row.lastSeenAt ?? '',
      CSV_PRICING_BASIS,
    ]
      .map(csvEscape)
      .join(',');

  return [headers.join(','), ...rows.map(serializeRow), serializeRow(totals)].join('\n');
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

      const { days, windowStart, windowEnd } = buildWindow(req.query);
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

  async function exportFinanceReport(req: ServerRequest, res: Response) {
    if (!isEnabled(process.env[USAGE_TRACKING_ENABLED])) {
      return res.status(503).json({ error: 'Usage tracking is disabled' });
    }

    try {
      const { filter, error } = buildUsageFilter(req.query);
      if (!filter) {
        return res.status(400).json({ error });
      }

      const { days, windowStart, windowEnd } = buildWindow(req.query);
      filter.createdAt = { $gte: windowStart, $lte: windowEnd };

      const overviewSummary = await deps.summarizeUsageOverview(filter);
      const activeUsers = Math.max(overviewSummary.activeUsers, 1);
      const [usageByUser, usageRecordCount] = await Promise.all([
        deps.summarizeUsageByUser(filter, { limit: activeUsers, offset: 0 }),
        deps.countUsageRecords(filter),
      ]);

      const usageRecords =
        usageRecordCount > 0
          ? await deps.findUsageRecords(filter, {
              limit: usageRecordCount,
              offset: 0,
              sort: { createdAt: 1 },
            })
          : [];

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
            role: user.role ?? 'USER',
            provider: user.provider ?? 'local',
          },
        ]),
      );

      const costByUser = new Map<
        string,
        Omit<
          FinanceReportRow,
          | 'name'
          | 'username'
          | 'email'
          | 'role'
          | 'provider'
          | 'graphUserId'
          | 'orgTeam'
          | 'orgRole'
          | 'orgCompany'
          | 'orgOfficeLocation'
          | 'requestCount'
          | 'totalTokens'
          | 'inputTokens'
          | 'outputTokens'
          | 'cacheCreationTokens'
          | 'cacheReadTokens'
          | 'firstSeenAt'
          | 'lastSeenAt'
        >
      >();

      for (const record of usageRecords) {
        const userId = record.user?.toString() ?? '';
        if (!userId) {
          continue;
        }

        let current = costByUser.get(userId);
        if (!current) {
          current = {
            userId,
            pricedRequestCount: 0,
            unpricedRequestCount: 0,
            unpricedTotalTokens: 0,
            estimatedInputCostUsd: 0,
            estimatedOutputCostUsd: 0,
            estimatedCacheWriteCostUsd: 0,
            estimatedCacheReadCostUsd: 0,
            estimatedTotalCostUsd: 0,
            modelsUsed: [],
          };
          costByUser.set(userId, current);
        }

        const estimatedCost = estimateUsageCost(record, deps);
        const modelLabel = record.model || record.provider || record.endpoint || 'unknown';
        if (!current.modelsUsed.includes(modelLabel)) {
          current.modelsUsed.push(modelLabel);
        }

        if (estimatedCost.priced) {
          current.pricedRequestCount += 1;
          current.estimatedInputCostUsd += estimatedCost.inputCostUsd;
          current.estimatedOutputCostUsd += estimatedCost.outputCostUsd;
          current.estimatedCacheWriteCostUsd += estimatedCost.cacheWriteCostUsd;
          current.estimatedCacheReadCostUsd += estimatedCost.cacheReadCostUsd;
          current.estimatedTotalCostUsd += estimatedCost.totalCostUsd;
        } else {
          current.unpricedRequestCount += 1;
          current.unpricedTotalTokens += record.totalTokens ?? 0;
        }
      }

      const orgMetadataByUserId =
        deps.resolveFinanceUserOrgMetadata != null
          ? await deps.resolveFinanceUserOrgMetadata(users, req.user)
          : new Map<string, FinanceUserOrgMetadata>();

      const rows: FinanceReportRow[] = usageByUser.map((item) => {
        const user = usersById.get(item.userId);
        const cost = costByUser.get(item.userId);
        const orgMetadata = orgMetadataByUserId.get(item.userId);
        return {
          userId: item.userId,
          name: user?.name ?? '',
          username: user?.username ?? '',
          email: user?.email ?? '',
          role: user?.role ?? 'USER',
          provider: user?.provider ?? 'local',
          graphUserId: orgMetadata?.graphUserId ?? '',
          orgTeam: orgMetadata?.team ?? '',
          orgRole: orgMetadata?.role ?? '',
          orgCompany: orgMetadata?.company ?? '',
          orgOfficeLocation: orgMetadata?.officeLocation ?? '',
          requestCount: item.requestCount,
          pricedRequestCount: cost?.pricedRequestCount ?? 0,
          unpricedRequestCount: cost?.unpricedRequestCount ?? item.requestCount,
          totalTokens: item.totalTokens,
          inputTokens: item.inputTokens,
          outputTokens: item.outputTokens,
          cacheCreationTokens: item.cacheCreationTokens,
          cacheReadTokens: item.cacheReadTokens,
          unpricedTotalTokens: cost?.unpricedTotalTokens ?? item.totalTokens,
          estimatedInputCostUsd: cost?.estimatedInputCostUsd ?? 0,
          estimatedOutputCostUsd: cost?.estimatedOutputCostUsd ?? 0,
          estimatedCacheWriteCostUsd: cost?.estimatedCacheWriteCostUsd ?? 0,
          estimatedCacheReadCostUsd: cost?.estimatedCacheReadCostUsd ?? 0,
          estimatedTotalCostUsd: cost?.estimatedTotalCostUsd ?? 0,
          modelsUsed: [...(cost?.modelsUsed ?? [])].sort(),
          firstSeenAt: item.firstSeenAt?.toISOString(),
          lastSeenAt: item.lastSeenAt?.toISOString(),
        };
      });

      const totals: FinanceReportRow = {
        userId: 'TOTAL',
        name: 'Totals',
        username: '',
        email: '',
        role: '',
        provider: '',
        graphUserId: '',
        orgTeam: '',
        orgRole: '',
        orgCompany: '',
        orgOfficeLocation: '',
        requestCount: overviewSummary.requestCount,
        pricedRequestCount: rows.reduce((sum, row) => sum + row.pricedRequestCount, 0),
        unpricedRequestCount: rows.reduce((sum, row) => sum + row.unpricedRequestCount, 0),
        totalTokens: overviewSummary.totalTokens,
        inputTokens: overviewSummary.inputTokens,
        outputTokens: overviewSummary.outputTokens,
        cacheCreationTokens: overviewSummary.cacheCreationTokens,
        cacheReadTokens: overviewSummary.cacheReadTokens,
        unpricedTotalTokens: rows.reduce((sum, row) => sum + row.unpricedTotalTokens, 0),
        estimatedInputCostUsd: rows.reduce((sum, row) => sum + row.estimatedInputCostUsd, 0),
        estimatedOutputCostUsd: rows.reduce((sum, row) => sum + row.estimatedOutputCostUsd, 0),
        estimatedCacheWriteCostUsd: rows.reduce(
          (sum, row) => sum + row.estimatedCacheWriteCostUsd,
          0,
        ),
        estimatedCacheReadCostUsd: rows.reduce(
          (sum, row) => sum + row.estimatedCacheReadCostUsd,
          0,
        ),
        estimatedTotalCostUsd: rows.reduce((sum, row) => sum + row.estimatedTotalCostUsd, 0),
        modelsUsed: [],
        firstSeenAt: overviewSummary.firstSeenAt?.toISOString(),
        lastSeenAt: overviewSummary.lastSeenAt?.toISOString(),
      };

      const generatedAt = new Date().toISOString();
      const csv = buildFinanceReportCsv(rows, totals, {
        generatedAt,
        days,
        windowStart: windowStart.toISOString(),
        windowEnd: windowEnd.toISOString(),
      });

      res.setHeader('Content-Type', 'text/csv; charset=utf-8');
      res.setHeader(
        'Content-Disposition',
        `attachment; filename="cortex-finance-usage-${days}d-${generatedAt.slice(0, 10)}.csv"`,
      );

      return res.status(200).send(csv);
    } catch (error) {
      logger.error('[adminUsage] exportFinanceReport error:', error);
      return res.status(500).json({ error: 'Failed to export finance usage report' });
    }
  }

  return { listUsage, getUsageSummary, exportFinanceReport };
}
