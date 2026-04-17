import type { FilterQuery, Model } from 'mongoose';
import logger from '~/config/winston';
import type { IUsage } from '~/schema/usage';
import type { UsageRecordData } from '~/types';

export interface UsageQueryOptions {
  limit?: number;
  offset?: number;
  sort?: Record<string, 1 | -1>;
}

export interface UsageSummaryQueryOptions {
  limit?: number;
  offset?: number;
}

export interface UsageUserSummary {
  userId: string;
  requestCount: number;
  inputTokens: number;
  outputTokens: number;
  totalTokens: number;
  cacheCreationTokens: number;
  cacheReadTokens: number;
  avgLatencyMs: number | null;
  firstSeenAt?: Date;
  lastSeenAt?: Date;
}

export interface UsageOverviewSummary {
  requestCount: number;
  inputTokens: number;
  outputTokens: number;
  totalTokens: number;
  cacheCreationTokens: number;
  cacheReadTokens: number;
  avgLatencyMs: number | null;
  activeUsers: number;
  firstSeenAt?: Date;
  lastSeenAt?: Date;
}

export function createUsageMethods(mongoose: typeof import('mongoose')) {
  async function createUsageRecords(records: UsageRecordData[]): Promise<IUsage[]> {
    if (!records.length) {
      return [];
    }

    try {
      const Usage = mongoose.models.Usage as Model<IUsage>;
      return (await Usage.insertMany(records, { ordered: false })) as IUsage[];
    } catch (error) {
      logger.error('[createUsageRecords] Error creating usage records', error);
      throw new Error('Error creating usage records');
    }
  }

  async function findUsageRecords(
    filter: FilterQuery<IUsage> = {},
    options: UsageQueryOptions = {},
  ): Promise<IUsage[]> {
    try {
      const Usage = mongoose.models.Usage as Model<IUsage>;
      const {
        limit = 50,
        offset = 0,
        sort = { createdAt: -1 },
      } = options;

      return (await Usage.find(filter).sort(sort).skip(offset).limit(limit).lean()) as IUsage[];
    } catch (error) {
      logger.error('[findUsageRecords] Error finding usage records', error);
      throw new Error('Error finding usage records');
    }
  }

  async function countUsageRecords(filter: FilterQuery<IUsage> = {}): Promise<number> {
    try {
      const Usage = mongoose.models.Usage as Model<IUsage>;
      return await Usage.countDocuments(filter);
    } catch (error) {
      logger.error('[countUsageRecords] Error counting usage records', error);
      throw new Error('Error counting usage records');
    }
  }

  async function summarizeUsageByUser(
    filter: FilterQuery<IUsage> = {},
    options: UsageSummaryQueryOptions = {},
  ): Promise<UsageUserSummary[]> {
    try {
      const Usage = mongoose.models.Usage as Model<IUsage>;
      const { limit = 50, offset = 0 } = options;

      return (await Usage.aggregate([
        { $match: filter },
        {
          $group: {
            _id: '$user',
            requestCount: { $sum: 1 },
            inputTokens: { $sum: '$inputTokens' },
            outputTokens: { $sum: '$outputTokens' },
            totalTokens: { $sum: '$totalTokens' },
            cacheCreationTokens: { $sum: { $ifNull: ['$cacheCreationTokens', 0] } },
            cacheReadTokens: { $sum: { $ifNull: ['$cacheReadTokens', 0] } },
            latencyTotal: { $sum: { $ifNull: ['$latencyMs', 0] } },
            latencyCount: {
              $sum: {
                $cond: [{ $ifNull: ['$latencyMs', false] }, 1, 0],
              },
            },
            firstSeenAt: { $min: '$createdAt' },
            lastSeenAt: { $max: '$createdAt' },
          },
        },
        { $sort: { totalTokens: -1, requestCount: -1, lastSeenAt: -1 } },
        { $skip: offset },
        { $limit: limit },
        {
          $project: {
            _id: 0,
            userId: { $toString: '$_id' },
            requestCount: 1,
            inputTokens: 1,
            outputTokens: 1,
            totalTokens: 1,
            cacheCreationTokens: 1,
            cacheReadTokens: 1,
            firstSeenAt: 1,
            lastSeenAt: 1,
            avgLatencyMs: {
              $cond: [
                { $gt: ['$latencyCount', 0] },
                { $divide: ['$latencyTotal', '$latencyCount'] },
                null,
              ],
            },
          },
        },
      ])) as UsageUserSummary[];
    } catch (error) {
      logger.error('[summarizeUsageByUser] Error summarizing usage by user', error);
      throw new Error('Error summarizing usage by user');
    }
  }

  async function summarizeUsageOverview(
    filter: FilterQuery<IUsage> = {},
  ): Promise<UsageOverviewSummary> {
    try {
      const Usage = mongoose.models.Usage as Model<IUsage>;
      const [result] = (await Usage.aggregate([
        { $match: filter },
        {
          $group: {
            _id: null,
            requestCount: { $sum: 1 },
            inputTokens: { $sum: '$inputTokens' },
            outputTokens: { $sum: '$outputTokens' },
            totalTokens: { $sum: '$totalTokens' },
            cacheCreationTokens: { $sum: { $ifNull: ['$cacheCreationTokens', 0] } },
            cacheReadTokens: { $sum: { $ifNull: ['$cacheReadTokens', 0] } },
            latencyTotal: { $sum: { $ifNull: ['$latencyMs', 0] } },
            latencyCount: {
              $sum: {
                $cond: [{ $ifNull: ['$latencyMs', false] }, 1, 0],
              },
            },
            activeUsersSet: { $addToSet: '$user' },
            firstSeenAt: { $min: '$createdAt' },
            lastSeenAt: { $max: '$createdAt' },
          },
        },
        {
          $project: {
            _id: 0,
            requestCount: 1,
            inputTokens: 1,
            outputTokens: 1,
            totalTokens: 1,
            cacheCreationTokens: 1,
            cacheReadTokens: 1,
            firstSeenAt: 1,
            lastSeenAt: 1,
            activeUsers: { $size: '$activeUsersSet' },
            avgLatencyMs: {
              $cond: [
                { $gt: ['$latencyCount', 0] },
                { $divide: ['$latencyTotal', '$latencyCount'] },
                null,
              ],
            },
          },
        },
      ])) as UsageOverviewSummary[];

      return (
        result ?? {
          requestCount: 0,
          inputTokens: 0,
          outputTokens: 0,
          totalTokens: 0,
          cacheCreationTokens: 0,
          cacheReadTokens: 0,
          avgLatencyMs: null,
          activeUsers: 0,
        }
      );
    } catch (error) {
      logger.error('[summarizeUsageOverview] Error summarizing usage overview', error);
      throw new Error('Error summarizing usage overview');
    }
  }

  return {
    createUsageRecords,
    findUsageRecords,
    countUsageRecords,
    summarizeUsageByUser,
    summarizeUsageOverview,
  };
}

export type UsageMethods = ReturnType<typeof createUsageMethods>;
