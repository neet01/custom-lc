import type { FilterQuery, Model } from 'mongoose';
import logger from '~/config/winston';
import type { ITeamsArchiveBackfillState } from '~/schema/teamsArchiveBackfillState';
import type { ITeamsArchiveConversation } from '~/schema/teamsArchiveConversation';
import type { ITeamsArchiveMessage } from '~/schema/teamsArchiveMessage';
import type { ITeamsArchiveSyncLease } from '~/schema/teamsArchiveSyncLease';
import type { ITeamsArchiveSyncJob } from '~/schema/teamsArchiveSyncJob';
import type {
  TeamsArchiveBackfillStateData,
  TeamsArchiveConversationData,
  TeamsArchiveMessageData,
  TeamsArchiveSyncJobData,
} from '~/types/teamsArchive';

export interface TeamsArchiveQueryOptions {
  limit?: number;
  offset?: number;
  sort?: Record<string, 1 | -1>;
}

interface TeamsArchiveSyncLeaseData {
  leaseKey: string;
  leaseType: 'user' | 'slot';
  ownerToken: string;
  user?: string;
  leaseExpiresAt: Date;
  lastHeartbeatAt?: Date;
}

export function createTeamsArchiveMethods(mongoose: typeof import('mongoose')) {
  async function getTeamsArchiveBackfillState(
    userId: string,
  ): Promise<ITeamsArchiveBackfillState | null> {
    try {
      const TeamsArchiveBackfillState = mongoose.models
        .TeamsArchiveBackfillState as Model<ITeamsArchiveBackfillState>;
      return (await TeamsArchiveBackfillState.findOne({ user: userId }).lean()) as
        | ITeamsArchiveBackfillState
        | null;
    } catch (error) {
      logger.error('[getTeamsArchiveBackfillState] Error finding Teams archive backfill state', error);
      throw new Error('Error finding Teams archive backfill state');
    }
  }

  async function upsertTeamsArchiveBackfillState(
    record: TeamsArchiveBackfillStateData,
  ): Promise<ITeamsArchiveBackfillState> {
    try {
      const TeamsArchiveBackfillState = mongoose.models
        .TeamsArchiveBackfillState as Model<ITeamsArchiveBackfillState>;

      return (await TeamsArchiveBackfillState.findOneAndUpdate(
        { user: record.user },
        { $set: record },
        { new: true, upsert: true, setDefaultsOnInsert: true },
      )) as ITeamsArchiveBackfillState;
    } catch (error) {
      logger.error('[upsertTeamsArchiveBackfillState] Error upserting Teams archive backfill state', error);
      throw new Error('Error upserting Teams archive backfill state');
    }
  }

  async function upsertTeamsArchiveConversation(
    record: TeamsArchiveConversationData,
  ): Promise<ITeamsArchiveConversation> {
    try {
      const TeamsArchiveConversation = mongoose.models
        .TeamsArchiveConversation as Model<ITeamsArchiveConversation>;

      return (await TeamsArchiveConversation.findOneAndUpdate(
        { user: record.user, graphChatId: record.graphChatId },
        { $set: record },
        { new: true, upsert: true, setDefaultsOnInsert: true },
      )) as ITeamsArchiveConversation;
    } catch (error) {
      logger.error('[upsertTeamsArchiveConversation] Error upserting Teams archive conversation', error);
      throw new Error('Error upserting Teams archive conversation');
    }
  }

  async function bulkUpsertTeamsArchiveMessages(
    records: TeamsArchiveMessageData[],
  ): Promise<number> {
    try {
      if (!Array.isArray(records) || records.length === 0) {
        return 0;
      }

      const TeamsArchiveMessage = mongoose.models.TeamsArchiveMessage as Model<ITeamsArchiveMessage>;
      const operations = records.map((record) => ({
        updateOne: {
          filter: { user: record.user, graphMessageId: record.graphMessageId },
          update: { $set: record },
          upsert: true,
        },
      }));

      const result = await TeamsArchiveMessage.bulkWrite(operations, { ordered: false });
      return result ? records.length : 0;
    } catch (error) {
      logger.error('[bulkUpsertTeamsArchiveMessages] Error upserting Teams archive messages', error);
      throw new Error('Error upserting Teams archive messages');
    }
  }

  async function createTeamsArchiveSyncJob(
    record: TeamsArchiveSyncJobData,
  ): Promise<ITeamsArchiveSyncJob> {
    try {
      const TeamsArchiveSyncJob = mongoose.models.TeamsArchiveSyncJob as Model<ITeamsArchiveSyncJob>;
      return (await TeamsArchiveSyncJob.create(record)) as ITeamsArchiveSyncJob;
    } catch (error) {
      logger.error('[createTeamsArchiveSyncJob] Error creating Teams archive sync job', error);
      throw new Error('Error creating Teams archive sync job');
    }
  }

  async function updateTeamsArchiveSyncJob(
    id: string,
    updates: Partial<TeamsArchiveSyncJobData>,
  ): Promise<ITeamsArchiveSyncJob | null> {
    try {
      const TeamsArchiveSyncJob = mongoose.models.TeamsArchiveSyncJob as Model<ITeamsArchiveSyncJob>;
      return (await TeamsArchiveSyncJob.findByIdAndUpdate(id, { $set: updates }, { new: true })) as
        | ITeamsArchiveSyncJob
        | null;
    } catch (error) {
      logger.error('[updateTeamsArchiveSyncJob] Error updating Teams archive sync job', error);
      throw new Error('Error updating Teams archive sync job');
    }
  }

  async function findTeamsArchiveConversations(
    filter: FilterQuery<ITeamsArchiveConversation> = {},
    options: TeamsArchiveQueryOptions = {},
  ): Promise<ITeamsArchiveConversation[]> {
    try {
      const TeamsArchiveConversation = mongoose.models
        .TeamsArchiveConversation as Model<ITeamsArchiveConversation>;
      const { limit = 50, offset = 0, sort = { lastMessageAt: -1, updatedAt: -1 } } = options;
      return (await TeamsArchiveConversation.find(filter)
        .sort(sort)
        .skip(offset)
        .limit(limit)
        .lean()) as ITeamsArchiveConversation[];
    } catch (error) {
      logger.error('[findTeamsArchiveConversations] Error finding Teams archive conversations', error);
      throw new Error('Error finding Teams archive conversations');
    }
  }

  async function findTeamsArchiveMessages(
    filter: FilterQuery<ITeamsArchiveMessage> = {},
    options: TeamsArchiveQueryOptions = {},
  ): Promise<ITeamsArchiveMessage[]> {
    try {
      const TeamsArchiveMessage = mongoose.models.TeamsArchiveMessage as Model<ITeamsArchiveMessage>;
      const { limit = 50, offset = 0, sort = { sentDateTime: -1, createdAt: -1 } } = options;
      return (await TeamsArchiveMessage.find(filter)
        .sort(sort)
        .skip(offset)
        .limit(limit)
        .lean()) as ITeamsArchiveMessage[];
    } catch (error) {
      logger.error('[findTeamsArchiveMessages] Error finding Teams archive messages', error);
      throw new Error('Error finding Teams archive messages');
    }
  }

  async function findLatestTeamsArchiveSyncJob(
    filter: FilterQuery<ITeamsArchiveSyncJob> = {},
  ): Promise<ITeamsArchiveSyncJob | null> {
    try {
      const TeamsArchiveSyncJob = mongoose.models.TeamsArchiveSyncJob as Model<ITeamsArchiveSyncJob>;
      return (await TeamsArchiveSyncJob.findOne(filter).sort({ createdAt: -1 }).lean()) as
        | ITeamsArchiveSyncJob
        | null;
    } catch (error) {
      logger.error('[findLatestTeamsArchiveSyncJob] Error finding Teams archive sync job', error);
      throw new Error('Error finding Teams archive sync job');
    }
  }

  async function findTeamsArchiveSyncJobById(id: string): Promise<ITeamsArchiveSyncJob | null> {
    try {
      const TeamsArchiveSyncJob = mongoose.models.TeamsArchiveSyncJob as Model<ITeamsArchiveSyncJob>;
      return (await TeamsArchiveSyncJob.findById(id).lean()) as ITeamsArchiveSyncJob | null;
    } catch (error) {
      logger.error('[findTeamsArchiveSyncJobById] Error finding Teams archive sync job by id', error);
      throw new Error('Error finding Teams archive sync job by id');
    }
  }

  async function countTeamsArchiveConversations(
    filter: FilterQuery<ITeamsArchiveConversation> = {},
  ): Promise<number> {
    try {
      const TeamsArchiveConversation = mongoose.models
        .TeamsArchiveConversation as Model<ITeamsArchiveConversation>;
      return await TeamsArchiveConversation.countDocuments(filter);
    } catch (error) {
      logger.error('[countTeamsArchiveConversations] Error counting Teams archive conversations', error);
      throw new Error('Error counting Teams archive conversations');
    }
  }

  async function countTeamsArchiveMessages(
    filter: FilterQuery<ITeamsArchiveMessage> = {},
  ): Promise<number> {
    try {
      const TeamsArchiveMessage = mongoose.models.TeamsArchiveMessage as Model<ITeamsArchiveMessage>;
      return await TeamsArchiveMessage.countDocuments(filter);
    } catch (error) {
      logger.error('[countTeamsArchiveMessages] Error counting Teams archive messages', error);
      throw new Error('Error counting Teams archive messages');
    }
  }

  async function updateTeamsArchiveConversation(
    id: string,
    updates: Partial<TeamsArchiveConversationData>,
  ): Promise<ITeamsArchiveConversation | null> {
    try {
      const TeamsArchiveConversation = mongoose.models
        .TeamsArchiveConversation as Model<ITeamsArchiveConversation>;
      return (await TeamsArchiveConversation.findByIdAndUpdate(id, { $set: updates }, { new: true })) as
        | ITeamsArchiveConversation
        | null;
    } catch (error) {
      logger.error('[updateTeamsArchiveConversation] Error updating Teams archive conversation', error);
      throw new Error('Error updating Teams archive conversation');
    }
  }

  async function acquireTeamsArchiveSyncLease(
    record: TeamsArchiveSyncLeaseData,
  ): Promise<ITeamsArchiveSyncLease | null> {
    try {
      const TeamsArchiveSyncLease = mongoose.models.TeamsArchiveSyncLease as Model<ITeamsArchiveSyncLease>;
      const now = new Date();

      return (await TeamsArchiveSyncLease.findOneAndUpdate(
        {
          leaseKey: record.leaseKey,
          $or: [{ leaseExpiresAt: { $lte: now } }, { ownerToken: record.ownerToken }],
        },
        {
          $set: {
            leaseType: record.leaseType,
            ownerToken: record.ownerToken,
            user: record.user,
            leaseExpiresAt: record.leaseExpiresAt,
            lastHeartbeatAt: record.lastHeartbeatAt || now,
          },
          $setOnInsert: {
            leaseKey: record.leaseKey,
          },
        },
        {
          new: true,
          upsert: true,
          setDefaultsOnInsert: true,
        },
      )) as ITeamsArchiveSyncLease | null;
    } catch (error) {
      if ((error as { code?: number })?.code === 11000) {
        return null;
      }

      logger.error('[acquireTeamsArchiveSyncLease] Error acquiring Teams archive sync lease', error);
      throw new Error('Error acquiring Teams archive sync lease');
    }
  }

  async function refreshTeamsArchiveSyncLease(
    leaseKey: string,
    ownerToken: string,
    leaseExpiresAt: Date,
  ): Promise<ITeamsArchiveSyncLease | null> {
    try {
      const TeamsArchiveSyncLease = mongoose.models.TeamsArchiveSyncLease as Model<ITeamsArchiveSyncLease>;
      return (await TeamsArchiveSyncLease.findOneAndUpdate(
        { leaseKey, ownerToken },
        {
          $set: {
            leaseExpiresAt,
            lastHeartbeatAt: new Date(),
          },
        },
        { new: true },
      )) as ITeamsArchiveSyncLease | null;
    } catch (error) {
      logger.error('[refreshTeamsArchiveSyncLease] Error refreshing Teams archive sync lease', error);
      throw new Error('Error refreshing Teams archive sync lease');
    }
  }

  async function releaseTeamsArchiveSyncLease(
    leaseKey: string,
    ownerToken: string,
  ): Promise<boolean> {
    try {
      const TeamsArchiveSyncLease = mongoose.models.TeamsArchiveSyncLease as Model<ITeamsArchiveSyncLease>;
      const result = await TeamsArchiveSyncLease.deleteOne({ leaseKey, ownerToken });
      return Boolean(result?.deletedCount);
    } catch (error) {
      logger.error('[releaseTeamsArchiveSyncLease] Error releasing Teams archive sync lease', error);
      throw new Error('Error releasing Teams archive sync lease');
    }
  }

  async function countActiveTeamsArchiveSyncLeases(
    filter: FilterQuery<ITeamsArchiveSyncLease> = {},
  ): Promise<number> {
    try {
      const TeamsArchiveSyncLease = mongoose.models.TeamsArchiveSyncLease as Model<ITeamsArchiveSyncLease>;
      return await TeamsArchiveSyncLease.countDocuments(filter);
    } catch (error) {
      logger.error('[countActiveTeamsArchiveSyncLeases] Error counting Teams archive sync leases', error);
      throw new Error('Error counting Teams archive sync leases');
    }
  }

  return {
    getTeamsArchiveBackfillState,
    upsertTeamsArchiveBackfillState,
    upsertTeamsArchiveConversation,
    bulkUpsertTeamsArchiveMessages,
    createTeamsArchiveSyncJob,
    updateTeamsArchiveSyncJob,
    findTeamsArchiveConversations,
    findTeamsArchiveMessages,
    findLatestTeamsArchiveSyncJob,
    findTeamsArchiveSyncJobById,
    countTeamsArchiveConversations,
    countTeamsArchiveMessages,
    updateTeamsArchiveConversation,
    acquireTeamsArchiveSyncLease,
    refreshTeamsArchiveSyncLease,
    releaseTeamsArchiveSyncLease,
    countActiveTeamsArchiveSyncLeases,
  };
}

export type TeamsArchiveMethods = ReturnType<typeof createTeamsArchiveMethods>;
