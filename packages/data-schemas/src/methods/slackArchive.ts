import type { FilterQuery, Model } from 'mongoose';
import logger from '~/config/winston';
import type { ISlackArchiveConversation } from '~/schema/slackArchiveConversation';
import type { ISlackIdentityLink } from '~/schema/slackIdentityLink';
import type { ISlackArchiveMessage } from '~/schema/slackArchiveMessage';
import type { ISlackArchiveSyncJob } from '~/schema/slackArchiveSyncJob';
import type { ISlackArchiveSyncLease } from '~/schema/slackArchiveSyncLease';
import type { ISlackWorkspaceInstall } from '~/schema/slackWorkspaceInstall';
import type {
  SlackArchiveConversationData,
  SlackArchiveMessageData,
  SlackArchiveSyncJobData,
} from '~/types/slackArchive';
import type { SlackIdentityLinkData } from '~/types/slackIdentityLink';
import type { SlackWorkspaceInstallData } from '~/types/slackWorkspaceInstall';

export interface SlackArchiveQueryOptions {
  limit?: number;
  offset?: number;
  sort?: Record<string, 1 | -1>;
}

interface SlackArchiveSyncLeaseData {
  leaseKey: string;
  leaseType: 'user' | 'slot';
  ownerToken: string;
  user?: string;
  leaseExpiresAt: Date;
  lastHeartbeatAt?: Date;
}

export function createSlackArchiveMethods(mongoose: typeof import('mongoose')) {
  async function upsertSlackWorkspaceInstall(
    record: SlackWorkspaceInstallData,
  ): Promise<ISlackWorkspaceInstall> {
    try {
      const SlackWorkspaceInstall = mongoose.models
        .SlackWorkspaceInstall as Model<ISlackWorkspaceInstall>;

      const filter =
        record.enterpriseId || record.teamId
          ? {
              enterpriseId: record.enterpriseId || null,
              teamId: record.teamId || null,
            }
          : {
              botUserId: record.botUserId,
            };

      return (await SlackWorkspaceInstall.findOneAndUpdate(
        filter,
        { $set: record },
        { new: true, upsert: true, setDefaultsOnInsert: true },
      )) as ISlackWorkspaceInstall;
    } catch (error) {
      logger.error('[upsertSlackWorkspaceInstall] Error upserting Slack workspace install', error);
      throw new Error('Error upserting Slack workspace install');
    }
  }

  async function findSlackWorkspaceInstall(
    filter: FilterQuery<ISlackWorkspaceInstall> = {},
  ): Promise<ISlackWorkspaceInstall | null> {
    try {
      const SlackWorkspaceInstall = mongoose.models
        .SlackWorkspaceInstall as Model<ISlackWorkspaceInstall>;
      return (await SlackWorkspaceInstall.findOne(filter).sort({ updatedAt: -1 }).lean()) as
        | ISlackWorkspaceInstall
        | null;
    } catch (error) {
      logger.error('[findSlackWorkspaceInstall] Error finding Slack workspace install', error);
      throw new Error('Error finding Slack workspace install');
    }
  }

  async function upsertSlackIdentityLink(record: SlackIdentityLinkData): Promise<ISlackIdentityLink> {
    try {
      const SlackIdentityLink = mongoose.models.SlackIdentityLink as Model<ISlackIdentityLink>;

      return (await SlackIdentityLink.findOneAndUpdate(
        {
          user: record.user,
          slackUserId: record.slackUserId,
          teamId: record.teamId || null,
          enterpriseId: record.enterpriseId || null,
        },
        { $set: record },
        { new: true, upsert: true, setDefaultsOnInsert: true },
      )) as ISlackIdentityLink;
    } catch (error) {
      logger.error('[upsertSlackIdentityLink] Error upserting Slack identity link', error);
      throw new Error('Error upserting Slack identity link');
    }
  }

  async function findSlackIdentityLink(
    filter: FilterQuery<ISlackIdentityLink> = {},
  ): Promise<ISlackIdentityLink | null> {
    try {
      const SlackIdentityLink = mongoose.models.SlackIdentityLink as Model<ISlackIdentityLink>;
      return (await SlackIdentityLink.findOne(filter).sort({ updatedAt: -1 }).lean()) as
        | ISlackIdentityLink
        | null;
    } catch (error) {
      logger.error('[findSlackIdentityLink] Error finding Slack identity link', error);
      throw new Error('Error finding Slack identity link');
    }
  }

  async function upsertSlackArchiveConversation(
    record: SlackArchiveConversationData,
  ): Promise<ISlackArchiveConversation> {
    try {
      const SlackArchiveConversation = mongoose.models
        .SlackArchiveConversation as Model<ISlackArchiveConversation>;

      return (await SlackArchiveConversation.findOneAndUpdate(
        { user: record.user, slackConversationId: record.slackConversationId },
        { $set: record },
        { new: true, upsert: true, setDefaultsOnInsert: true },
      )) as ISlackArchiveConversation;
    } catch (error) {
      logger.error('[upsertSlackArchiveConversation] Error upserting Slack archive conversation', error);
      throw new Error('Error upserting Slack archive conversation');
    }
  }

  async function bulkUpsertSlackArchiveMessages(records: SlackArchiveMessageData[]): Promise<number> {
    try {
      if (!Array.isArray(records) || records.length === 0) {
        return 0;
      }

      const SlackArchiveMessage = mongoose.models.SlackArchiveMessage as Model<ISlackArchiveMessage>;
      const operations = records.map((record) => ({
        updateOne: {
          filter: {
            user: record.user,
            slackConversationId: record.slackConversationId,
            slackMessageTs: record.slackMessageTs,
          },
          update: { $set: record },
          upsert: true,
        },
      }));

      const result = await SlackArchiveMessage.bulkWrite(operations, { ordered: false });
      return result ? records.length : 0;
    } catch (error) {
      logger.error('[bulkUpsertSlackArchiveMessages] Error upserting Slack archive messages', error);
      throw new Error('Error upserting Slack archive messages');
    }
  }

  async function createSlackArchiveSyncJob(record: SlackArchiveSyncJobData): Promise<ISlackArchiveSyncJob> {
    try {
      const SlackArchiveSyncJob = mongoose.models.SlackArchiveSyncJob as Model<ISlackArchiveSyncJob>;
      return (await SlackArchiveSyncJob.create(record)) as ISlackArchiveSyncJob;
    } catch (error) {
      logger.error('[createSlackArchiveSyncJob] Error creating Slack archive sync job', error);
      throw new Error('Error creating Slack archive sync job');
    }
  }

  async function updateSlackArchiveSyncJob(
    id: string,
    updates: Partial<SlackArchiveSyncJobData>,
  ): Promise<ISlackArchiveSyncJob | null> {
    try {
      const SlackArchiveSyncJob = mongoose.models.SlackArchiveSyncJob as Model<ISlackArchiveSyncJob>;
      return (await SlackArchiveSyncJob.findByIdAndUpdate(id, { $set: updates }, { new: true })) as
        | ISlackArchiveSyncJob
        | null;
    } catch (error) {
      logger.error('[updateSlackArchiveSyncJob] Error updating Slack archive sync job', error);
      throw new Error('Error updating Slack archive sync job');
    }
  }

  async function findSlackArchiveConversations(
    filter: FilterQuery<ISlackArchiveConversation> = {},
    options: SlackArchiveQueryOptions = {},
  ): Promise<ISlackArchiveConversation[]> {
    try {
      const SlackArchiveConversation = mongoose.models
        .SlackArchiveConversation as Model<ISlackArchiveConversation>;
      const { limit = 50, offset = 0, sort = { lastMessageAt: -1, updatedAt: -1 } } = options;
      return (await SlackArchiveConversation.find(filter)
        .sort(sort)
        .skip(offset)
        .limit(limit)
        .lean()) as ISlackArchiveConversation[];
    } catch (error) {
      logger.error('[findSlackArchiveConversations] Error finding Slack archive conversations', error);
      throw new Error('Error finding Slack archive conversations');
    }
  }

  async function findSlackArchiveMessages(
    filter: FilterQuery<ISlackArchiveMessage> = {},
    options: SlackArchiveQueryOptions = {},
  ): Promise<ISlackArchiveMessage[]> {
    try {
      const SlackArchiveMessage = mongoose.models.SlackArchiveMessage as Model<ISlackArchiveMessage>;
      const { limit = 50, offset = 0, sort = { sentAt: -1, createdAt: -1 } } = options;
      return (await SlackArchiveMessage.find(filter)
        .sort(sort)
        .skip(offset)
        .limit(limit)
        .lean()) as ISlackArchiveMessage[];
    } catch (error) {
      logger.error('[findSlackArchiveMessages] Error finding Slack archive messages', error);
      throw new Error('Error finding Slack archive messages');
    }
  }

  async function findLatestSlackArchiveSyncJob(
    filter: FilterQuery<ISlackArchiveSyncJob> = {},
  ): Promise<ISlackArchiveSyncJob | null> {
    try {
      const SlackArchiveSyncJob = mongoose.models.SlackArchiveSyncJob as Model<ISlackArchiveSyncJob>;
      return (await SlackArchiveSyncJob.findOne(filter).sort({ createdAt: -1 }).lean()) as
        | ISlackArchiveSyncJob
        | null;
    } catch (error) {
      logger.error('[findLatestSlackArchiveSyncJob] Error finding Slack archive sync job', error);
      throw new Error('Error finding Slack archive sync job');
    }
  }

  async function countSlackArchiveConversations(
    filter: FilterQuery<ISlackArchiveConversation> = {},
  ): Promise<number> {
    try {
      const SlackArchiveConversation = mongoose.models
        .SlackArchiveConversation as Model<ISlackArchiveConversation>;
      return await SlackArchiveConversation.countDocuments(filter);
    } catch (error) {
      logger.error('[countSlackArchiveConversations] Error counting Slack archive conversations', error);
      throw new Error('Error counting Slack archive conversations');
    }
  }

  async function countSlackArchiveMessages(
    filter: FilterQuery<ISlackArchiveMessage> = {},
  ): Promise<number> {
    try {
      const SlackArchiveMessage = mongoose.models.SlackArchiveMessage as Model<ISlackArchiveMessage>;
      return await SlackArchiveMessage.countDocuments(filter);
    } catch (error) {
      logger.error('[countSlackArchiveMessages] Error counting Slack archive messages', error);
      throw new Error('Error counting Slack archive messages');
    }
  }

  async function deleteSlackArchiveConversations(
    filter: FilterQuery<ISlackArchiveConversation> = {},
  ): Promise<number> {
    try {
      const SlackArchiveConversation = mongoose.models
        .SlackArchiveConversation as Model<ISlackArchiveConversation>;
      const result = await SlackArchiveConversation.deleteMany(filter);
      return result?.deletedCount || 0;
    } catch (error) {
      logger.error('[deleteSlackArchiveConversations] Error deleting Slack archive conversations', error);
      throw new Error('Error deleting Slack archive conversations');
    }
  }

  async function deleteSlackArchiveMessages(
    filter: FilterQuery<ISlackArchiveMessage> = {},
  ): Promise<number> {
    try {
      const SlackArchiveMessage = mongoose.models.SlackArchiveMessage as Model<ISlackArchiveMessage>;
      const result = await SlackArchiveMessage.deleteMany(filter);
      return result?.deletedCount || 0;
    } catch (error) {
      logger.error('[deleteSlackArchiveMessages] Error deleting Slack archive messages', error);
      throw new Error('Error deleting Slack archive messages');
    }
  }

  async function deleteSlackArchiveSyncJobs(
    filter: FilterQuery<ISlackArchiveSyncJob> = {},
  ): Promise<number> {
    try {
      const SlackArchiveSyncJob = mongoose.models.SlackArchiveSyncJob as Model<ISlackArchiveSyncJob>;
      const result = await SlackArchiveSyncJob.deleteMany(filter);
      return result?.deletedCount || 0;
    } catch (error) {
      logger.error('[deleteSlackArchiveSyncJobs] Error deleting Slack archive sync jobs', error);
      throw new Error('Error deleting Slack archive sync jobs');
    }
  }

  async function deleteSlackArchiveSyncLeases(
    filter: FilterQuery<ISlackArchiveSyncLease> = {},
  ): Promise<number> {
    try {
      const SlackArchiveSyncLease = mongoose.models.SlackArchiveSyncLease as Model<ISlackArchiveSyncLease>;
      const result = await SlackArchiveSyncLease.deleteMany(filter);
      return result?.deletedCount || 0;
    } catch (error) {
      logger.error('[deleteSlackArchiveSyncLeases] Error deleting Slack archive sync leases', error);
      throw new Error('Error deleting Slack archive sync leases');
    }
  }

  async function updateSlackArchiveConversation(
    id: string,
    updates: Partial<SlackArchiveConversationData>,
  ): Promise<ISlackArchiveConversation | null> {
    try {
      const SlackArchiveConversation = mongoose.models
        .SlackArchiveConversation as Model<ISlackArchiveConversation>;
      return (await SlackArchiveConversation.findByIdAndUpdate(id, { $set: updates }, { new: true })) as
        | ISlackArchiveConversation
        | null;
    } catch (error) {
      logger.error('[updateSlackArchiveConversation] Error updating Slack archive conversation', error);
      throw new Error('Error updating Slack archive conversation');
    }
  }

  async function acquireSlackArchiveSyncLease(
    record: SlackArchiveSyncLeaseData,
  ): Promise<ISlackArchiveSyncLease | null> {
    try {
      const SlackArchiveSyncLease = mongoose.models.SlackArchiveSyncLease as Model<ISlackArchiveSyncLease>;
      return (await SlackArchiveSyncLease.findOneAndUpdate(
        {
          leaseKey: record.leaseKey,
          $or: [
            { leaseExpiresAt: { $lte: new Date() } },
            { ownerToken: record.ownerToken },
          ],
        },
        {
          $set: record,
          $setOnInsert: {
            leaseKey: record.leaseKey,
          },
        },
        {
          new: true,
          upsert: true,
        },
      )) as ISlackArchiveSyncLease | null;
    } catch (error) {
      logger.error('[acquireSlackArchiveSyncLease] Error acquiring Slack archive sync lease', error);
      throw new Error('Error acquiring Slack archive sync lease');
    }
  }

  async function refreshSlackArchiveSyncLease(
    leaseKey: string,
    ownerToken: string,
    leaseExpiresAt: Date,
  ): Promise<ISlackArchiveSyncLease | null> {
    try {
      const SlackArchiveSyncLease = mongoose.models.SlackArchiveSyncLease as Model<ISlackArchiveSyncLease>;
      return (await SlackArchiveSyncLease.findOneAndUpdate(
        { leaseKey, ownerToken },
        { $set: { leaseExpiresAt, lastHeartbeatAt: new Date() } },
        { new: true },
      )) as ISlackArchiveSyncLease | null;
    } catch (error) {
      logger.error('[refreshSlackArchiveSyncLease] Error refreshing Slack archive sync lease', error);
      throw new Error('Error refreshing Slack archive sync lease');
    }
  }

  async function releaseSlackArchiveSyncLease(leaseKey: string, ownerToken: string): Promise<boolean> {
    try {
      const SlackArchiveSyncLease = mongoose.models.SlackArchiveSyncLease as Model<ISlackArchiveSyncLease>;
      const result = await SlackArchiveSyncLease.deleteOne({ leaseKey, ownerToken });
      return Boolean(result?.deletedCount);
    } catch (error) {
      logger.error('[releaseSlackArchiveSyncLease] Error releasing Slack archive sync lease', error);
      throw new Error('Error releasing Slack archive sync lease');
    }
  }

  async function countActiveSlackArchiveSyncLeases(
    filter: FilterQuery<ISlackArchiveSyncLease> = {},
  ): Promise<number> {
    try {
      const SlackArchiveSyncLease = mongoose.models.SlackArchiveSyncLease as Model<ISlackArchiveSyncLease>;
      return await SlackArchiveSyncLease.countDocuments(filter);
    } catch (error) {
      logger.error('[countActiveSlackArchiveSyncLeases] Error counting Slack archive sync leases', error);
      throw new Error('Error counting Slack archive sync leases');
    }
  }

  return {
    upsertSlackWorkspaceInstall,
    findSlackWorkspaceInstall,
    upsertSlackIdentityLink,
    findSlackIdentityLink,
    upsertSlackArchiveConversation,
    bulkUpsertSlackArchiveMessages,
    createSlackArchiveSyncJob,
    updateSlackArchiveSyncJob,
    findSlackArchiveConversations,
    findSlackArchiveMessages,
    findLatestSlackArchiveSyncJob,
    countSlackArchiveConversations,
    countSlackArchiveMessages,
    deleteSlackArchiveConversations,
    deleteSlackArchiveMessages,
    deleteSlackArchiveSyncJobs,
    deleteSlackArchiveSyncLeases,
    updateSlackArchiveConversation,
    acquireSlackArchiveSyncLease,
    refreshSlackArchiveSyncLease,
    releaseSlackArchiveSyncLease,
    countActiveSlackArchiveSyncLeases,
  };
}

export type SlackArchiveMethods = ReturnType<typeof createSlackArchiveMethods>;
