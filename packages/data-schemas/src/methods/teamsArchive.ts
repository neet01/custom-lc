import type { FilterQuery, Model } from 'mongoose';
import logger from '~/config/winston';
import type { ITeamsArchiveConversation } from '~/schema/teamsArchiveConversation';
import type { ITeamsArchiveMessage } from '~/schema/teamsArchiveMessage';
import type { ITeamsArchiveSyncJob } from '~/schema/teamsArchiveSyncJob';
import type {
  TeamsArchiveConversationData,
  TeamsArchiveMessageData,
  TeamsArchiveSyncJobData,
} from '~/types/teamsArchive';

export interface TeamsArchiveQueryOptions {
  limit?: number;
  offset?: number;
  sort?: Record<string, 1 | -1>;
}

export function createTeamsArchiveMethods(mongoose: typeof import('mongoose')) {
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

  return {
    upsertTeamsArchiveConversation,
    bulkUpsertTeamsArchiveMessages,
    createTeamsArchiveSyncJob,
    updateTeamsArchiveSyncJob,
    findTeamsArchiveConversations,
    findTeamsArchiveMessages,
    findLatestTeamsArchiveSyncJob,
    countTeamsArchiveConversations,
    countTeamsArchiveMessages,
  };
}

export type TeamsArchiveMethods = ReturnType<typeof createTeamsArchiveMethods>;
