import type { FilterQuery, Model } from 'mongoose';
import logger from '~/config/winston';
import type { IEnterpriseMemoryEntity } from '~/schema/enterpriseMemoryEntity';
import type { IEnterpriseMemoryRelationship } from '~/schema/enterpriseMemoryRelationship';
import type { IEnterpriseMemoryChunk } from '~/schema/enterpriseMemoryChunk';
import type { IEnterpriseMemoryJob } from '~/schema/enterpriseMemoryJob';
import type {
  EnterpriseMemoryChunkData,
  EnterpriseMemoryEntityData,
  EnterpriseMemoryJobData,
  EnterpriseMemoryRelationshipData,
} from '~/types/enterpriseMemory';

export interface EnterpriseMemoryQueryOptions {
  limit?: number;
  offset?: number;
  sort?: Record<string, 1 | -1>;
}

export function createEnterpriseMemoryMethods(mongoose: typeof import('mongoose')) {
  async function upsertEnterpriseMemoryEntity(
    record: EnterpriseMemoryEntityData,
  ): Promise<IEnterpriseMemoryEntity> {
    try {
      const EnterpriseMemoryEntity = mongoose.models
        .EnterpriseMemoryEntity as Model<IEnterpriseMemoryEntity>;

      return (await EnterpriseMemoryEntity.findOneAndUpdate(
        {
          visibilityScope: record.visibilityScope || 'user',
          user: record.user || null,
          tenantId: record.tenantId || null,
          source: record.source,
          entityType: record.entityType,
          canonicalKey: record.canonicalKey,
        },
        { $set: record },
        { new: true, upsert: true, setDefaultsOnInsert: true },
      )) as IEnterpriseMemoryEntity;
    } catch (error) {
      logger.error('[upsertEnterpriseMemoryEntity] Error upserting enterprise memory entity', error);
      throw new Error('Error upserting enterprise memory entity');
    }
  }

  async function bulkUpsertEnterpriseMemoryRelationships(
    records: EnterpriseMemoryRelationshipData[],
  ): Promise<number> {
    try {
      if (!Array.isArray(records) || records.length === 0) {
        return 0;
      }

      const EnterpriseMemoryRelationship = mongoose.models
        .EnterpriseMemoryRelationship as Model<IEnterpriseMemoryRelationship>;
      const operations = records.map((record) => ({
        updateOne: {
          filter: {
            visibilityScope: record.visibilityScope || 'user',
            user: record.user || null,
            tenantId: record.tenantId || null,
            source: record.source,
            relationshipType: record.relationshipType,
            fromEntityId: record.fromEntityId,
            toEntityId: record.toEntityId,
            sourceRecordType: record.sourceRecordType || '',
            sourceRecordId: record.sourceRecordId || '',
          },
          update: {
            $set: {
              ...record,
              sourceRecordType: record.sourceRecordType || '',
              sourceRecordId: record.sourceRecordId || '',
            },
          },
          upsert: true,
        },
      }));

      const result = await EnterpriseMemoryRelationship.bulkWrite(operations, { ordered: false });
      return result ? records.length : 0;
    } catch (error) {
      logger.error(
        '[bulkUpsertEnterpriseMemoryRelationships] Error upserting enterprise memory relationships',
        error,
      );
      throw new Error('Error upserting enterprise memory relationships');
    }
  }

  async function bulkUpsertEnterpriseMemoryChunks(
    records: EnterpriseMemoryChunkData[],
  ): Promise<number> {
    try {
      if (!Array.isArray(records) || records.length === 0) {
        return 0;
      }

      const EnterpriseMemoryChunk = mongoose.models
        .EnterpriseMemoryChunk as Model<IEnterpriseMemoryChunk>;
      const operations = records.map((record) => ({
        updateOne: {
          filter: {
            visibilityScope: record.visibilityScope || 'user',
            user: record.user || null,
            tenantId: record.tenantId || null,
            source: record.source,
            sourceRecordType: record.sourceRecordType,
            sourceRecordId: record.sourceRecordId,
            chunkType: record.chunkType,
            orderIndex: record.orderIndex || 0,
          },
          update: { $set: { ...record, orderIndex: record.orderIndex || 0 } },
          upsert: true,
        },
      }));

      const result = await EnterpriseMemoryChunk.bulkWrite(operations, { ordered: false });
      return result ? records.length : 0;
    } catch (error) {
      logger.error('[bulkUpsertEnterpriseMemoryChunks] Error upserting enterprise memory chunks', error);
      throw new Error('Error upserting enterprise memory chunks');
    }
  }

  async function createEnterpriseMemoryJob(
    record: EnterpriseMemoryJobData,
  ): Promise<IEnterpriseMemoryJob> {
    try {
      const EnterpriseMemoryJob = mongoose.models.EnterpriseMemoryJob as Model<IEnterpriseMemoryJob>;
      return (await EnterpriseMemoryJob.create(record)) as IEnterpriseMemoryJob;
    } catch (error) {
      logger.error('[createEnterpriseMemoryJob] Error creating enterprise memory job', error);
      throw new Error('Error creating enterprise memory job');
    }
  }

  async function updateEnterpriseMemoryJob(
    id: string,
    updates: Partial<EnterpriseMemoryJobData>,
  ): Promise<IEnterpriseMemoryJob | null> {
    try {
      const EnterpriseMemoryJob = mongoose.models.EnterpriseMemoryJob as Model<IEnterpriseMemoryJob>;
      return (await EnterpriseMemoryJob.findByIdAndUpdate(id, { $set: updates }, { new: true })) as
        | IEnterpriseMemoryJob
        | null;
    } catch (error) {
      logger.error('[updateEnterpriseMemoryJob] Error updating enterprise memory job', error);
      throw new Error('Error updating enterprise memory job');
    }
  }

  async function findEnterpriseMemoryEntities(
    filter: FilterQuery<IEnterpriseMemoryEntity> = {},
    options: EnterpriseMemoryQueryOptions = {},
  ): Promise<IEnterpriseMemoryEntity[]> {
    try {
      const EnterpriseMemoryEntity = mongoose.models
        .EnterpriseMemoryEntity as Model<IEnterpriseMemoryEntity>;
      const { limit = 100, offset = 0, sort = { updatedAt: -1 } } = options;
      return (await EnterpriseMemoryEntity.find(filter)
        .sort(sort)
        .skip(offset)
        .limit(limit)
        .lean()) as IEnterpriseMemoryEntity[];
    } catch (error) {
      logger.error('[findEnterpriseMemoryEntities] Error finding enterprise memory entities', error);
      throw new Error('Error finding enterprise memory entities');
    }
  }

  async function findEnterpriseMemoryChunks(
    filter: FilterQuery<IEnterpriseMemoryChunk> = {},
    options: EnterpriseMemoryQueryOptions = {},
  ): Promise<IEnterpriseMemoryChunk[]> {
    try {
      const EnterpriseMemoryChunk = mongoose.models
        .EnterpriseMemoryChunk as Model<IEnterpriseMemoryChunk>;
      const { limit = 100, offset = 0, sort = { sourceTimestamp: -1, updatedAt: -1 } } = options;
      return (await EnterpriseMemoryChunk.find(filter)
        .sort(sort)
        .skip(offset)
        .limit(limit)
        .lean()) as IEnterpriseMemoryChunk[];
    } catch (error) {
      logger.error('[findEnterpriseMemoryChunks] Error finding enterprise memory chunks', error);
      throw new Error('Error finding enterprise memory chunks');
    }
  }

  async function countEnterpriseMemoryChunks(
    filter: FilterQuery<IEnterpriseMemoryChunk> = {},
  ): Promise<number> {
    try {
      const EnterpriseMemoryChunk = mongoose.models
        .EnterpriseMemoryChunk as Model<IEnterpriseMemoryChunk>;
      return await EnterpriseMemoryChunk.countDocuments(filter);
    } catch (error) {
      logger.error('[countEnterpriseMemoryChunks] Error counting enterprise memory chunks', error);
      throw new Error('Error counting enterprise memory chunks');
    }
  }

  async function countDistinctEnterpriseMemoryChunkField(
    field: string,
    filter: FilterQuery<IEnterpriseMemoryChunk> = {},
  ): Promise<number> {
    try {
      const EnterpriseMemoryChunk = mongoose.models
        .EnterpriseMemoryChunk as Model<IEnterpriseMemoryChunk>;
      const values = await EnterpriseMemoryChunk.distinct(field, filter);
      return values.filter((value) => value !== undefined && value !== null && value !== '').length;
    } catch (error) {
      logger.error(
        '[countDistinctEnterpriseMemoryChunkField] Error counting distinct enterprise memory chunk field values',
        error,
      );
      throw new Error('Error counting distinct enterprise memory chunk field values');
    }
  }

  async function findLatestEnterpriseMemoryJob(
    filter: FilterQuery<IEnterpriseMemoryJob> = {},
  ): Promise<IEnterpriseMemoryJob | null> {
    try {
      const EnterpriseMemoryJob = mongoose.models.EnterpriseMemoryJob as Model<IEnterpriseMemoryJob>;
      return (await EnterpriseMemoryJob.findOne(filter).sort({ createdAt: -1 }).lean()) as
        | IEnterpriseMemoryJob
        | null;
    } catch (error) {
      logger.error('[findLatestEnterpriseMemoryJob] Error finding enterprise memory job', error);
      throw new Error('Error finding enterprise memory job');
    }
  }

  async function deleteEnterpriseMemoryEntities(
    filter: FilterQuery<IEnterpriseMemoryEntity> = {},
  ): Promise<number> {
    try {
      const EnterpriseMemoryEntity = mongoose.models
        .EnterpriseMemoryEntity as Model<IEnterpriseMemoryEntity>;
      const result = await EnterpriseMemoryEntity.deleteMany(filter);
      return result?.deletedCount || 0;
    } catch (error) {
      logger.error('[deleteEnterpriseMemoryEntities] Error deleting enterprise memory entities', error);
      throw new Error('Error deleting enterprise memory entities');
    }
  }

  async function deleteEnterpriseMemoryRelationships(
    filter: FilterQuery<IEnterpriseMemoryRelationship> = {},
  ): Promise<number> {
    try {
      const EnterpriseMemoryRelationship = mongoose.models
        .EnterpriseMemoryRelationship as Model<IEnterpriseMemoryRelationship>;
      const result = await EnterpriseMemoryRelationship.deleteMany(filter);
      return result?.deletedCount || 0;
    } catch (error) {
      logger.error(
        '[deleteEnterpriseMemoryRelationships] Error deleting enterprise memory relationships',
        error,
      );
      throw new Error('Error deleting enterprise memory relationships');
    }
  }

  async function deleteEnterpriseMemoryChunks(
    filter: FilterQuery<IEnterpriseMemoryChunk> = {},
  ): Promise<number> {
    try {
      const EnterpriseMemoryChunk = mongoose.models
        .EnterpriseMemoryChunk as Model<IEnterpriseMemoryChunk>;
      const result = await EnterpriseMemoryChunk.deleteMany(filter);
      return result?.deletedCount || 0;
    } catch (error) {
      logger.error('[deleteEnterpriseMemoryChunks] Error deleting enterprise memory chunks', error);
      throw new Error('Error deleting enterprise memory chunks');
    }
  }

  async function deleteEnterpriseMemoryJobs(
    filter: FilterQuery<IEnterpriseMemoryJob> = {},
  ): Promise<number> {
    try {
      const EnterpriseMemoryJob = mongoose.models.EnterpriseMemoryJob as Model<IEnterpriseMemoryJob>;
      const result = await EnterpriseMemoryJob.deleteMany(filter);
      return result?.deletedCount || 0;
    } catch (error) {
      logger.error('[deleteEnterpriseMemoryJobs] Error deleting enterprise memory jobs', error);
      throw new Error('Error deleting enterprise memory jobs');
    }
  }

  return {
    upsertEnterpriseMemoryEntity,
    bulkUpsertEnterpriseMemoryRelationships,
    bulkUpsertEnterpriseMemoryChunks,
    createEnterpriseMemoryJob,
    updateEnterpriseMemoryJob,
    findEnterpriseMemoryEntities,
    findEnterpriseMemoryChunks,
    countEnterpriseMemoryChunks,
    countDistinctEnterpriseMemoryChunkField,
    findLatestEnterpriseMemoryJob,
    deleteEnterpriseMemoryEntities,
    deleteEnterpriseMemoryRelationships,
    deleteEnterpriseMemoryChunks,
    deleteEnterpriseMemoryJobs,
  };
}

export type EnterpriseMemoryMethods = ReturnType<typeof createEnterpriseMemoryMethods>;
