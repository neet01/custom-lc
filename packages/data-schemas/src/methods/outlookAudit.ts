import type { FilterQuery, Model } from 'mongoose';
import logger from '~/config/winston';
import type { IOutlookAudit } from '~/schema/outlookAudit';
import type { OutlookAuditData } from '~/types';

export interface OutlookAuditQueryOptions {
  limit?: number;
  offset?: number;
  sort?: Record<string, 1 | -1>;
}

export function createOutlookAuditMethods(mongoose: typeof import('mongoose')) {
  async function createOutlookAudit(record: OutlookAuditData): Promise<IOutlookAudit> {
    try {
      const OutlookAudit = mongoose.models.OutlookAudit as Model<IOutlookAudit>;
      return (await OutlookAudit.create(record)) as IOutlookAudit;
    } catch (error) {
      logger.error('[createOutlookAudit] Error creating Outlook audit record', error);
      throw new Error('Error creating Outlook audit record');
    }
  }

  async function findOutlookAudits(
    filter: FilterQuery<IOutlookAudit> = {},
    options: OutlookAuditQueryOptions = {},
  ): Promise<IOutlookAudit[]> {
    try {
      const OutlookAudit = mongoose.models.OutlookAudit as Model<IOutlookAudit>;
      const { limit = 50, offset = 0, sort = { createdAt: -1 } } = options;

      return (await OutlookAudit.find(filter)
        .sort(sort)
        .skip(offset)
        .limit(limit)
        .lean()) as IOutlookAudit[];
    } catch (error) {
      logger.error('[findOutlookAudits] Error finding Outlook audit records', error);
      throw new Error('Error finding Outlook audit records');
    }
  }

  async function countOutlookAudits(filter: FilterQuery<IOutlookAudit> = {}): Promise<number> {
    try {
      const OutlookAudit = mongoose.models.OutlookAudit as Model<IOutlookAudit>;
      return await OutlookAudit.countDocuments(filter);
    } catch (error) {
      logger.error('[countOutlookAudits] Error counting Outlook audit records', error);
      throw new Error('Error counting Outlook audit records');
    }
  }

  return {
    createOutlookAudit,
    findOutlookAudits,
    countOutlookAudits,
  };
}

export type OutlookAuditMethods = ReturnType<typeof createOutlookAuditMethods>;
