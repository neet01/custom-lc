import type { FilterQuery, Model } from 'mongoose';
import logger from '~/config/winston';
import type { IIssueReport } from '~/schema/issueReport';
import type { IssueReportData } from '~/types';

export interface IssueReportQueryOptions {
  limit?: number;
  offset?: number;
  sort?: Record<string, 1 | -1>;
}

export function createIssueReportMethods(mongoose: typeof import('mongoose')) {
  async function createIssueReport(record: IssueReportData): Promise<IIssueReport> {
    try {
      const IssueReport = mongoose.models.IssueReport as Model<IIssueReport>;
      return (await IssueReport.create(record)) as IIssueReport;
    } catch (error) {
      logger.error('[createIssueReport] Error creating issue report', error);
      throw new Error('Error creating issue report');
    }
  }

  async function findIssueReports(
    filter: FilterQuery<IIssueReport> = {},
    options: IssueReportQueryOptions = {},
  ): Promise<IIssueReport[]> {
    try {
      const IssueReport = mongoose.models.IssueReport as Model<IIssueReport>;
      const { limit = 50, offset = 0, sort = { createdAt: -1 } } = options;

      return (await IssueReport.find(filter)
        .sort(sort)
        .skip(offset)
        .limit(limit)
        .lean()) as IIssueReport[];
    } catch (error) {
      logger.error('[findIssueReports] Error finding issue reports', error);
      throw new Error('Error finding issue reports');
    }
  }

  async function countIssueReports(filter: FilterQuery<IIssueReport> = {}): Promise<number> {
    try {
      const IssueReport = mongoose.models.IssueReport as Model<IIssueReport>;
      return await IssueReport.countDocuments(filter);
    } catch (error) {
      logger.error('[countIssueReports] Error counting issue reports', error);
      throw new Error('Error counting issue reports');
    }
  }

  return {
    createIssueReport,
    findIssueReports,
    countIssueReports,
  };
}

export type IssueReportMethods = ReturnType<typeof createIssueReportMethods>;
