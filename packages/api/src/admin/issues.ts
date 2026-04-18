import { logger, isValidObjectIdString } from '@librechat/data-schemas';
import type { FilterQuery } from 'mongoose';
import type { Response } from 'express';
import type { ServerRequest } from '~/types/http';
import { parsePagination } from './pagination';

type IssueReportRecord = {
  _id?: { toString(): string };
  user?: { toString(): string };
  conversationId: string;
  messageId: string;
  category: string;
  status: string;
  description?: string;
  model?: string;
  endpoint?: string;
  messagePreview?: string;
  error?: boolean;
  fileIds?: string[];
  toolName?: string;
  mcpServer?: string;
  createdAt?: Date;
  updatedAt?: Date;
};

type AdminIssueRecord = {
  id: string;
  userId: string;
  reporterName: string;
  reporterEmail: string;
  reporterAvatar: string;
  reporterRole: string;
  conversationId: string;
  messageId: string;
  category: string;
  status: string;
  description?: string;
  model?: string;
  endpoint?: string;
  messagePreview?: string;
  error?: boolean;
  fileIds?: string[];
  toolName?: string;
  mcpServer?: string;
  createdAt?: string;
  updatedAt?: string;
};

type AdminUserRecord = {
  _id?: { toString(): string };
  name?: string;
  username?: string;
  email?: string;
  avatar?: string;
  role?: string;
  provider?: string;
};

export interface AdminIssuesDeps {
  findIssueReports: (
    filter?: FilterQuery<IssueReportRecord>,
    options?: {
      limit?: number;
      offset?: number;
      sort?: Record<string, 1 | -1>;
    },
  ) => Promise<IssueReportRecord[]>;
  countIssueReports: (filter?: FilterQuery<IssueReportRecord>) => Promise<number>;
  findUsers: (
    searchCriteria: FilterQuery<AdminUserRecord>,
    fieldsToSelect?: string | string[] | null,
    options?: { limit?: number; offset?: number; sort?: Record<string, 1 | -1> },
  ) => Promise<AdminUserRecord[]>;
}

const USER_FIELDS = '_id name username email avatar role provider';

function buildIssueFilter(query: ServerRequest['query']) {
  const filter: FilterQuery<IssueReportRecord> = {};

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

  const category = typeof query.category === 'string' ? query.category : undefined;
  if (category) {
    filter.category = category;
  }

  const status = typeof query.status === 'string' ? query.status : undefined;
  if (status) {
    filter.status = status;
  }

  return { filter };
}

function mapIssue(
  issue: IssueReportRecord,
  usersById: Map<
    string,
    {
      name: string;
      username: string;
      email: string;
      avatar: string;
      role: string;
    }
  >,
): AdminIssueRecord {
  const user = usersById.get(issue.user?.toString() ?? '');

  return {
    id: issue._id?.toString() ?? '',
    userId: issue.user?.toString() ?? '',
    reporterName: user?.name || user?.username || user?.email || '',
    reporterEmail: user?.email || '',
    reporterAvatar: user?.avatar || '',
    reporterRole: user?.role || 'USER',
    conversationId: issue.conversationId,
    messageId: issue.messageId,
    category: issue.category,
    status: issue.status,
    description: issue.description,
    model: issue.model,
    endpoint: issue.endpoint,
    messagePreview: issue.messagePreview,
    error: issue.error,
    fileIds: issue.fileIds,
    toolName: issue.toolName,
    mcpServer: issue.mcpServer,
    createdAt: issue.createdAt?.toISOString(),
    updatedAt: issue.updatedAt?.toISOString(),
  };
}

export function createAdminIssuesHandlers(deps: AdminIssuesDeps) {
  async function listIssues(req: ServerRequest, res: Response) {
    try {
      const { limit, offset } = parsePagination(req.query);
      const { filter, error } = buildIssueFilter(req.query);
      if (!filter) {
        return res.status(400).json({ error });
      }

      const [issues, total] = await Promise.all([
        deps.findIssueReports(filter, { limit, offset, sort: { createdAt: -1 } }),
        deps.countIssueReports(filter),
      ]);

      const userIds = [...new Set(issues.map((issue) => issue.user?.toString()).filter(Boolean))];
      const users =
        userIds.length > 0
          ? await deps.findUsers({ _id: { $in: userIds } }, USER_FIELDS, {
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
          },
        ]),
      );

      return res.status(200).json({
        issues: issues.map((issue) => mapIssue(issue, usersById)),
        total,
        limit,
        offset,
      });
    } catch (error) {
      logger.error('[adminIssues] listIssues error:', error);
      return res.status(500).json({ error: 'Failed to list issue reports' });
    }
  }

  return { listIssues };
}
