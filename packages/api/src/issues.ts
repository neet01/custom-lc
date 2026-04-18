import { logger } from '@librechat/data-schemas';
import type { Response } from 'express';
import type { ServerRequest } from '~/types/http';

type IssueReportInput = {
  user: string;
  conversationId: string;
  messageId: string;
  category: string;
  status?: string;
  description?: string;
  model?: string;
  endpoint?: string;
  messagePreview?: string;
  error?: boolean;
  fileIds?: string[];
  toolName?: string;
  mcpServer?: string;
};

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

export interface IssueDeps {
  createIssueReport: (record: IssueReportInput) => Promise<IssueReportRecord>;
}

function toIssueResponse(issue: IssueReportRecord) {
  return {
    id: issue._id?.toString() ?? '',
    userId: issue.user?.toString() ?? '',
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

export function createIssueHandlers(deps: IssueDeps) {
  async function reportIssue(req: ServerRequest, res: Response) {
    const userId = req.user?.id;
    const body = (req.body ?? {}) as Record<string, unknown>;
    const conversationId =
      typeof body.conversationId === 'string' ? body.conversationId : undefined;
    const messageId = typeof body.messageId === 'string' ? body.messageId : undefined;
    const category = typeof body.category === 'string' ? body.category : undefined;
    const description = typeof body.description === 'string' ? body.description : undefined;
    const model = typeof body.model === 'string' ? body.model : undefined;
    const endpoint = typeof body.endpoint === 'string' ? body.endpoint : undefined;
    const messagePreview =
      typeof body.messagePreview === 'string' ? body.messagePreview : undefined;
    const error = typeof body.error === 'boolean' ? body.error : undefined;
    const fileIds = Array.isArray(body.fileIds) ? body.fileIds : undefined;
    const toolName = typeof body.toolName === 'string' ? body.toolName : undefined;
    const mcpServer = typeof body.mcpServer === 'string' ? body.mcpServer : undefined;

    if (!userId) {
      return res.status(401).json({ error: 'Authentication required' });
    }

    if (!conversationId || !messageId || !category) {
      return res
        .status(400)
        .json({ error: 'conversationId, messageId, and category are required' });
    }

    try {
      const issue = await deps.createIssueReport({
        user: userId,
        conversationId,
        messageId,
        category,
        description: description?.trim() || undefined,
        model,
        endpoint,
        messagePreview,
        error,
        fileIds: Array.isArray(fileIds)
          ? fileIds.filter((value): value is string => typeof value === 'string')
          : undefined,
        toolName,
        mcpServer,
      });

      return res.status(201).json({ issue: toIssueResponse(issue) });
    } catch (error) {
      logger.error('[issues] reportIssue error:', error);
      return res.status(500).json({ error: 'Failed to create issue report' });
    }
  }

  return { reportIssue };
}
