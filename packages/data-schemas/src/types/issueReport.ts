import type { IssueReportCategory, IssueReportStatus } from '~/schema/issueReport';

export type { IssueReportCategory, IssueReportStatus };

export interface IssueReportData {
  user: string;
  conversationId: string;
  messageId: string;
  category: IssueReportCategory;
  status?: IssueReportStatus;
  description?: string;
  model?: string;
  endpoint?: string;
  messagePreview?: string;
  error?: boolean;
  fileIds?: string[];
  toolName?: string;
  mcpServer?: string;
}
