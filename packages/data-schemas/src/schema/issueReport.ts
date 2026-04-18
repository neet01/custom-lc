import mongoose, { Schema, Document, Types } from 'mongoose';

export type IssueReportCategory =
  | 'bad_response'
  | 'faulty_mcp_tool'
  | 'bad_file_transformation'
  | 'timeout_or_error'
  | 'auth_or_permissions'
  | 'other';

export type IssueReportStatus = 'open' | 'triaged' | 'resolved';

export interface IIssueReport extends Document {
  user: Types.ObjectId;
  conversationId: string;
  messageId: string;
  category: IssueReportCategory;
  status: IssueReportStatus;
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
  tenantId?: string;
}

const issueReportSchema = new Schema<IIssueReport>(
  {
    user: {
      type: mongoose.Schema.Types.ObjectId,
      ref: 'User',
      index: true,
      required: true,
    },
    conversationId: {
      type: String,
      required: true,
      index: true,
    },
    messageId: {
      type: String,
      required: true,
      index: true,
    },
    category: {
      type: String,
      enum: [
        'bad_response',
        'faulty_mcp_tool',
        'bad_file_transformation',
        'timeout_or_error',
        'auth_or_permissions',
        'other',
      ],
      required: true,
      index: true,
    },
    status: {
      type: String,
      enum: ['open', 'triaged', 'resolved'],
      default: 'open',
      required: true,
      index: true,
    },
    description: {
      type: String,
      maxlength: 2000,
    },
    model: {
      type: String,
      index: true,
    },
    endpoint: {
      type: String,
      index: true,
    },
    messagePreview: {
      type: String,
      maxlength: 500,
    },
    error: {
      type: Boolean,
      default: false,
      index: true,
    },
    fileIds: {
      type: [String],
      default: undefined,
    },
    toolName: {
      type: String,
      index: true,
    },
    mcpServer: {
      type: String,
      index: true,
    },
    tenantId: {
      type: String,
      index: true,
    },
  },
  {
    timestamps: true,
  },
);

issueReportSchema.index({ user: 1, createdAt: -1 });
issueReportSchema.index({ status: 1, createdAt: -1 });
issueReportSchema.index({ category: 1, createdAt: -1 });
issueReportSchema.index({ tenantId: 1, createdAt: -1 });

export default issueReportSchema;
