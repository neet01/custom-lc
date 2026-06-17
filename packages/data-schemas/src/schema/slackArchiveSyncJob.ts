import mongoose, { Schema, Document, Types } from 'mongoose';

export type SlackArchiveSyncStatus = 'running' | 'success' | 'partial' | 'failure' | 'cancelled';

export interface ISlackArchiveSyncJob extends Document {
  user: Types.ObjectId;
  status: SlackArchiveSyncStatus;
  mode?: string;
  phase?: string;
  checkpoint?: Record<string, unknown>;
  stats?: Record<string, unknown>;
  requestedConversationLimit?: number;
  requestedMessagesPerConversation?: number;
  discoveredConversationCount?: number;
  processedConversationCount?: number;
  skippedConversationCount?: number;
  conversationCount?: number;
  messageCount?: number;
  errorMessage?: string;
  startedAt?: Date;
  completedAt?: Date;
  tenantId?: string;
  createdAt?: Date;
  updatedAt?: Date;
}

const slackArchiveSyncJobSchema = new Schema<ISlackArchiveSyncJob>(
  {
    user: {
      type: mongoose.Schema.Types.ObjectId,
      ref: 'User',
      required: true,
      index: true,
    },
    status: {
      type: String,
      enum: ['running', 'success', 'partial', 'failure', 'cancelled'],
      required: true,
      index: true,
    },
    mode: {
      type: String,
      maxlength: 64,
    },
    phase: {
      type: String,
      maxlength: 64,
      index: true,
    },
    checkpoint: {
      type: Schema.Types.Mixed,
    },
    stats: {
      type: Schema.Types.Mixed,
    },
    requestedConversationLimit: {
      type: Number,
      default: 0,
    },
    requestedMessagesPerConversation: {
      type: Number,
      default: 0,
    },
    discoveredConversationCount: {
      type: Number,
      default: 0,
    },
    processedConversationCount: {
      type: Number,
      default: 0,
    },
    skippedConversationCount: {
      type: Number,
      default: 0,
    },
    conversationCount: {
      type: Number,
      default: 0,
    },
    messageCount: {
      type: Number,
      default: 0,
    },
    errorMessage: {
      type: String,
      maxlength: 1000,
    },
    startedAt: {
      type: Date,
      index: true,
    },
    completedAt: {
      type: Date,
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

slackArchiveSyncJobSchema.index({ user: 1, createdAt: -1 });
slackArchiveSyncJobSchema.index({ tenantId: 1, createdAt: -1 });

export default slackArchiveSyncJobSchema;
