import mongoose, { Schema, Document, Types } from 'mongoose';

export type TeamsArchiveSyncStatus = 'running' | 'success' | 'failure';

export interface ITeamsArchiveSyncJob extends Document {
  user: Types.ObjectId;
  status: TeamsArchiveSyncStatus;
  mode?: string;
  conversationCount?: number;
  messageCount?: number;
  errorMessage?: string;
  startedAt?: Date;
  completedAt?: Date;
  tenantId?: string;
  createdAt?: Date;
  updatedAt?: Date;
}

const teamsArchiveSyncJobSchema = new Schema<ITeamsArchiveSyncJob>(
  {
    user: {
      type: mongoose.Schema.Types.ObjectId,
      ref: 'User',
      required: true,
      index: true,
    },
    status: {
      type: String,
      enum: ['running', 'success', 'failure'],
      required: true,
      index: true,
    },
    mode: {
      type: String,
      maxlength: 64,
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

teamsArchiveSyncJobSchema.index({ user: 1, createdAt: -1 });
teamsArchiveSyncJobSchema.index({ tenantId: 1, createdAt: -1 });

export default teamsArchiveSyncJobSchema;
