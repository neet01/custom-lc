import mongoose, { Schema, Document, Types } from 'mongoose';

export interface ITeamsArchiveBackfillState extends Document {
  user: Types.ObjectId;
  status: 'idle' | 'discovering' | 'syncing' | 'complete' | 'failed';
  nextChatPageLink?: string;
  discoveryComplete?: boolean;
  discoveredChatCount?: number;
  completedChatCount?: number;
  pendingChatCount?: number;
  runningChatCount?: number;
  failedChatCount?: number;
  totalMessageCount?: number;
  lastSyncJobId?: string;
  lastProjectionJobId?: string;
  lastDiscoveredAt?: Date;
  lastCompletedAt?: Date;
  lastHeartbeatAt?: Date;
  errorMessage?: string;
  tenantId?: string;
  createdAt?: Date;
  updatedAt?: Date;
}

const teamsArchiveBackfillStateSchema = new Schema<ITeamsArchiveBackfillState>(
  {
    user: {
      type: mongoose.Schema.Types.ObjectId,
      ref: 'User',
      required: true,
      index: true,
      unique: true,
    },
    status: {
      type: String,
      enum: ['idle', 'discovering', 'syncing', 'complete', 'failed'],
      required: true,
      default: 'idle',
      index: true,
    },
    nextChatPageLink: {
      type: String,
      maxlength: 4096,
    },
    discoveryComplete: {
      type: Boolean,
      default: false,
      index: true,
    },
    discoveredChatCount: {
      type: Number,
      default: 0,
    },
    completedChatCount: {
      type: Number,
      default: 0,
    },
    pendingChatCount: {
      type: Number,
      default: 0,
    },
    runningChatCount: {
      type: Number,
      default: 0,
    },
    failedChatCount: {
      type: Number,
      default: 0,
    },
    totalMessageCount: {
      type: Number,
      default: 0,
    },
    lastSyncJobId: {
      type: String,
      maxlength: 128,
      index: true,
    },
    lastProjectionJobId: {
      type: String,
      maxlength: 128,
      index: true,
    },
    lastDiscoveredAt: {
      type: Date,
      index: true,
    },
    lastCompletedAt: {
      type: Date,
      index: true,
    },
    lastHeartbeatAt: {
      type: Date,
      index: true,
    },
    errorMessage: {
      type: String,
      maxlength: 2000,
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

teamsArchiveBackfillStateSchema.index({ tenantId: 1, status: 1, updatedAt: -1 });

export default teamsArchiveBackfillStateSchema;
