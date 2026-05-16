import mongoose, { Schema, Document, Types } from 'mongoose';
import type { EnterpriseMemoryVisibilityScope } from './enterpriseMemoryEntity';

export type EnterpriseMemoryJobStatus = 'pending' | 'running' | 'success' | 'failure';

export interface IEnterpriseMemoryJob extends Document {
  user?: Types.ObjectId;
  tenantId?: string;
  visibilityScope: EnterpriseMemoryVisibilityScope;
  source: string;
  jobType: string;
  status: EnterpriseMemoryJobStatus;
  sourceRecordType?: string;
  sourceRecordId?: string;
  checkpoint?: Record<string, unknown>;
  stats?: Record<string, unknown>;
  errorMessage?: string;
  startedAt?: Date;
  completedAt?: Date;
  lastHeartbeatAt?: Date;
  createdAt?: Date;
  updatedAt?: Date;
}

const enterpriseMemoryJobSchema = new Schema<IEnterpriseMemoryJob>(
  {
    user: {
      type: mongoose.Schema.Types.ObjectId,
      ref: 'User',
      index: true,
    },
    tenantId: {
      type: String,
      index: true,
    },
    visibilityScope: {
      type: String,
      enum: ['user', 'tenant'],
      required: true,
      default: 'user',
      index: true,
    },
    source: {
      type: String,
      required: true,
      maxlength: 64,
      index: true,
    },
    jobType: {
      type: String,
      required: true,
      maxlength: 64,
      index: true,
    },
    status: {
      type: String,
      enum: ['pending', 'running', 'success', 'failure'],
      required: true,
      default: 'pending',
      index: true,
    },
    sourceRecordType: {
      type: String,
      maxlength: 64,
      index: true,
    },
    sourceRecordId: {
      type: String,
      maxlength: 256,
      index: true,
    },
    checkpoint: {
      type: Schema.Types.Mixed,
    },
    stats: {
      type: Schema.Types.Mixed,
    },
    errorMessage: {
      type: String,
      maxlength: 4096,
    },
    startedAt: {
      type: Date,
      index: true,
    },
    completedAt: {
      type: Date,
      index: true,
    },
    lastHeartbeatAt: {
      type: Date,
      index: true,
    },
  },
  {
    timestamps: true,
  },
);

enterpriseMemoryJobSchema.index({ source: 1, jobType: 1, status: 1, createdAt: -1 });
enterpriseMemoryJobSchema.index({ tenantId: 1, source: 1, createdAt: -1 });

export default enterpriseMemoryJobSchema;
