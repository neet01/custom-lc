import mongoose, { Schema, Document, Types } from 'mongoose';
import type { EnterpriseMemoryVisibilityScope } from './enterpriseMemoryEntity';

export interface IEnterpriseMemoryChunk extends Document {
  user?: Types.ObjectId;
  tenantId?: string;
  visibilityScope: EnterpriseMemoryVisibilityScope;
  source: string;
  sourceRecordType: string;
  sourceRecordId: string;
  sourceParentRecordId?: string;
  parentEntityId?: Types.ObjectId;
  entityIds?: Types.ObjectId[];
  chunkType: string;
  title?: string;
  text: string;
  summary?: string;
  orderIndex: number;
  sourceTimestamp?: Date;
  metadata?: Record<string, unknown>;
  createdAt?: Date;
  updatedAt?: Date;
}

const enterpriseMemoryChunkSchema = new Schema<IEnterpriseMemoryChunk>(
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
    sourceRecordType: {
      type: String,
      required: true,
      maxlength: 64,
      index: true,
    },
    sourceRecordId: {
      type: String,
      required: true,
      maxlength: 256,
      index: true,
    },
    sourceParentRecordId: {
      type: String,
      maxlength: 256,
      index: true,
    },
    parentEntityId: {
      type: mongoose.Schema.Types.ObjectId,
      ref: 'EnterpriseMemoryEntity',
      index: true,
    },
    entityIds: {
      type: [mongoose.Schema.Types.ObjectId],
      ref: 'EnterpriseMemoryEntity',
      default: undefined,
      index: true,
    },
    chunkType: {
      type: String,
      required: true,
      maxlength: 64,
      index: true,
    },
    title: {
      type: String,
      maxlength: 512,
    },
    text: {
      type: String,
      required: true,
    },
    summary: {
      type: String,
      maxlength: 2048,
    },
    orderIndex: {
      type: Number,
      required: true,
      default: 0,
      index: true,
    },
    sourceTimestamp: {
      type: Date,
      index: true,
    },
    metadata: {
      type: Schema.Types.Mixed,
    },
  },
  {
    timestamps: true,
  },
);

enterpriseMemoryChunkSchema.index(
  {
    visibilityScope: 1,
    user: 1,
    tenantId: 1,
    source: 1,
    sourceRecordType: 1,
    sourceRecordId: 1,
    chunkType: 1,
    orderIndex: 1,
  },
  { unique: true },
);
enterpriseMemoryChunkSchema.index({ parentEntityId: 1, sourceTimestamp: -1 });
enterpriseMemoryChunkSchema.index({ source: 1, sourceRecordType: 1, sourceTimestamp: -1 });
enterpriseMemoryChunkSchema.index({
  title: 'text',
  text: 'text',
  summary: 'text',
});

export default enterpriseMemoryChunkSchema;
