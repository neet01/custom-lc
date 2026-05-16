import mongoose, { Schema, Document, Types } from 'mongoose';

export type EnterpriseMemoryVisibilityScope = 'user' | 'tenant';

export interface IEnterpriseMemoryEntity extends Document {
  user?: Types.ObjectId;
  tenantId?: string;
  visibilityScope: EnterpriseMemoryVisibilityScope;
  source: string;
  entityType: string;
  canonicalKey: string;
  displayName: string;
  aliases?: string[];
  summary?: string;
  sourceRecordType?: string;
  sourceRecordId?: string;
  sourceParentRecordId?: string;
  sourceUpdatedAt?: Date;
  attributes?: Record<string, unknown>;
  createdAt?: Date;
  updatedAt?: Date;
}

const enterpriseMemoryEntitySchema = new Schema<IEnterpriseMemoryEntity>(
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
    entityType: {
      type: String,
      required: true,
      maxlength: 64,
      index: true,
    },
    canonicalKey: {
      type: String,
      required: true,
      maxlength: 256,
      index: true,
    },
    displayName: {
      type: String,
      required: true,
      maxlength: 512,
    },
    aliases: {
      type: [String],
      default: undefined,
    },
    summary: {
      type: String,
      maxlength: 2048,
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
    sourceParentRecordId: {
      type: String,
      maxlength: 256,
      index: true,
    },
    sourceUpdatedAt: {
      type: Date,
      index: true,
    },
    attributes: {
      type: Schema.Types.Mixed,
    },
  },
  {
    timestamps: true,
  },
);

enterpriseMemoryEntitySchema.index(
  { visibilityScope: 1, user: 1, tenantId: 1, source: 1, entityType: 1, canonicalKey: 1 },
  { unique: true },
);
enterpriseMemoryEntitySchema.index({ source: 1, entityType: 1, displayName: 1 });
enterpriseMemoryEntitySchema.index({ tenantId: 1, source: 1, entityType: 1, updatedAt: -1 });

export default enterpriseMemoryEntitySchema;
