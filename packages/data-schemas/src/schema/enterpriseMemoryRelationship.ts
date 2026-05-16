import mongoose, { Schema, Document, Types } from 'mongoose';
import type { EnterpriseMemoryVisibilityScope } from './enterpriseMemoryEntity';

export interface IEnterpriseMemoryRelationship extends Document {
  user?: Types.ObjectId;
  tenantId?: string;
  visibilityScope: EnterpriseMemoryVisibilityScope;
  source: string;
  relationshipType: string;
  fromEntityId: Types.ObjectId;
  toEntityId: Types.ObjectId;
  sourceRecordType?: string;
  sourceRecordId?: string;
  sourceUpdatedAt?: Date;
  attributes?: Record<string, unknown>;
  createdAt?: Date;
  updatedAt?: Date;
}

const enterpriseMemoryRelationshipSchema = new Schema<IEnterpriseMemoryRelationship>(
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
    relationshipType: {
      type: String,
      required: true,
      maxlength: 64,
      index: true,
    },
    fromEntityId: {
      type: mongoose.Schema.Types.ObjectId,
      ref: 'EnterpriseMemoryEntity',
      required: true,
      index: true,
    },
    toEntityId: {
      type: mongoose.Schema.Types.ObjectId,
      ref: 'EnterpriseMemoryEntity',
      required: true,
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

enterpriseMemoryRelationshipSchema.index(
  {
    visibilityScope: 1,
    user: 1,
    tenantId: 1,
    source: 1,
    relationshipType: 1,
    fromEntityId: 1,
    toEntityId: 1,
    sourceRecordType: 1,
    sourceRecordId: 1,
  },
  { unique: true },
);
enterpriseMemoryRelationshipSchema.index({ fromEntityId: 1, relationshipType: 1 });
enterpriseMemoryRelationshipSchema.index({ toEntityId: 1, relationshipType: 1 });

export default enterpriseMemoryRelationshipSchema;
