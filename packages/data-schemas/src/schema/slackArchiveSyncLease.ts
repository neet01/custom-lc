import mongoose, { Schema, Document, Types } from 'mongoose';

export interface ISlackArchiveSyncLease extends Document {
  leaseKey: string;
  leaseType: 'user' | 'slot';
  ownerToken: string;
  user?: Types.ObjectId;
  leaseExpiresAt: Date;
  lastHeartbeatAt?: Date;
  tenantId?: string;
  createdAt?: Date;
  updatedAt?: Date;
}

const slackArchiveSyncLeaseSchema = new Schema<ISlackArchiveSyncLease>(
  {
    leaseKey: {
      type: String,
      required: true,
      unique: true,
    },
    leaseType: {
      type: String,
      enum: ['user', 'slot'],
      required: true,
      index: true,
    },
    ownerToken: {
      type: String,
      required: true,
      index: true,
    },
    user: {
      type: mongoose.Schema.Types.ObjectId,
      ref: 'User',
      index: true,
    },
    leaseExpiresAt: {
      type: Date,
      required: true,
      index: true,
    },
    lastHeartbeatAt: {
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

slackArchiveSyncLeaseSchema.index({ leaseType: 1, leaseExpiresAt: 1 });
slackArchiveSyncLeaseSchema.index({ user: 1, leaseType: 1 });

export default slackArchiveSyncLeaseSchema;
