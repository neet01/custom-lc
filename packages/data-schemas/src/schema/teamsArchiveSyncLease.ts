import mongoose, { Schema, Document, Types } from 'mongoose';

export type TeamsArchiveSyncLeaseType = 'user' | 'slot';

export interface ITeamsArchiveSyncLease extends Document {
  leaseKey: string;
  leaseType: TeamsArchiveSyncLeaseType;
  ownerToken: string;
  user?: Types.ObjectId;
  leaseExpiresAt: Date;
  lastHeartbeatAt?: Date;
  tenantId?: string;
  createdAt?: Date;
  updatedAt?: Date;
}

const teamsArchiveSyncLeaseSchema = new Schema<ITeamsArchiveSyncLease>(
  {
    leaseKey: {
      type: String,
      required: true,
      unique: true,
      maxlength: 128,
      index: true,
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
      maxlength: 128,
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

teamsArchiveSyncLeaseSchema.index({ leaseType: 1, leaseExpiresAt: 1 });
teamsArchiveSyncLeaseSchema.index({ user: 1, leaseType: 1 });

export default teamsArchiveSyncLeaseSchema;
