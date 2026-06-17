import mongoose, { Schema, Document, Types } from 'mongoose';

export type SlackWorkspaceInstallStatus = 'active' | 'revoked' | 'error';

export interface ISlackWorkspaceInstall extends Document {
  installedByUser?: Types.ObjectId;
  teamId?: string;
  teamName?: string;
  enterpriseId?: string;
  enterpriseName?: string;
  botUserId?: string;
  botAccessToken?: string;
  botScopes?: string;
  userScopes?: string;
  installPayload?: Record<string, unknown>;
  status: SlackWorkspaceInstallStatus;
  installedAt?: Date;
  lastValidatedAt?: Date;
  tenantId?: string;
  createdAt?: Date;
  updatedAt?: Date;
}

const slackWorkspaceInstallSchema = new Schema<ISlackWorkspaceInstall>(
  {
    installedByUser: {
      type: mongoose.Schema.Types.ObjectId,
      ref: 'User',
      index: true,
    },
    teamId: {
      type: String,
      maxlength: 64,
      index: true,
    },
    teamName: {
      type: String,
      maxlength: 256,
    },
    enterpriseId: {
      type: String,
      maxlength: 64,
      index: true,
    },
    enterpriseName: {
      type: String,
      maxlength: 256,
    },
    botUserId: {
      type: String,
      maxlength: 128,
      index: true,
    },
    botAccessToken: {
      type: String,
      maxlength: 4096,
    },
    botScopes: {
      type: String,
      maxlength: 4096,
    },
    userScopes: {
      type: String,
      maxlength: 4096,
    },
    installPayload: {
      type: Schema.Types.Mixed,
    },
    status: {
      type: String,
      enum: ['active', 'revoked', 'error'],
      default: 'active',
      index: true,
    },
    installedAt: {
      type: Date,
      index: true,
    },
    lastValidatedAt: {
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

slackWorkspaceInstallSchema.index({ enterpriseId: 1, teamId: 1 }, { unique: true, sparse: true });
slackWorkspaceInstallSchema.index({ teamId: 1, status: 1 }, { sparse: true });

export default slackWorkspaceInstallSchema;
