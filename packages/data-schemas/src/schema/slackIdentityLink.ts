import mongoose, { Schema, Document, Types } from 'mongoose';

export type SlackIdentityLinkStatus = 'linked' | 'pending' | 'revoked';

export interface ISlackIdentityLink extends Document {
  user: Types.ObjectId;
  slackUserId: string;
  teamId?: string;
  teamName?: string;
  enterpriseId?: string;
  enterpriseName?: string;
  slackEmail?: string;
  slackDisplayName?: string;
  status: SlackIdentityLinkStatus;
  source?: 'oauth_install' | 'admin_link' | 'manual';
  linkedAt?: Date;
  lastVerifiedAt?: Date;
  tenantId?: string;
  createdAt?: Date;
  updatedAt?: Date;
}

const slackIdentityLinkSchema = new Schema<ISlackIdentityLink>(
  {
    user: {
      type: mongoose.Schema.Types.ObjectId,
      ref: 'User',
      required: true,
      index: true,
    },
    slackUserId: {
      type: String,
      required: true,
      maxlength: 128,
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
    slackEmail: {
      type: String,
      maxlength: 320,
      index: true,
    },
    slackDisplayName: {
      type: String,
      maxlength: 256,
    },
    status: {
      type: String,
      enum: ['linked', 'pending', 'revoked'],
      default: 'linked',
      index: true,
    },
    source: {
      type: String,
      enum: ['oauth_install', 'admin_link', 'manual'],
      default: 'oauth_install',
    },
    linkedAt: {
      type: Date,
      index: true,
    },
    lastVerifiedAt: {
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

slackIdentityLinkSchema.index(
  { user: 1, slackUserId: 1, teamId: 1, enterpriseId: 1 },
  { unique: true, sparse: true },
);
slackIdentityLinkSchema.index(
  { slackUserId: 1, teamId: 1, enterpriseId: 1, status: 1 },
  { sparse: true },
);

export default slackIdentityLinkSchema;
