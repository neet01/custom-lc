import mongoose, { Schema, Document, Types } from 'mongoose';

export interface ISlackArchiveParticipant {
  slackUserId?: string;
  displayName?: string;
  realName?: string;
  username?: string;
  email?: string;
  isBot?: boolean;
  isAppUser?: boolean;
}

export interface ISlackArchiveConversation extends Document {
  user: Types.ObjectId;
  slackConversationId: string;
  teamId?: string;
  enterpriseId?: string;
  conversationType?: 'public_channel' | 'private_channel' | 'im' | 'mpim';
  name?: string;
  topic?: string;
  purpose?: string;
  isArchived?: boolean;
  isShared?: boolean;
  isExtShared?: boolean;
  isOrgShared?: boolean;
  isSlackConnect?: boolean;
  participants?: ISlackArchiveParticipant[];
  syncStatus?: 'pending' | 'running' | 'complete' | 'failed';
  syncCursor?: string;
  syncError?: string;
  syncAttemptCount?: number;
  sourceDiscoveredAt?: Date;
  sourceLastMessageAt?: Date;
  syncStartedAt?: Date;
  syncCompletedAt?: Date;
  lastMessageSyncAt?: Date;
  lastMessageAt?: Date;
  lastHumanMessageAt?: Date;
  lastMeaningfulMessageAt?: Date;
  lastSystemMessageAt?: Date;
  lastSyncedAt?: Date;
  sourceUpdatedAt?: Date;
  messageCount?: number;
  humanMessageCount?: number;
  systemMessageCount?: number;
  meaningfulMessageCount?: number;
  tenantId?: string;
  createdAt?: Date;
  updatedAt?: Date;
}

const participantSchema = new Schema<ISlackArchiveParticipant>(
  {
    slackUserId: { type: String, maxlength: 128 },
    displayName: { type: String, maxlength: 256 },
    realName: { type: String, maxlength: 256 },
    username: { type: String, maxlength: 256 },
    email: { type: String, maxlength: 320 },
    isBot: { type: Boolean, default: false },
    isAppUser: { type: Boolean, default: false },
  },
  { _id: false },
);

const slackArchiveConversationSchema = new Schema<ISlackArchiveConversation>(
  {
    user: {
      type: mongoose.Schema.Types.ObjectId,
      ref: 'User',
      required: true,
      index: true,
    },
    slackConversationId: {
      type: String,
      required: true,
      index: true,
    },
    teamId: {
      type: String,
      maxlength: 64,
      index: true,
    },
    enterpriseId: {
      type: String,
      maxlength: 64,
      index: true,
    },
    conversationType: {
      type: String,
      enum: ['public_channel', 'private_channel', 'im', 'mpim'],
      index: true,
    },
    name: {
      type: String,
      maxlength: 256,
    },
    topic: {
      type: String,
      maxlength: 1024,
    },
    purpose: {
      type: String,
      maxlength: 2048,
    },
    isArchived: {
      type: Boolean,
      default: false,
      index: true,
    },
    isShared: {
      type: Boolean,
      default: false,
      index: true,
    },
    isExtShared: {
      type: Boolean,
      default: false,
      index: true,
    },
    isOrgShared: {
      type: Boolean,
      default: false,
      index: true,
    },
    isSlackConnect: {
      type: Boolean,
      default: false,
      index: true,
    },
    participants: {
      type: [participantSchema],
      default: undefined,
    },
    syncStatus: {
      type: String,
      enum: ['pending', 'running', 'complete', 'failed'],
      default: 'pending',
      index: true,
    },
    syncCursor: {
      type: String,
      maxlength: 4096,
    },
    syncError: {
      type: String,
      maxlength: 2000,
    },
    syncAttemptCount: {
      type: Number,
      default: 0,
    },
    sourceDiscoveredAt: {
      type: Date,
      index: true,
    },
    sourceLastMessageAt: {
      type: Date,
      index: true,
    },
    syncStartedAt: {
      type: Date,
      index: true,
    },
    syncCompletedAt: {
      type: Date,
      index: true,
    },
    lastMessageSyncAt: {
      type: Date,
      index: true,
    },
    lastMessageAt: {
      type: Date,
      index: true,
    },
    lastHumanMessageAt: {
      type: Date,
      index: true,
    },
    lastMeaningfulMessageAt: {
      type: Date,
      index: true,
    },
    lastSystemMessageAt: {
      type: Date,
      index: true,
    },
    lastSyncedAt: {
      type: Date,
      index: true,
    },
    sourceUpdatedAt: {
      type: Date,
      index: true,
    },
    messageCount: {
      type: Number,
      default: 0,
    },
    humanMessageCount: {
      type: Number,
      default: 0,
      index: true,
    },
    systemMessageCount: {
      type: Number,
      default: 0,
      index: true,
    },
    meaningfulMessageCount: {
      type: Number,
      default: 0,
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

slackArchiveConversationSchema.index({ user: 1, slackConversationId: 1 }, { unique: true });
slackArchiveConversationSchema.index({ user: 1, conversationType: 1, lastMeaningfulMessageAt: -1 });
slackArchiveConversationSchema.index({ user: 1, lastMessageAt: -1, updatedAt: -1 });
slackArchiveConversationSchema.index({ tenantId: 1, lastMessageAt: -1 });

export default slackArchiveConversationSchema;
