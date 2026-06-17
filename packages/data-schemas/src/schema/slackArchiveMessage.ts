import mongoose, { Schema, Document, Types } from 'mongoose';

export interface ISlackArchiveReaction {
  name?: string;
  count?: number;
  users?: string[];
}

export interface ISlackArchiveMention {
  slackUserId?: string;
  displayName?: string;
}

export interface ISlackArchiveAttachment {
  id?: string;
  title?: string;
  text?: string;
  fallback?: string;
  serviceName?: string;
  titleLink?: string;
}

export interface ISlackArchiveFile {
  id?: string;
  name?: string;
  title?: string;
  mimetype?: string;
  filetype?: string;
  prettyType?: string;
  urlPrivate?: string;
}

export interface ISlackArchiveMessage extends Document {
  user: Types.ObjectId;
  slackConversationId: string;
  slackMessageTs: string;
  teamId?: string;
  enterpriseId?: string;
  slackUserId?: string;
  botId?: string;
  username?: string;
  displayName?: string;
  subtype?: string;
  text?: string;
  normalizedText?: string;
  threadTs?: string;
  parentUserId?: string;
  replyCount?: number;
  replyUsers?: string[];
  latestReplyTs?: string;
  reactions?: ISlackArchiveReaction[];
  mentions?: ISlackArchiveMention[];
  attachments?: ISlackArchiveAttachment[];
  files?: ISlackArchiveFile[];
  raw?: Record<string, unknown>;
  sentAt?: Date;
  editedAt?: Date;
  deletedAt?: Date;
  normalizedTextLength?: number;
  isSystemLikeMessage?: boolean;
  isChunkable?: boolean;
  skipChunkReason?: string;
  tenantId?: string;
  createdAt?: Date;
  updatedAt?: Date;
}

const reactionSchema = new Schema<ISlackArchiveReaction>(
  {
    name: { type: String, maxlength: 128 },
    count: { type: Number, default: 0 },
    users: {
      type: [String],
      default: undefined,
    },
  },
  { _id: false },
);

const mentionSchema = new Schema<ISlackArchiveMention>(
  {
    slackUserId: { type: String, maxlength: 128 },
    displayName: { type: String, maxlength: 256 },
  },
  { _id: false },
);

const attachmentSchema = new Schema<ISlackArchiveAttachment>(
  {
    id: { type: String, maxlength: 128 },
    title: { type: String, maxlength: 512 },
    text: { type: String },
    fallback: { type: String, maxlength: 2048 },
    serviceName: { type: String, maxlength: 256 },
    titleLink: { type: String, maxlength: 2048 },
  },
  { _id: false },
);

const fileSchema = new Schema<ISlackArchiveFile>(
  {
    id: { type: String, maxlength: 128 },
    name: { type: String, maxlength: 512 },
    title: { type: String, maxlength: 512 },
    mimetype: { type: String, maxlength: 256 },
    filetype: { type: String, maxlength: 128 },
    prettyType: { type: String, maxlength: 256 },
    urlPrivate: { type: String, maxlength: 2048 },
  },
  { _id: false },
);

const slackArchiveMessageSchema = new Schema<ISlackArchiveMessage>(
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
    slackMessageTs: {
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
    slackUserId: {
      type: String,
      maxlength: 128,
      index: true,
    },
    botId: {
      type: String,
      maxlength: 128,
      index: true,
    },
    username: {
      type: String,
      maxlength: 256,
    },
    displayName: {
      type: String,
      maxlength: 256,
      index: true,
    },
    subtype: {
      type: String,
      maxlength: 128,
      index: true,
    },
    text: {
      type: String,
    },
    normalizedText: {
      type: String,
    },
    threadTs: {
      type: String,
      maxlength: 64,
      index: true,
    },
    parentUserId: {
      type: String,
      maxlength: 128,
      index: true,
    },
    replyCount: {
      type: Number,
      default: 0,
    },
    replyUsers: {
      type: [String],
      default: undefined,
    },
    latestReplyTs: {
      type: String,
      maxlength: 64,
    },
    reactions: {
      type: [reactionSchema],
      default: undefined,
    },
    mentions: {
      type: [mentionSchema],
      default: undefined,
    },
    attachments: {
      type: [attachmentSchema],
      default: undefined,
    },
    files: {
      type: [fileSchema],
      default: undefined,
    },
    raw: {
      type: Schema.Types.Mixed,
    },
    sentAt: {
      type: Date,
      index: true,
    },
    editedAt: {
      type: Date,
      index: true,
    },
    deletedAt: {
      type: Date,
      index: true,
    },
    normalizedTextLength: {
      type: Number,
      default: 0,
      index: true,
    },
    isSystemLikeMessage: {
      type: Boolean,
      default: false,
      index: true,
    },
    isChunkable: {
      type: Boolean,
      default: false,
      index: true,
    },
    skipChunkReason: {
      type: String,
      maxlength: 128,
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

slackArchiveMessageSchema.index({ user: 1, slackConversationId: 1, slackMessageTs: 1 }, { unique: true });
slackArchiveMessageSchema.index({ user: 1, slackConversationId: 1, sentAt: -1 });
slackArchiveMessageSchema.index({ user: 1, threadTs: 1, sentAt: 1 });
slackArchiveMessageSchema.index({ user: 1, slackUserId: 1, sentAt: -1 });
slackArchiveMessageSchema.index({ tenantId: 1, slackConversationId: 1, sentAt: -1 });

export default slackArchiveMessageSchema;
