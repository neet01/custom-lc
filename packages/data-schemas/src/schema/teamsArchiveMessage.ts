import mongoose, { Schema, Document, Types } from 'mongoose';

export interface ITeamsArchiveAttachment {
  id?: string;
  name?: string;
  contentType?: string;
  contentUrl?: string;
}

export interface ITeamsArchiveMention {
  id?: string;
  displayName?: string;
  mentionedUserId?: string;
}

export interface ITeamsArchiveMessage extends Document {
  user: Types.ObjectId;
  graphChatId: string;
  graphMessageId: string;
  replyToId?: string;
  fromDisplayName?: string;
  fromEmail?: string;
  fromUserId?: string;
  subject?: string;
  summary?: string;
  importance?: string;
  messageType?: string;
  bodyContentType?: string;
  bodyPreview?: string;
  bodyContent?: string;
  bodyText?: string;
  attachments?: ITeamsArchiveAttachment[];
  mentions?: ITeamsArchiveMention[];
  webUrl?: string;
  sentDateTime?: Date;
  lastModifiedDateTime?: Date;
  deletedDateTime?: Date;
  etag?: string;
  tenantId?: string;
  createdAt?: Date;
  updatedAt?: Date;
}

const attachmentSchema = new Schema<ITeamsArchiveAttachment>(
  {
    id: { type: String, maxlength: 128 },
    name: { type: String, maxlength: 512 },
    contentType: { type: String, maxlength: 256 },
    contentUrl: { type: String, maxlength: 2048 },
  },
  { _id: false },
);

const mentionSchema = new Schema<ITeamsArchiveMention>(
  {
    id: { type: String, maxlength: 128 },
    displayName: { type: String, maxlength: 256 },
    mentionedUserId: { type: String, maxlength: 128 },
  },
  { _id: false },
);

const teamsArchiveMessageSchema = new Schema<ITeamsArchiveMessage>(
  {
    user: {
      type: mongoose.Schema.Types.ObjectId,
      ref: 'User',
      required: true,
      index: true,
    },
    graphChatId: {
      type: String,
      required: true,
      index: true,
    },
    graphMessageId: {
      type: String,
      required: true,
      index: true,
    },
    replyToId: {
      type: String,
      maxlength: 128,
      index: true,
    },
    fromDisplayName: {
      type: String,
      maxlength: 256,
      index: true,
    },
    fromEmail: {
      type: String,
      maxlength: 320,
      index: true,
    },
    fromUserId: {
      type: String,
      maxlength: 128,
      index: true,
    },
    subject: {
      type: String,
      maxlength: 512,
    },
    summary: {
      type: String,
      maxlength: 1024,
    },
    importance: {
      type: String,
      maxlength: 32,
      index: true,
    },
    messageType: {
      type: String,
      maxlength: 64,
      index: true,
    },
    bodyContentType: {
      type: String,
      maxlength: 32,
    },
    bodyPreview: {
      type: String,
      maxlength: 2048,
    },
    bodyContent: {
      type: String,
    },
    bodyText: {
      type: String,
    },
    attachments: {
      type: [attachmentSchema],
      default: undefined,
    },
    mentions: {
      type: [mentionSchema],
      default: undefined,
    },
    webUrl: {
      type: String,
      maxlength: 2048,
    },
    sentDateTime: {
      type: Date,
      index: true,
    },
    lastModifiedDateTime: {
      type: Date,
      index: true,
    },
    deletedDateTime: {
      type: Date,
      index: true,
    },
    etag: {
      type: String,
      maxlength: 256,
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

teamsArchiveMessageSchema.index({ user: 1, graphMessageId: 1 }, { unique: true });
teamsArchiveMessageSchema.index({ user: 1, graphChatId: 1, sentDateTime: -1 });
teamsArchiveMessageSchema.index({ tenantId: 1, graphChatId: 1, sentDateTime: -1 });

export default teamsArchiveMessageSchema;
