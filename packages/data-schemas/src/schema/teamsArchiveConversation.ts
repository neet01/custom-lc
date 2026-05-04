import mongoose, { Schema, Document, Types } from 'mongoose';

export interface ITeamsArchiveParticipant {
  displayName?: string;
  email?: string;
  userId?: string;
}

export interface ITeamsArchiveConversation extends Document {
  user: Types.ObjectId;
  graphChatId: string;
  chatType?: string;
  topic?: string;
  webUrl?: string;
  participants?: ITeamsArchiveParticipant[];
  lastMessageAt?: Date;
  lastSyncedAt?: Date;
  sourceUpdatedAt?: Date;
  messageCount?: number;
  tenantId?: string;
  createdAt?: Date;
  updatedAt?: Date;
}

const participantSchema = new Schema<ITeamsArchiveParticipant>(
  {
    displayName: { type: String, maxlength: 256 },
    email: { type: String, maxlength: 320 },
    userId: { type: String, maxlength: 128 },
  },
  { _id: false },
);

const teamsArchiveConversationSchema = new Schema<ITeamsArchiveConversation>(
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
    chatType: {
      type: String,
      maxlength: 64,
      index: true,
    },
    topic: {
      type: String,
      maxlength: 512,
    },
    webUrl: {
      type: String,
      maxlength: 2048,
    },
    participants: {
      type: [participantSchema],
      default: undefined,
    },
    lastMessageAt: {
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
    tenantId: {
      type: String,
      index: true,
    },
  },
  {
    timestamps: true,
  },
);

teamsArchiveConversationSchema.index({ user: 1, graphChatId: 1 }, { unique: true });
teamsArchiveConversationSchema.index({ user: 1, lastMessageAt: -1 });
teamsArchiveConversationSchema.index({ tenantId: 1, lastMessageAt: -1 });

export default teamsArchiveConversationSchema;
