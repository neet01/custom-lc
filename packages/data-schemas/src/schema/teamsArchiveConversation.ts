import mongoose, { Schema, Document, Types } from 'mongoose';

export interface ITeamsArchiveParticipant {
  displayName?: string;
  email?: string;
  userId?: string;
  source?: 'graph' | 'inferred_from_messages' | 'inferred_from_mentions' | 'mixed' | 'unknown';
  confidence?: 'high' | 'medium' | 'low';
}

export interface ITeamsArchiveConversation extends Document {
  user: Types.ObjectId;
  graphChatId: string;
  chatType?: string;
  topic?: string;
  webUrl?: string;
  participants?: ITeamsArchiveParticipant[];
  syncStatus?: 'pending' | 'running' | 'complete' | 'failed';
  syncCursor?: string;
  syncError?: string;
  sourceDiscoveredAt?: Date;
  sourceLastMessageAt?: Date;
  syncStartedAt?: Date;
  syncCompletedAt?: Date;
  lastMessageSyncAt?: Date;
  lastMessageAt?: Date;
  lastSyncedAt?: Date;
  sourceUpdatedAt?: Date;
  messageCount?: number;
  participantMetadataSource?: 'graph' | 'inferred_from_messages' | 'inferred_from_mentions' | 'mixed' | 'unknown';
  participantConfidence?: 'high' | 'medium' | 'low';
  participantDegraded?: boolean;
  participantStats?: Record<string, unknown>;
  tenantId?: string;
  createdAt?: Date;
  updatedAt?: Date;
}

const participantSchema = new Schema<ITeamsArchiveParticipant>(
  {
    displayName: { type: String, maxlength: 256 },
    email: { type: String, maxlength: 320 },
    userId: { type: String, maxlength: 128 },
    source: {
      type: String,
      enum: ['graph', 'inferred_from_messages', 'inferred_from_mentions', 'mixed', 'unknown'],
      default: 'unknown',
    },
    confidence: {
      type: String,
      enum: ['high', 'medium', 'low'],
      default: 'low',
    },
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
    participantMetadataSource: {
      type: String,
      enum: ['graph', 'inferred_from_messages', 'inferred_from_mentions', 'mixed', 'unknown'],
      default: 'unknown',
      index: true,
    },
    participantConfidence: {
      type: String,
      enum: ['high', 'medium', 'low'],
      default: 'low',
      index: true,
    },
    participantDegraded: {
      type: Boolean,
      default: false,
      index: true,
    },
    participantStats: {
      type: Schema.Types.Mixed,
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
teamsArchiveConversationSchema.index({ user: 1, syncStatus: 1, sourceUpdatedAt: -1, updatedAt: -1 });
teamsArchiveConversationSchema.index({ user: 1, sourceLastMessageAt: -1, updatedAt: -1 });
teamsArchiveConversationSchema.index({ user: 1, participantDegraded: 1, updatedAt: -1 });
teamsArchiveConversationSchema.index({ tenantId: 1, lastMessageAt: -1 });

export default teamsArchiveConversationSchema;
