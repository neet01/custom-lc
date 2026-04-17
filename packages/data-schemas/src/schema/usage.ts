import mongoose, { Schema, Document, Types } from 'mongoose';
import type { UsageSource } from '~/types';

export interface IUsage extends Document {
  user: Types.ObjectId;
  conversationId: string;
  messageId?: string;
  requestId?: string;
  sessionId?: string;
  model?: string;
  provider?: string;
  endpoint?: string;
  context?: string;
  source?: UsageSource;
  inputTokens: number;
  outputTokens: number;
  totalTokens: number;
  cacheCreationTokens?: number;
  cacheReadTokens?: number;
  latencyMs?: number;
  createdAt?: Date;
  updatedAt?: Date;
  tenantId?: string;
}

const usageSchema = new Schema<IUsage>(
  {
    user: {
      type: mongoose.Schema.Types.ObjectId,
      ref: 'User',
      index: true,
      required: true,
    },
    conversationId: {
      type: String,
      index: true,
      required: true,
    },
    messageId: {
      type: String,
      index: true,
    },
    requestId: {
      type: String,
      index: true,
    },
    sessionId: {
      type: String,
      index: true,
    },
    model: {
      type: String,
      index: true,
    },
    provider: {
      type: String,
      index: true,
    },
    endpoint: {
      type: String,
      index: true,
    },
    context: {
      type: String,
      index: true,
    },
    source: {
      type: String,
      enum: ['agent', 'assistant', 'tool', 'system'],
      default: 'system',
      index: true,
    },
    inputTokens: {
      type: Number,
      required: true,
      min: 0,
    },
    outputTokens: {
      type: Number,
      required: true,
      min: 0,
    },
    totalTokens: {
      type: Number,
      required: true,
      min: 0,
    },
    cacheCreationTokens: {
      type: Number,
      min: 0,
    },
    cacheReadTokens: {
      type: Number,
      min: 0,
    },
    latencyMs: {
      type: Number,
      min: 0,
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

usageSchema.index({ user: 1, createdAt: -1 });
usageSchema.index({ conversationId: 1, createdAt: -1 });
usageSchema.index({ tenantId: 1, createdAt: -1 });

export default usageSchema;
