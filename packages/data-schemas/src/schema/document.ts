import mongoose, { Schema, Document, Types } from 'mongoose';

export type DocumentPipelineStatus = 'pending' | 'processing' | 'ready' | 'failed';
export type DocumentExtractionKind = 'none' | 'text' | 'structured';

export interface IDocumentRecord extends Document {
  user: Types.ObjectId;
  sourceFileId: string;
  conversationId?: string;
  messageId?: string;
  filename: string;
  mimeType: string;
  bytes: number;
  source: string;
  context?: string;
  status: DocumentPipelineStatus;
  extractionKind: DocumentExtractionKind;
  latestVersionId?: Types.ObjectId;
  currentJobId?: Types.ObjectId;
  tenantId?: string;
  createdAt?: Date;
  updatedAt?: Date;
}

const documentSchema = new Schema<IDocumentRecord>(
  {
    user: {
      type: mongoose.Schema.Types.ObjectId,
      ref: 'User',
      required: true,
      index: true,
    },
    sourceFileId: {
      type: String,
      required: true,
      index: true,
    },
    conversationId: {
      type: String,
      index: true,
    },
    messageId: {
      type: String,
      index: true,
    },
    filename: {
      type: String,
      required: true,
    },
    mimeType: {
      type: String,
      required: true,
      index: true,
    },
    bytes: {
      type: Number,
      required: true,
      default: 0,
    },
    source: {
      type: String,
      required: true,
    },
    context: {
      type: String,
      index: true,
    },
    status: {
      type: String,
      enum: ['pending', 'processing', 'ready', 'failed'],
      required: true,
      default: 'pending',
      index: true,
    },
    extractionKind: {
      type: String,
      enum: ['none', 'text', 'structured'],
      required: true,
      default: 'none',
    },
    latestVersionId: {
      type: mongoose.Schema.Types.ObjectId,
      ref: 'DocumentVersion',
    },
    currentJobId: {
      type: mongoose.Schema.Types.ObjectId,
      ref: 'DocumentJob',
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

documentSchema.index({ sourceFileId: 1, tenantId: 1 }, { unique: true });
documentSchema.index({ user: 1, createdAt: -1 });
documentSchema.index({ tenantId: 1, createdAt: -1 });

export default documentSchema;
