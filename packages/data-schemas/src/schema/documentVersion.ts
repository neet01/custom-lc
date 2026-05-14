import mongoose, { Schema, Document, Types } from 'mongoose';
import type { DocumentExtractionKind, DocumentPipelineStatus } from './document';

export interface IDocumentVersion extends Document {
  documentId: Types.ObjectId;
  sourceFileId: string;
  versionNumber: number;
  filename: string;
  mimeType: string;
  bytes: number;
  source: string;
  context?: string;
  sourceFilepath?: string;
  status: DocumentPipelineStatus;
  extractionKind: DocumentExtractionKind;
  textLength?: number;
  chunkCount?: number;
  tenantId?: string;
  createdAt?: Date;
  updatedAt?: Date;
}

const documentVersionSchema = new Schema<IDocumentVersion>(
  {
    documentId: {
      type: mongoose.Schema.Types.ObjectId,
      ref: 'Document',
      required: true,
      index: true,
    },
    sourceFileId: {
      type: String,
      required: true,
      index: true,
    },
    versionNumber: {
      type: Number,
      required: true,
      default: 1,
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
    sourceFilepath: {
      type: String,
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
    textLength: {
      type: Number,
      default: 0,
    },
    chunkCount: {
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

documentVersionSchema.index({ documentId: 1, versionNumber: 1 }, { unique: true });
documentVersionSchema.index({ tenantId: 1, createdAt: -1 });

export default documentVersionSchema;
