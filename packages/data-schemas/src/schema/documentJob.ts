import mongoose, { Schema, Document, Types } from 'mongoose';

export type DocumentJobType = 'extract' | 'chunk' | 'index';
export type DocumentJobStatus = 'pending' | 'running' | 'success' | 'failure';

export interface IDocumentJob extends Document {
  documentId: Types.ObjectId;
  documentVersionId: Types.ObjectId;
  user: Types.ObjectId;
  jobType: DocumentJobType;
  status: DocumentJobStatus;
  attempts?: number;
  errorMessage?: string;
  startedAt?: Date;
  completedAt?: Date;
  tenantId?: string;
  createdAt?: Date;
  updatedAt?: Date;
}

const documentJobSchema = new Schema<IDocumentJob>(
  {
    documentId: {
      type: mongoose.Schema.Types.ObjectId,
      ref: 'Document',
      required: true,
      index: true,
    },
    documentVersionId: {
      type: mongoose.Schema.Types.ObjectId,
      ref: 'DocumentVersion',
      required: true,
      index: true,
    },
    user: {
      type: mongoose.Schema.Types.ObjectId,
      ref: 'User',
      required: true,
      index: true,
    },
    jobType: {
      type: String,
      enum: ['extract', 'chunk', 'index'],
      required: true,
      index: true,
    },
    status: {
      type: String,
      enum: ['pending', 'running', 'success', 'failure'],
      required: true,
      default: 'pending',
      index: true,
    },
    attempts: {
      type: Number,
      default: 0,
    },
    errorMessage: {
      type: String,
      maxlength: 1000,
    },
    startedAt: {
      type: Date,
      index: true,
    },
    completedAt: {
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

documentJobSchema.index({ documentId: 1, createdAt: -1 });
documentJobSchema.index({ tenantId: 1, createdAt: -1 });

export default documentJobSchema;
