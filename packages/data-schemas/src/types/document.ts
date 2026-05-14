import type { Types } from 'mongoose';
import type {
  DocumentExtractionKind,
  DocumentPipelineStatus,
} from '~/schema/document';
import type { DocumentJobStatus, DocumentJobType } from '~/schema/documentJob';

export type { DocumentExtractionKind, DocumentPipelineStatus, DocumentJobStatus, DocumentJobType };

export interface DocumentRecordData {
  user: string;
  sourceFileId: string;
  conversationId?: string;
  messageId?: string;
  filename: string;
  mimeType: string;
  bytes: number;
  source: string;
  context?: string;
  status?: DocumentPipelineStatus;
  extractionKind?: DocumentExtractionKind;
  latestVersionId?: Types.ObjectId | string;
  currentJobId?: Types.ObjectId | string;
}

export interface DocumentVersionData {
  documentId: string | Types.ObjectId;
  sourceFileId: string;
  versionNumber?: number;
  filename: string;
  mimeType: string;
  bytes: number;
  source: string;
  context?: string;
  sourceFilepath?: string;
  status?: DocumentPipelineStatus;
  extractionKind?: DocumentExtractionKind;
  textLength?: number;
  chunkCount?: number;
}

export interface DocumentJobData {
  documentId: string | Types.ObjectId;
  documentVersionId: string | Types.ObjectId;
  user: string;
  jobType: DocumentJobType;
  status?: DocumentJobStatus;
  attempts?: number;
  errorMessage?: string;
  startedAt?: Date;
  completedAt?: Date;
}
