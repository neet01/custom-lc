import type { Model, FilterQuery } from 'mongoose';
import type { IDocumentRecord } from '~/schema/document';
import type { IDocumentVersion } from '~/schema/documentVersion';
import type { IDocumentJob } from '~/schema/documentJob';

export function createDocumentMethods(mongoose: typeof import('mongoose')) {
  async function getDocumentBySourceFileId(sourceFileId: string): Promise<IDocumentRecord | null> {
    const Document = mongoose.models.Document as Model<IDocumentRecord>;
    return Document.findOne({ sourceFileId }).lean();
  }

  async function createDocument(
    data: Partial<IDocumentRecord> & { sourceFileId: string },
  ): Promise<IDocumentRecord | null> {
    const Document = mongoose.models.Document as Model<IDocumentRecord>;
    return Document.findOneAndUpdate(
      { sourceFileId: data.sourceFileId },
      data,
      { new: true, upsert: true },
    ).lean();
  }

  async function updateDocument(
    filter: FilterQuery<IDocumentRecord>,
    update: Partial<IDocumentRecord>,
  ): Promise<IDocumentRecord | null> {
    const Document = mongoose.models.Document as Model<IDocumentRecord>;
    return Document.findOneAndUpdate(filter, { $set: update }, { new: true }).lean();
  }

  async function createDocumentVersion(
    data: Partial<IDocumentVersion> & { documentId: IDocumentVersion['documentId'] },
  ): Promise<IDocumentVersion | null> {
    const DocumentVersion = mongoose.models.DocumentVersion as Model<IDocumentVersion>;
    return DocumentVersion.create(data);
  }

  async function createDocumentJob(
    data: Partial<IDocumentJob> & { documentId: IDocumentJob['documentId']; documentVersionId: IDocumentJob['documentVersionId'] },
  ): Promise<IDocumentJob | null> {
    const DocumentJob = mongoose.models.DocumentJob as Model<IDocumentJob>;
    return DocumentJob.create(data);
  }

  async function updateDocumentJob(
    filter: FilterQuery<IDocumentJob>,
    update: Partial<IDocumentJob>,
  ): Promise<IDocumentJob | null> {
    const DocumentJob = mongoose.models.DocumentJob as Model<IDocumentJob>;
    return DocumentJob.findOneAndUpdate(filter, { $set: update }, { new: true }).lean();
  }

  return {
    getDocumentBySourceFileId,
    createDocument,
    updateDocument,
    createDocumentVersion,
    createDocumentJob,
    updateDocumentJob,
  };
}

export type DocumentMethods = ReturnType<typeof createDocumentMethods>;
