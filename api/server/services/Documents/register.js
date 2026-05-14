const {
  FileSources,
  textMimeTypes,
  excelMimeTypes,
  applicationMimeTypes,
  documentParserMimeTypes,
} = require('librechat-data-provider');
const { logger } = require('@librechat/data-schemas');
const db = require('~/models');

const documentIntelligenceMimeTypes = [
  textMimeTypes,
  excelMimeTypes,
  applicationMimeTypes,
  ...documentParserMimeTypes,
];

const isIndexableDocumentType = (mimeType) =>
  typeof mimeType === 'string' &&
  documentIntelligenceMimeTypes.some((pattern) => pattern.test(mimeType));

const deriveExtractionKind = (file) => {
  if (file?.source === FileSources.text) {
    return 'text';
  }
  return 'none';
};

const deriveInitialJobType = (file) => {
  if (file?.source === FileSources.text) {
    return 'chunk';
  }
  return 'extract';
};

async function registerDocumentUpload({
  file,
  userId,
  conversationId,
  messageId,
}) {
  if (!file?.file_id || !isIndexableDocumentType(file.type)) {
    return null;
  }

  const existing = await db.getDocumentBySourceFileId(file.file_id);
  if (existing) {
    return existing;
  }

  const extractionKind = deriveExtractionKind(file);
  const textLength = typeof file.text === 'string' ? Buffer.byteLength(file.text, 'utf8') : 0;

  const document = await db.createDocument({
    user: userId,
    sourceFileId: file.file_id,
    conversationId,
    messageId,
    filename: file.filename,
    mimeType: file.type,
    bytes: file.bytes ?? 0,
    source: file.source,
    context: file.context,
    status: 'pending',
    extractionKind,
  });

  if (!document?._id) {
    throw new Error(`Failed to create document record for file ${file.file_id}`);
  }

  const version = await db.createDocumentVersion({
    documentId: document._id,
    sourceFileId: file.file_id,
    versionNumber: 1,
    filename: file.filename,
    mimeType: file.type,
    bytes: file.bytes ?? 0,
    source: file.source,
    context: file.context,
    sourceFilepath: file.filepath,
    status: 'pending',
    extractionKind,
    textLength,
    chunkCount: 0,
  });

  if (!version?._id) {
    throw new Error(`Failed to create document version for file ${file.file_id}`);
  }

  const job = await db.createDocumentJob({
    documentId: document._id,
    documentVersionId: version._id,
    user: userId,
    jobType: deriveInitialJobType(file),
    status: 'pending',
    attempts: 0,
  });

  await db.updateDocument(
    { _id: document._id },
    {
      latestVersionId: version._id,
      currentJobId: job?._id,
    },
  );

  logger.info(
    `[Documents] Registered file ${file.file_id} as document ${document._id.toString()} (job=${job?._id?.toString?.() ?? 'n/a'})`,
  );

  return document;
}

async function maybeRegisterDocumentUpload(params) {
  try {
    return await registerDocumentUpload(params);
  } catch (error) {
    logger.error('[Documents] Failed to register uploaded document:', error);
    return null;
  }
}

module.exports = {
  isIndexableDocumentType,
  registerDocumentUpload,
  maybeRegisterDocumentUpload,
};
