const { Readable } = require('stream');
const { v4: uuidv4 } = require('uuid');
const {
  EModelEndpoint,
  FileSources,
  FileContext,
  checkOpenAIStorage,
} = require('librechat-data-provider');
const { getOpenAIClient } = require('~/server/controllers/assistants/helpers');
const { getStrategyFunctions } = require('~/server/services/Files/strategies');
const { getFileStrategy } = require('~/server/utils/getFileStrategy');
const db = require('~/models');
const {
  inspectWordDocumentBuffer,
  transformWordDocumentBuffer,
} = require('~/server/services/Files/WordDocuments/transform');

async function streamToBuffer(stream) {
  const chunks = [];
  for await (const chunk of stream) {
    chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
  }
  return Buffer.concat(chunks);
}

async function getWordDocumentFileBuffer(req, res, file) {
  const { getDownloadStream } = getStrategyFunctions(file.source);
  if (!getDownloadStream) {
    throw new Error(`No download stream method implemented for file source: ${file.source}`);
  }

  if (checkOpenAIStorage(file.source)) {
    req.body = { ...req.body, model: file.model };
    const endpointMap = {
      [FileSources.openai]: EModelEndpoint.assistants,
      [FileSources.azure]: EModelEndpoint.azureAssistants,
    };
    const { openai } = await getOpenAIClient({
      req,
      res,
      overrideEndpoint: endpointMap[file.source],
    });

    const passThrough = await getDownloadStream(file.file_id, openai);
    const stream =
      passThrough.body && typeof passThrough.body.getReader === 'function'
        ? Readable.fromWeb(passThrough.body)
        : passThrough.body;

    return streamToBuffer(stream);
  }

  const stream = await getDownloadStream(req, file.filepath);
  return streamToBuffer(stream);
}

function resolveGeneratedFileExportStrategy(req) {
  const candidates = [
    (() => {
      try {
        return getFileStrategy(req?.config ?? {}, { context: FileContext.message_attachment });
      } catch (_error) {
        return undefined;
      }
    })(),
    req?.config?.fileStrategy,
    FileSources.local,
  ].filter(Boolean);

  const checked = new Set();
  for (const candidate of candidates) {
    if (checked.has(candidate)) {
      continue;
    }
    checked.add(candidate);

    try {
      const strategy = getStrategyFunctions(candidate);
      if (strategy?.saveBuffer) {
        return {
          fileStrategy: candidate,
          saveBuffer: strategy.saveBuffer,
        };
      }
    } catch (_error) {
      continue;
    }
  }

  throw new Error('No valid file strategy available for generated Word document exports');
}

async function saveGeneratedWordDocument({
  req,
  sourceFile,
  generated,
  conversationId,
  messageId,
}) {
  const { fileStrategy, saveBuffer } = resolveGeneratedFileExportStrategy(req);

  const outputFileId = uuidv4();
  const storedFileName = `${outputFileId}__${generated.filename}`;
  const filepath = await saveBuffer({
    userId: req.user.id,
    buffer: generated.buffer,
    fileName: storedFileName,
    basePath: 'uploads',
  });

  return db.createFile(
    {
      user: req.user.id,
      file_id: outputFileId,
      bytes: generated.bytes,
      filepath,
      filename: generated.filename,
      type: generated.mimeType,
      context: FileContext.message_attachment,
      source: fileStrategy,
      conversationId: conversationId ?? sourceFile.conversationId,
      messageId: messageId ?? sourceFile.messageId,
    },
    true,
  );
}

async function inspectWordDocumentFile({
  req,
  res,
  sourceFile,
  maxPreviewParagraphs,
}) {
  const sourceBuffer = await getWordDocumentFileBuffer(req, res, sourceFile);
  return inspectWordDocumentBuffer({
    buffer: sourceBuffer,
    sourceFilename: sourceFile.filename,
    maxPreviewParagraphs,
  });
}

async function transformWordDocumentFile({
  req,
  res,
  sourceFile,
  replaceText,
  redactPhrases,
  redactionText,
  prependText,
  appendText,
  replacementText,
  outputFilename,
  conversationId,
  messageId,
}) {
  const sourceBuffer = await getWordDocumentFileBuffer(req, res, sourceFile);
  const generated = await transformWordDocumentBuffer({
    buffer: sourceBuffer,
    sourceFilename: sourceFile.filename,
    replaceText,
    redactPhrases,
    redactionText,
    prependText,
    appendText,
    replacementText,
    outputFilename,
  });

  const file = await saveGeneratedWordDocument({
    req,
    sourceFile,
    generated,
    conversationId,
    messageId,
  });

  return {
    file,
    summary: generated.summary,
  };
}

module.exports = {
  getWordDocumentFileBuffer,
  inspectWordDocumentFile,
  saveGeneratedWordDocument,
  streamToBuffer,
  transformWordDocumentFile,
};
