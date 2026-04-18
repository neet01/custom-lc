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
  inspectSpreadsheetBuffer,
  transformSpreadsheetBuffer,
} = require('~/server/services/Files/Spreadsheets/transform');

async function streamToBuffer(stream) {
  const chunks = [];
  for await (const chunk of stream) {
    chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
  }
  return Buffer.concat(chunks);
}

async function getSpreadsheetFileBuffer(req, res, file) {
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

async function saveGeneratedSpreadsheet({
  req,
  sourceFile,
  generated,
  conversationId,
  messageId,
}) {
  const fileStrategy = getFileStrategy(req.config, { context: FileContext.message_attachment });
  const { saveBuffer } = getStrategyFunctions(fileStrategy);
  if (!saveBuffer) {
    throw new Error(
      `File strategy "${fileStrategy}" does not support generated spreadsheet exports`,
    );
  }

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

async function inspectSpreadsheetFile({
  req,
  res,
  sourceFile,
  maxPreviewRows,
}) {
  const sourceBuffer = await getSpreadsheetFileBuffer(req, res, sourceFile);
  return inspectSpreadsheetBuffer({
    buffer: sourceBuffer,
    sourceFilename: sourceFile.filename,
    maxPreviewRows,
  });
}

async function transformSpreadsheetFile({
  req,
  res,
  sourceFile,
  removeColumns,
  keepColumns,
  redactColumns,
  redactionText,
  sheetNames,
  outputFormat,
  conversationId,
  messageId,
}) {
  const sourceBuffer = await getSpreadsheetFileBuffer(req, res, sourceFile);
  const generated = await transformSpreadsheetBuffer({
    buffer: sourceBuffer,
    sourceFilename: sourceFile.filename,
    removeColumns,
    keepColumns,
    redactColumns,
    redactionText,
    sheetNames,
    outputFormat,
  });

  const file = await saveGeneratedSpreadsheet({
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
  streamToBuffer,
  getSpreadsheetFileBuffer,
  saveGeneratedSpreadsheet,
  inspectSpreadsheetFile,
  transformSpreadsheetFile,
};
