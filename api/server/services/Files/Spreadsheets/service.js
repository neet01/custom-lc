const { Readable } = require('stream');
const { v4: uuidv4 } = require('uuid');
const { logger } = require('@librechat/data-schemas');
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
const {
  SpreadsheetWorkerError,
  isSpreadsheetWorkerEnabled,
  inspectSpreadsheetWithWorker,
  shouldFallbackToJs,
  shouldUseSpreadsheetWorker,
  supportsPythonSpreadsheetWorker,
  transformSpreadsheetWithWorker,
} = require('~/server/services/Files/Spreadsheets/workerClient');

async function streamToBuffer(stream) {
  const chunks = [];
  for await (const chunk of stream) {
    chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
  }
  return Buffer.concat(chunks);
}

function resolveDocumentFileSource(req, file) {
  return file?.source || req?.config?.fileStrategy || FileSources.s3;
}

async function getSpreadsheetFileBuffer(req, res, file) {
  const fileSource = resolveDocumentFileSource(req, file);
  const { getDownloadStream } = getStrategyFunctions(fileSource);
  if (!getDownloadStream) {
    throw new Error(`No download stream method implemented for file source: ${fileSource}`);
  }

  if (checkOpenAIStorage(fileSource)) {
    req.body = { ...req.body, model: file.model };
    const endpointMap = {
      [FileSources.openai]: EModelEndpoint.assistants,
      [FileSources.azure]: EModelEndpoint.azureAssistants,
    };
    const { openai } = await getOpenAIClient({
      req,
      res,
      overrideEndpoint: endpointMap[fileSource],
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
    FileSources.s3,
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

  throw new Error('No valid file strategy available for generated spreadsheet exports');
}

async function saveGeneratedSpreadsheet({
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

async function inspectSpreadsheetFile({
  req,
  res,
  sourceFile,
  maxPreviewRows,
}) {
  const sourceBuffer = await getSpreadsheetFileBuffer(req, res, sourceFile);
  const workerSelected = shouldUseSpreadsheetWorker(sourceFile.filename);
  logger.info('[SpreadsheetService] Inspect routing decision', {
    fileId: sourceFile?.file_id,
    filename: sourceFile?.filename,
    workerEnabled: isSpreadsheetWorkerEnabled(),
    fileSupportedByWorker: supportsPythonSpreadsheetWorker(sourceFile?.filename),
    workerSelected,
    maxPreviewRows,
  });

  if (workerSelected) {
    try {
      logger.info('[SpreadsheetService] Dispatching spreadsheet inspection to Python worker', {
        fileId: sourceFile?.file_id,
        filename: sourceFile?.filename,
      });
      return await inspectSpreadsheetWithWorker({
        buffer: sourceBuffer,
        sourceFilename: sourceFile.filename,
        maxPreviewRows,
      });
    } catch (error) {
      const canFallback = shouldFallbackToJs() && error instanceof SpreadsheetWorkerError;
      logger.warn('[SpreadsheetService] Python worker inspection failed', {
        fileId: sourceFile?.file_id,
        filename: sourceFile?.filename,
        canFallback,
        code: error?.code,
        status: error?.status,
        error: error?.message,
      });
      if (!canFallback) {
        throw error;
      }
    }
  }

  logger.info('[SpreadsheetService] Using JavaScript inspection path', {
    fileId: sourceFile?.file_id,
    filename: sourceFile?.filename,
  });

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
  operations,
  conversationId,
  messageId,
}) {
  const sourceBuffer = await getSpreadsheetFileBuffer(req, res, sourceFile);
  let generated;
  const workerSelected = shouldUseSpreadsheetWorker(sourceFile.filename);

  logger.info('[SpreadsheetService] Transform routing decision', {
    fileId: sourceFile?.file_id,
    filename: sourceFile?.filename,
    workerEnabled: isSpreadsheetWorkerEnabled(),
    fileSupportedByWorker: supportsPythonSpreadsheetWorker(sourceFile?.filename),
    workerSelected,
    outputFormat,
    operationCount: Array.isArray(operations) ? operations.length : 0,
    operationTypes: Array.isArray(operations)
      ? operations.map((operation) => operation?.type).filter(Boolean)
      : [],
  });

  if (workerSelected) {
    try {
      logger.info('[SpreadsheetService] Dispatching spreadsheet transform to Python worker', {
        fileId: sourceFile?.file_id,
        filename: sourceFile?.filename,
      });
      generated = await transformSpreadsheetWithWorker({
        buffer: sourceBuffer,
        sourceFilename: sourceFile.filename,
        removeColumns,
        keepColumns,
        redactColumns,
        redactionText,
        sheetNames,
        outputFormat,
        operations,
      });
    } catch (error) {
      const canFallback = shouldFallbackToJs() && error instanceof SpreadsheetWorkerError;
      logger.warn('[SpreadsheetService] Python worker transform failed', {
        fileId: sourceFile?.file_id,
        filename: sourceFile?.filename,
        canFallback,
        code: error?.code,
        status: error?.status,
        error: error?.message,
      });
      if (!canFallback) {
        throw error;
      }
    }
  }

  if (!generated) {
    logger.info('[SpreadsheetService] Using JavaScript transform path', {
      fileId: sourceFile?.file_id,
      filename: sourceFile?.filename,
      outputFormat,
    });
    generated = await transformSpreadsheetBuffer({
      buffer: sourceBuffer,
      sourceFilename: sourceFile.filename,
      removeColumns,
      keepColumns,
      redactColumns,
      redactionText,
      sheetNames,
      outputFormat,
      operations,
    });
  }

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
