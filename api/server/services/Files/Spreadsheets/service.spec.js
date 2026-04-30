jest.mock('uuid', () => ({
  v4: jest.fn(() => 'generated-file-id'),
}));

jest.mock('librechat-data-provider', () => ({
  EModelEndpoint: {
    assistants: 'assistants',
    azureAssistants: 'azureAssistants',
  },
  FileSources: {
    s3: 's3',
    local: 'local',
    openai: 'openai',
    azure: 'azure',
  },
  FileContext: {
    message_attachment: 'message_attachment',
  },
  checkOpenAIStorage: jest.fn(() => false),
}));

jest.mock('~/server/controllers/assistants/helpers', () => ({
  getOpenAIClient: jest.fn(),
}));

jest.mock('~/server/services/Files/strategies', () => ({
  getStrategyFunctions: jest.fn(),
}));

jest.mock('~/server/utils/getFileStrategy', () => ({
  getFileStrategy: jest.fn(),
}));

jest.mock('~/models', () => ({
  createFile: jest.fn(),
}));

jest.mock('~/server/services/Files/Spreadsheets/transform', () => ({
  inspectSpreadsheetBuffer: jest.fn(),
  transformSpreadsheetBuffer: jest.fn(),
}));

jest.mock('./workerClient', () => ({
  SpreadsheetWorkerError: class SpreadsheetWorkerError extends Error {
    constructor(message, options = {}) {
      super(message);
      this.name = 'SpreadsheetWorkerError';
      this.code = options.code;
      this.status = options.status;
    }
  },
  SpreadsheetWorkerUnavailableError: class SpreadsheetWorkerUnavailableError extends Error {
    constructor(message, options = {}) {
      super(message);
      this.name = 'SpreadsheetWorkerUnavailableError';
      this.code = options.code;
      this.status = options.status;
    }
  },
  inspectSpreadsheetWithWorker: jest.fn(),
  shouldFallbackToJs: jest.fn(() => false),
  shouldUseSpreadsheetWorker: jest.fn(() => false),
  transformSpreadsheetWithWorker: jest.fn(),
}));

const db = require('~/models');
const { getStrategyFunctions } = require('~/server/services/Files/strategies');
const { getFileStrategy } = require('~/server/utils/getFileStrategy');
const {
  inspectSpreadsheetBuffer,
  transformSpreadsheetBuffer,
} = require('~/server/services/Files/Spreadsheets/transform');
const workerClient = require('./workerClient');
const {
  getSpreadsheetFileBuffer,
  inspectSpreadsheetFile,
  saveGeneratedSpreadsheet,
  transformSpreadsheetFile,
} = require('./service');

describe('Spreadsheet file service', () => {
  const req = {
    user: { id: 'user-1' },
    config: { fileStrategy: 's3' },
    body: {},
  };
  const res = {};
  const sourceFile = {
    filename: 'runway.xlsx',
    filepath: 'uploads/runway.xlsx',
    conversationId: 'conversation-1',
    messageId: 'message-1',
  };

  beforeEach(() => {
    jest.clearAllMocks();
    getStrategyFunctions.mockImplementation((strategy) => {
      if (strategy === 's3') {
        return {
          getDownloadStream: jest.fn(async () => Readable.from([Buffer.from('input-buffer')])),
          saveBuffer: jest.fn(async ({ fileName }) => `uploads/${fileName}`),
        };
      }

      if (strategy === 'local-export') {
        return {
          saveBuffer: jest.fn(async ({ fileName }) => `exports/${fileName}`),
        };
      }

      return {};
    });
    getFileStrategy.mockReturnValue('local-export');
    db.createFile.mockImplementation(async (payload) => payload);
    inspectSpreadsheetBuffer.mockResolvedValue({ engine: 'js', sheetCount: 1 });
    transformSpreadsheetBuffer.mockResolvedValue({
      buffer: Buffer.from('js-output'),
      bytes: 9,
      filename: 'runway-transformed.xlsx',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      summary: { engine: 'js' },
    });
    workerClient.inspectSpreadsheetWithWorker.mockResolvedValue({
      engine: 'python_worker',
      sheetCount: 1,
    });
    workerClient.transformSpreadsheetWithWorker.mockResolvedValue({
      buffer: Buffer.from('python-output'),
      bufferBase64: Buffer.from('python-output').toString('base64'),
      bytes: 13,
      filename: 'runway-transformed.xlsx',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      summary: { engine: 'python_worker' },
    });
    workerClient.shouldUseSpreadsheetWorker.mockReturnValue(false);
    workerClient.shouldFallbackToJs.mockReturnValue(false);
  });

  it('streams the source spreadsheet through the configured strategy', async () => {
    const buffer = await getSpreadsheetFileBuffer(req, res, sourceFile);

    expect(buffer.toString()).toBe('input-buffer');
    expect(getStrategyFunctions).toHaveBeenCalledWith('s3');
  });

  it('routes inspection to the Python worker for supported xlsx files when enabled', async () => {
    workerClient.shouldUseSpreadsheetWorker.mockReturnValue(true);

    const result = await inspectSpreadsheetFile({
      req,
      res,
      sourceFile,
      maxPreviewRows: 7,
    });

    expect(workerClient.inspectSpreadsheetWithWorker).toHaveBeenCalledWith({
      buffer: Buffer.from('input-buffer'),
      sourceFilename: 'runway.xlsx',
      maxPreviewRows: 7,
    });
    expect(inspectSpreadsheetBuffer).not.toHaveBeenCalled();
    expect(result).toEqual({
      engine: 'python_worker',
      sheetCount: 1,
    });
  });

  it('falls back to JS inspection when the worker is not selected', async () => {
    const result = await inspectSpreadsheetFile({
      req,
      res,
      sourceFile: { ...sourceFile, filename: 'runway.csv' },
      maxPreviewRows: 4,
    });

    expect(workerClient.inspectSpreadsheetWithWorker).not.toHaveBeenCalled();
    expect(inspectSpreadsheetBuffer).toHaveBeenCalledWith({
      buffer: Buffer.from('input-buffer'),
      sourceFilename: 'runway.csv',
      maxPreviewRows: 4,
    });
    expect(result).toEqual({ engine: 'js', sheetCount: 1 });
  });

  it('throws the worker error by default when the Python-primary path rejects an operation', async () => {
    workerClient.shouldUseSpreadsheetWorker.mockReturnValue(true);
    workerClient.transformSpreadsheetWithWorker.mockRejectedValue(
      new workerClient.SpreadsheetWorkerError('unsupported', {
        code: 'UNSUPPORTED_OPERATION',
        status: 422,
      }),
    );

    await expect(
      transformSpreadsheetFile({
        req,
        res,
        sourceFile,
        removeColumns: [],
        keepColumns: [],
        redactColumns: [],
        redactionText: '[REDACTED]',
        sheetNames: [],
        outputFormat: 'xlsx',
        operations: [{ type: 'sort_rows', sheetName: 'Runway', columnName: 'Amount' }],
      }),
    ).rejects.toThrow('unsupported');

    expect(workerClient.transformSpreadsheetWithWorker).toHaveBeenCalled();
    expect(transformSpreadsheetBuffer).not.toHaveBeenCalled();
  });

  it('can still fall back to the JS transformer when fallback is explicitly enabled', async () => {
    workerClient.shouldUseSpreadsheetWorker.mockReturnValue(true);
    workerClient.shouldFallbackToJs.mockReturnValue(true);
    workerClient.transformSpreadsheetWithWorker.mockRejectedValue(
      new workerClient.SpreadsheetWorkerError('worker unavailable', {
        code: 'SPREADSHEET_WORKER_UNAVAILABLE',
        status: 503,
      }),
    );

    const result = await transformSpreadsheetFile({
      req,
      res,
      sourceFile,
      removeColumns: [],
      keepColumns: [],
      redactColumns: [],
      redactionText: '[REDACTED]',
      sheetNames: [],
      outputFormat: 'xlsx',
      operations: [{ type: 'sort_rows', sheetName: 'Runway', columnName: 'Amount' }],
    });

    expect(transformSpreadsheetBuffer).toHaveBeenCalledWith(
      expect.objectContaining({
        buffer: Buffer.from('input-buffer'),
        sourceFilename: 'runway.xlsx',
      }),
    );
    expect(result.summary).toEqual({ engine: 'js' });
  });

  it('uses the Python worker output when the transform succeeds', async () => {
    workerClient.shouldUseSpreadsheetWorker.mockReturnValue(true);

    const result = await transformSpreadsheetFile({
      req,
      res,
      sourceFile,
      removeColumns: [],
      keepColumns: ['Department'],
      redactColumns: [],
      redactionText: '[REDACTED]',
      sheetNames: ['Runway'],
      outputFormat: 'xlsx',
      operations: [{ type: 'add_column', sheetName: 'Runway', columnName: 'RunwayMonths' }],
      conversationId: 'conversation-2',
      messageId: 'message-2',
    });

    expect(transformSpreadsheetBuffer).not.toHaveBeenCalled();
    expect(workerClient.transformSpreadsheetWithWorker).toHaveBeenCalledWith({
      buffer: Buffer.from('input-buffer'),
      sourceFilename: 'runway.xlsx',
      removeColumns: [],
      keepColumns: ['Department'],
      redactColumns: [],
      redactionText: '[REDACTED]',
      sheetNames: ['Runway'],
      outputFormat: 'xlsx',
      operations: [{ type: 'add_column', sheetName: 'Runway', columnName: 'RunwayMonths' }],
    });
    expect(result.summary).toEqual({ engine: 'python_worker' });
    expect(result.file).toMatchObject({
      filename: 'runway-transformed.xlsx',
      source: 'local-export',
      conversationId: 'conversation-2',
      messageId: 'message-2',
    });
  });

  it('saves generated spreadsheets through the resolved export strategy', async () => {
    const file = await saveGeneratedSpreadsheet({
      req,
      sourceFile,
      generated: {
        buffer: Buffer.from('generated'),
        bytes: 9,
        filename: 'generated.xlsx',
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      },
      conversationId: 'conversation-9',
      messageId: 'message-9',
    });

    expect(db.createFile).toHaveBeenCalledWith(
      expect.objectContaining({
        file_id: 'generated-file-id',
        filename: 'generated.xlsx',
        source: 'local-export',
        conversationId: 'conversation-9',
        messageId: 'message-9',
      }),
      true,
    );
    expect(file.filepath).toBe('exports/generated-file-id__generated.xlsx');
  });
});
const { Readable } = require('stream');
