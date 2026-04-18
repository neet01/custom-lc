const { Readable } = require('stream');
const { FileSources, FileContext, SystemRoles } = require('librechat-data-provider');

jest.mock('~/models', () => ({
  createFile: jest.fn(async (data) => ({ ...data, object: 'file' })),
  getFiles: jest.fn(),
  getAgent: jest.fn(),
}));

jest.mock('~/server/services/Files/process', () => ({
  processDeleteRequest: jest.fn(),
  filterFile: jest.fn(),
  processFileUpload: jest.fn(),
  processAgentFileUpload: jest.fn(),
}));

jest.mock('~/server/services/Files/strategies', () => ({
  getStrategyFunctions: jest.fn(),
}));

jest.mock('~/server/controllers/assistants/helpers', () => ({
  getOpenAIClient: jest.fn(),
}));

jest.mock('~/server/services/Tools/credentials', () => ({
  loadAuthValues: jest.fn(),
}));

jest.mock('~/server/services/PermissionService', () => ({
  checkPermission: jest.fn(),
}));

jest.mock('~/server/services/Files', () => ({
  hasAccessToFilesViaAgent: jest.fn(),
}));

jest.mock('~/server/middleware/accessResources/fileAccess', () => ({
  fileAccess: (req, _res, next) => {
    req.fileAccess = { file: req.app.locals.testFile };
    next();
  },
}));

jest.mock('~/cache', () => ({
  getLogStores: jest.fn(() => ({
    get: jest.fn(),
    set: jest.fn(),
  })),
}));

jest.mock('@librechat/api', () => ({
  ...jest.requireActual('@librechat/api'),
  refreshS3FileUrls: jest.fn(),
}));

const db = require('~/models');
const { getStrategyFunctions } = require('~/server/services/Files/strategies');
const router = require('./files');

describe('Spreadsheet transform route', () => {
  let handlerLayer;

  beforeEach(() => {
    jest.clearAllMocks();
    handlerLayer = router.stack.find(
      (layer) =>
        layer.route?.path === '/:file_id/transform/spreadsheet' && layer.route.methods.post,
    );
  });

  async function invokeRoute({ testFile, body }) {
    const req = {
      params: { file_id: 'source-file-1' },
      body,
      user: { id: 'user-123', role: SystemRoles.USER },
      config: { fileStrategy: FileSources.local },
      app: { locals: { testFile } },
    };

    const res = {
      status: jest.fn(function status() {
        return this;
      }),
      json: jest.fn(function json() {
        return this;
      }),
    };

    const [fileAccessMiddleware, routeHandler] = handlerLayer.route.stack.map((layer) => layer.handle);
    await new Promise((resolve, reject) =>
      fileAccessMiddleware(req, res, (error) => (error ? reject(error) : resolve())),
    );
    await routeHandler(req, res);

    return { req, res };
  }

  it('creates a transformed spreadsheet file record and returns it', async () => {
    const csvData = Buffer.from(
      ['Employee,Salary,Department', 'Alice,150000,Finance', 'Bob,120000,Operations'].join('\n'),
      'utf8',
    );
    const saveBuffer = jest.fn().mockResolvedValue('/uploads/user-123/runway-transformed.csv');

    getStrategyFunctions.mockImplementation(() => ({
      getDownloadStream: jest.fn().mockResolvedValue(Readable.from([csvData])),
      saveBuffer,
    }));

    const { res } = await invokeRoute({
      testFile: {
        file_id: 'source-file-1',
        filename: 'runway.csv',
        filepath: '/uploads/user-123/runway.csv',
        type: 'text/csv',
        source: FileSources.local,
        conversationId: 'convo-1',
        messageId: 'msg-1',
      },
      body: {
        removeColumns: ['Salary'],
        outputFormat: 'csv',
      },
    });

    expect(res.status).toHaveBeenCalledWith(200);
    const payload = res.json.mock.calls[0][0];
    expect(payload.message).toBe('Spreadsheet transformed successfully');
    expect(payload.file.filename).toBe('runway-transformed.csv');
    expect(payload.file.context).toBe(FileContext.message_attachment);
    expect(payload.summary.matchedColumns.remove).toContain('Salary');
    expect(saveBuffer).toHaveBeenCalledWith(
      expect.objectContaining({
        userId: 'user-123',
        basePath: 'uploads',
      }),
    );
    expect(db.createFile).toHaveBeenCalledWith(
      expect.objectContaining({
        user: 'user-123',
        filename: 'runway-transformed.csv',
        type: 'text/csv',
        conversationId: 'convo-1',
        messageId: 'msg-1',
      }),
      true,
    );
  });

  it('rejects non-spreadsheet source files', async () => {
    const { res } = await invokeRoute({
      testFile: {
        file_id: 'source-file-1',
        filename: 'notes.txt',
        filepath: '/uploads/user-123/notes.txt',
        type: 'text/plain',
        source: FileSources.local,
      },
      body: {
        removeColumns: ['Salary'],
      },
    });

    expect(res.status).toHaveBeenCalledWith(400);
    expect(res.json.mock.calls[0][0].message).toContain('not a supported spreadsheet type');
  });
});
