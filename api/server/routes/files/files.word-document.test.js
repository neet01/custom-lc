const { Readable } = require('stream');
const { FileSources, FileContext, SystemRoles } = require('librechat-data-provider');
const JSZip = require('jszip');

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

async function createDocxBufferFromText(text) {
  const zip = new JSZip();
  const bodyXml = String(text)
    .split('\n')
    .map((paragraph) =>
      paragraph
        ? `<w:p><w:r><w:t xml:space="preserve">${paragraph}</w:t></w:r></w:p>`
        : '<w:p/>',
    )
    .join('');

  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`,
  );
  zip.folder('_rels').file(
    '.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`,
  );
  zip.folder('word').folder('_rels').file(
    'document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`,
  );
  zip.folder('word').file(
    'document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${bodyXml}
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
  </w:body>
</w:document>`,
  );

  return zip.generateAsync({ type: 'nodebuffer' });
}

describe('Word document transform route', () => {
  let handlerLayer;

  beforeEach(() => {
    jest.clearAllMocks();
    handlerLayer = router.stack.find(
      (layer) =>
        layer.route?.path === '/:file_id/transform/word-document' && layer.route.methods.post,
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

  it('creates a transformed Word document file record and returns it', async () => {
    const docxData = await createDocxBufferFromText('Budget memo\nRevenue is 500.');
    const saveBuffer = jest.fn().mockResolvedValue('/uploads/user-123/budget-transformed.docx');

    getStrategyFunctions.mockImplementation(() => ({
      getDownloadStream: jest.fn().mockResolvedValue(Readable.from([docxData])),
      saveBuffer,
    }));

    const { res } = await invokeRoute({
      testFile: {
        file_id: 'source-file-1',
        filename: 'budget.docx',
        filepath: '/uploads/user-123/budget.docx',
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        source: FileSources.local,
        conversationId: 'convo-1',
        messageId: 'msg-1',
      },
      body: {
        replaceText: [{ find: '500', replace: '650' }],
      },
    });

    expect(res.status).toHaveBeenCalledWith(200);
    const payload = res.json.mock.calls[0][0];
    expect(payload.message).toBe('Word document transformed successfully');
    expect(payload.file.filename).toBe('budget-transformed.docx');
    expect(payload.file.context).toBe(FileContext.message_attachment);
    expect(payload.summary.replacements[0].find).toBe('500');
    expect(saveBuffer).toHaveBeenCalledWith(
      expect.objectContaining({
        userId: 'user-123',
        basePath: 'uploads',
      }),
    );
    expect(db.createFile).toHaveBeenCalledWith(
      expect.objectContaining({
        user: 'user-123',
        filename: 'budget-transformed.docx',
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        conversationId: 'convo-1',
        messageId: 'msg-1',
      }),
      true,
    );
  });

  it('rejects non-docx source files', async () => {
    const { res } = await invokeRoute({
      testFile: {
        file_id: 'source-file-1',
        filename: 'notes.txt',
        filepath: '/uploads/user-123/notes.txt',
        type: 'text/plain',
        source: FileSources.local,
      },
      body: {
        replaceText: [{ find: 'foo', replace: 'bar' }],
      },
    });

    expect(res.status).toHaveBeenCalledWith(400);
    expect(res.json.mock.calls[0][0].message).toContain('not a supported Word document type');
  });
});
