jest.mock('~/models', () => ({
  getDocumentBySourceFileId: jest.fn(),
  createDocument: jest.fn(),
  createDocumentVersion: jest.fn(),
  createDocumentJob: jest.fn(),
  updateDocument: jest.fn(),
}));

const db = require('~/models');
const {
  isIndexableDocumentType,
  registerDocumentUpload,
} = require('./register');

describe('document upload registration', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  it('ignores non-indexable file types', async () => {
    expect(isIndexableDocumentType('image/png')).toBe(false);

    const result = await registerDocumentUpload({
      userId: 'user-1',
      file: {
        file_id: 'file-1',
        filename: 'diagram.png',
        type: 'image/png',
        bytes: 10,
        source: 's3',
      },
    });

    expect(result).toBeNull();
    expect(db.createDocument).not.toHaveBeenCalled();
  });

  it('creates document, version, and job records for indexable files', async () => {
    db.getDocumentBySourceFileId.mockResolvedValue(null);
    db.createDocument.mockResolvedValue({ _id: 'doc-1' });
    db.createDocumentVersion.mockResolvedValue({ _id: 'ver-1' });
    db.createDocumentJob.mockResolvedValue({ _id: 'job-1' });
    db.updateDocument.mockResolvedValue({ _id: 'doc-1' });

    const result = await registerDocumentUpload({
      userId: 'user-1',
      conversationId: 'conversation-1',
      messageId: 'message-1',
      file: {
        file_id: 'file-1',
        filename: 'spec.pdf',
        filepath: '/tmp/spec.pdf',
        type: 'application/pdf',
        bytes: 4096,
        source: 's3',
        context: 'message_attachment',
      },
    });

    expect(result).toEqual({ _id: 'doc-1' });
    expect(db.createDocument).toHaveBeenCalledWith(
      expect.objectContaining({
        user: 'user-1',
        sourceFileId: 'file-1',
        filename: 'spec.pdf',
        mimeType: 'application/pdf',
        status: 'pending',
        extractionKind: 'none',
      }),
    );
    expect(db.createDocumentVersion).toHaveBeenCalledWith(
      expect.objectContaining({
        documentId: 'doc-1',
        versionNumber: 1,
        sourceFileId: 'file-1',
      }),
    );
    expect(db.createDocumentJob).toHaveBeenCalledWith(
      expect.objectContaining({
        documentId: 'doc-1',
        documentVersionId: 'ver-1',
        jobType: 'extract',
        status: 'pending',
      }),
    );
  });

  it('uses chunk as the next job for text-backed uploads', async () => {
    db.getDocumentBySourceFileId.mockResolvedValue(null);
    db.createDocument.mockResolvedValue({ _id: 'doc-1' });
    db.createDocumentVersion.mockResolvedValue({ _id: 'ver-1' });
    db.createDocumentJob.mockResolvedValue({ _id: 'job-1' });
    db.updateDocument.mockResolvedValue({ _id: 'doc-1' });

    await registerDocumentUpload({
      userId: 'user-1',
      file: {
        file_id: 'file-1',
        filename: 'fallback.txt',
        filepath: '/tmp/fallback.txt',
        type: 'text/plain',
        bytes: 128,
        text: 'hello world',
        source: 'text',
        context: 'message_attachment',
      },
    });

    expect(db.createDocument).toHaveBeenCalledWith(
      expect.objectContaining({
        extractionKind: 'text',
      }),
    );
    expect(db.createDocumentVersion).toHaveBeenCalledWith(
      expect.objectContaining({
        textLength: Buffer.byteLength('hello world', 'utf8'),
        extractionKind: 'text',
      }),
    );
    expect(db.createDocumentJob).toHaveBeenCalledWith(
      expect.objectContaining({
        jobType: 'chunk',
      }),
    );
  });
});
