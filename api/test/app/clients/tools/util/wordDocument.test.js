jest.mock('~/models', () => ({
  getConvoFiles: jest.fn(),
  getFiles: jest.fn(),
}));

jest.mock('~/server/services/Files/permissions', () => ({
  filterFilesByAgentAccess: jest.fn((options) => Promise.resolve(options.files)),
}));

jest.mock('~/server/services/Files/WordDocuments/service', () => ({
  inspectWordDocumentFile: jest.fn(),
  transformWordDocumentFile: jest.fn(),
}));

const { getConvoFiles, getFiles } = require('~/models');
const { inspectWordDocumentFile, transformWordDocumentFile } = require(
  '~/server/services/Files/WordDocuments/service',
);
const { primeFiles, createWordDocumentTool, WORD_DOCUMENT_TOOL_NAME } = require(
  '~/app/clients/tools/util/wordDocument',
);

describe('word document tool', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  describe('primeFiles', () => {
    it('collects Word documents from the request and conversation', async () => {
      getConvoFiles.mockResolvedValue(['file-2', 'file-3']);
      getFiles.mockResolvedValue([
        {
          file_id: 'file-2',
          filename: 'memo.docx',
          type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        },
        {
          file_id: 'file-3',
          filename: 'notes.txt',
          type: 'text/plain',
        },
      ]);

      const result = await primeFiles({
        req: {
          body: {
            conversationId: 'convo-1',
            files: [
              {
                file_id: 'file-1',
                filename: 'proposal.docx',
                type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
              },
            ],
          },
          user: {
            id: 'user-1',
            role: 'USER',
          },
        },
        agentId: 'agent-1',
      });

      expect(result.files.map((file) => file.file_id)).toEqual(['file-1', 'file-2']);
      expect(result.toolContext).toContain(WORD_DOCUMENT_TOOL_NAME);
      expect(result.toolContext).toContain('proposal.docx');
      expect(result.toolContext).toContain('memo.docx');
    });
  });

  describe('createWordDocumentTool', () => {
    const req = {
      body: {
        conversationId: 'convo-1',
      },
      user: {
        id: 'user-1',
      },
    };

    const wordFile = {
      file_id: 'file-1',
      filename: 'proposal.docx',
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    };

    it('returns document inspection details', async () => {
      inspectWordDocumentFile.mockResolvedValue({
        filename: 'proposal.docx',
        paragraphCount: 2,
        wordCount: 8,
        previewParagraphs: ['Executive summary', 'The project is on track.'],
      });

      const wordTool = await createWordDocumentTool({
        req,
        res: {},
        files: [wordFile],
      });

      const result = await wordTool.func({
        action: 'inspect',
        file_id: 'file-1',
      });

      expect(result[0]).toContain('Document "proposal.docx" contains 2 paragraph(s)');
      expect(result[1][WORD_DOCUMENT_TOOL_NAME].inspection.paragraphCount).toBe(2);
      expect(inspectWordDocumentFile).toHaveBeenCalledWith(
        expect.objectContaining({
          sourceFile: wordFile,
        }),
      );
    });

    it('returns a generated Word document file artifact', async () => {
      transformWordDocumentFile.mockResolvedValue({
        file: {
          file_id: 'file-2',
          filename: 'proposal-redacted.docx',
          filepath: '/uploads/proposal-redacted.docx',
          type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        },
        summary: {
          redactions: [{ phrase: 'Internal Only', occurrences: 1 }],
        },
      });

      const wordTool = await createWordDocumentTool({
        req,
        res: {},
        files: [wordFile],
      });

      const result = await wordTool.func({
        action: 'transform',
        redactPhrases: ['Internal Only'],
        outputFilename: 'proposal-redacted.docx',
      });

      expect(result[0]).toContain('proposal-redacted.docx');
      expect(result[1].files).toHaveLength(1);
      expect(result[1].files[0].file_id).toBe('file-2');
      expect(result[1][WORD_DOCUMENT_TOOL_NAME].summary.redactions[0].phrase).toBe(
        'Internal Only',
      );
      expect(transformWordDocumentFile).toHaveBeenCalledWith(
        expect.objectContaining({
          sourceFile: wordFile,
          redactPhrases: ['Internal Only'],
          outputFilename: 'proposal-redacted.docx',
          conversationId: 'convo-1',
        }),
      );
    });

    it('makes a generated Word document immediately reusable within the same tool instance', async () => {
      transformWordDocumentFile
        .mockResolvedValueOnce({
          file: {
            file_id: 'file-2',
            filename: 'proposal-redacted.docx',
            filepath: '/uploads/proposal-redacted.docx',
            type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
          },
          summary: {
            redactions: [{ phrase: 'Internal Only', occurrences: 1 }],
          },
        })
        .mockResolvedValueOnce({
          file: {
            file_id: 'file-3',
            filename: 'proposal-redacted-v2.docx',
            filepath: '/uploads/proposal-redacted-v2.docx',
            type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
          },
          summary: {
            replacements: [{ find: 'Confidential', replace: 'Restricted' }],
          },
        });

      const wordTool = await createWordDocumentTool({
        req,
        res: {},
        files: [wordFile],
      });

      await wordTool.func({
        action: 'transform',
        file_id: 'file-1',
        redactPhrases: ['Internal Only'],
        outputFilename: 'proposal-redacted.docx',
      });

      const secondResult = await wordTool.func({
        action: 'transform',
        file_id: 'file-2',
        replaceText: [{ find: 'Confidential', replace: 'Restricted' }],
        outputFilename: 'proposal-redacted-v2.docx',
      });

      expect(transformWordDocumentFile).toHaveBeenNthCalledWith(
        2,
        expect.objectContaining({
          sourceFile: expect.objectContaining({
            file_id: 'file-2',
            filename: 'proposal-redacted.docx',
          }),
          outputFilename: 'proposal-redacted-v2.docx',
        }),
      );
      expect(req.body.files).toEqual(
        expect.arrayContaining([
          expect.objectContaining({ file_id: 'file-2', filename: 'proposal-redacted.docx' }),
        ]),
      );
      expect(secondResult[1].files[0].file_id).toBe('file-3');
    });
  });
});
