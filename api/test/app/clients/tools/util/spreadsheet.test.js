jest.mock('~/models', () => ({
  getConvoFiles: jest.fn(),
  getFiles: jest.fn(),
}));

jest.mock('~/server/services/Files/permissions', () => ({
  filterFilesByAgentAccess: jest.fn((options) => Promise.resolve(options.files)),
}));

jest.mock('~/server/services/Files/Spreadsheets/service', () => ({
  inspectSpreadsheetFile: jest.fn(),
  transformSpreadsheetFile: jest.fn(),
}));

const { getConvoFiles, getFiles } = require('~/models');
const { inspectSpreadsheetFile, transformSpreadsheetFile } = require(
  '~/server/services/Files/Spreadsheets/service',
);
const { primeFiles, createSpreadsheetTool, SPREADSHEET_TOOL_NAME } = require(
  '~/app/clients/tools/util/spreadsheet',
);

describe('spreadsheet tool', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  describe('primeFiles', () => {
    it('collects spreadsheet files from the request and conversation', async () => {
      getConvoFiles.mockResolvedValue(['file-2', 'file-3']);
      getFiles.mockResolvedValue([
        {
          file_id: 'file-2',
          filename: 'budget.xlsx',
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
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
                filename: 'runway.csv',
                type: 'text/csv',
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
      expect(result.toolContext).toContain(SPREADSHEET_TOOL_NAME);
      expect(result.toolContext).toContain('runway.csv');
      expect(result.toolContext).toContain('budget.xlsx');
    });
  });

  describe('createSpreadsheetTool', () => {
    const req = {
      body: {
        conversationId: 'convo-1',
      },
      user: {
        id: 'user-1',
      },
    };

    const spreadsheetFile = {
      file_id: 'file-1',
      filename: 'runway.xlsx',
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    };

    it('returns workbook inspection details', async () => {
      inspectSpreadsheetFile.mockResolvedValue({
        filename: 'runway.xlsx',
        sheetCount: 1,
        sheets: [
          {
            sheetName: 'Runway',
            rowCount: 2,
            columnCount: 3,
            columns: ['Cash', 'Burn', 'Month'],
            previewRows: [{ Cash: '1000000', Burn: '100000', Month: 'January' }],
          },
        ],
      });

      const spreadsheetTool = await createSpreadsheetTool({
        req,
        res: {},
        files: [spreadsheetFile],
      });

      const result = await spreadsheetTool.func({
        action: 'inspect',
        file_id: 'file-1',
      });

      expect(result[0]).toContain('Workbook "runway.xlsx" contains 1 sheet(s).');
      expect(result[1][SPREADSHEET_TOOL_NAME].inspection.sheetCount).toBe(1);
      expect(inspectSpreadsheetFile).toHaveBeenCalledWith(
        expect.objectContaining({
          sourceFile: spreadsheetFile,
        }),
      );
    });

    it('returns a generated spreadsheet file artifact', async () => {
      transformSpreadsheetFile.mockResolvedValue({
        file: {
          file_id: 'file-2',
          filename: 'runway-transformed.xlsx',
          filepath: '/uploads/runway-transformed.xlsx',
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        },
        summary: {
          matchedColumns: {
            remove: ['Salary'],
          },
        },
      });

      const spreadsheetTool = await createSpreadsheetTool({
        req,
        res: {},
        files: [spreadsheetFile],
      });

      const result = await spreadsheetTool.func({
        action: 'transform',
        removeColumns: ['Salary'],
      });

      expect(result[0]).toContain('runway-transformed.xlsx');
      expect(result[1].files).toHaveLength(1);
      expect(result[1].files[0].file_id).toBe('file-2');
      expect(result[1][SPREADSHEET_TOOL_NAME].summary.matchedColumns.remove).toEqual(['Salary']);
      expect(transformSpreadsheetFile).toHaveBeenCalledWith(
        expect.objectContaining({
          sourceFile: spreadsheetFile,
          removeColumns: ['Salary'],
          conversationId: 'convo-1',
        }),
      );
    });
  });
});
