const {
  isSpreadsheetTransformable,
  transformSpreadsheetBuffer,
} = require('./transform');
const { read, utils, write } = require('xlsx');

describe('Spreadsheet transform service', () => {
  it('removes and redacts columns from an xlsx workbook', async () => {
    const workbook = utils.book_new();
    const worksheet = utils.aoa_to_sheet([
      ['Employee', 'Email', 'RunwayMonths', 'Department'],
      ['Alice', 'alice@example.com', 18, 'Finance'],
      ['Bob', 'bob@example.com', 24, 'Operations'],
    ]);
    utils.book_append_sheet(workbook, worksheet, 'Runway');

    const inputBuffer = Buffer.from(write(workbook, { type: 'buffer', bookType: 'xlsx' }));
    const result = await transformSpreadsheetBuffer({
      buffer: inputBuffer,
      sourceFilename: 'runway.xlsx',
      removeColumns: ['RunwayMonths'],
      redactColumns: ['Email'],
      outputFormat: 'xlsx',
    });

    const outputWorkbook = read(result.buffer, { type: 'buffer' });
    const outputRows = utils.sheet_to_json(outputWorkbook.Sheets.Runway, {
      header: 1,
      raw: false,
      defval: '',
    });

    expect(outputRows).toEqual([
      ['Employee', 'Email', 'Department'],
      ['Alice', '[REDACTED]', 'Finance'],
      ['Bob', '[REDACTED]', 'Operations'],
    ]);
    expect(result.filename).toBe('runway-transformed.xlsx');
    expect(result.summary.matchedColumns.remove).toContain('RunwayMonths');
    expect(result.summary.matchedColumns.redact).toContain('Email');
  });

  it('writes csv output for a single-sheet spreadsheet transform', async () => {
    const inputBuffer = Buffer.from(
      ['Employee,Salary,Department', 'Alice,150000,Finance', 'Bob,120000,Operations'].join('\n'),
      'utf8',
    );

    const result = await transformSpreadsheetBuffer({
      buffer: inputBuffer,
      sourceFilename: 'salaries.csv',
      keepColumns: ['Employee', 'Department'],
      outputFormat: 'csv',
    });

    expect(result.mimeType).toBe('text/csv');
    expect(result.filename).toBe('salaries-transformed.csv');
    expect(result.buffer.toString('utf8').trim()).toBe(
      ['Employee,Department', 'Alice,Finance', 'Bob,Operations'].join('\n'),
    );
  });

  it('recognizes supported spreadsheet MIME types', () => {
    expect(isSpreadsheetTransformable('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')).toBe(true);
    expect(isSpreadsheetTransformable('text/csv')).toBe(true);
    expect(isSpreadsheetTransformable('application/pdf')).toBe(false);
  });
});
