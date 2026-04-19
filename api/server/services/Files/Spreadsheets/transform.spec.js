const {
  isSpreadsheetTransformable,
  transformSpreadsheetBuffer,
} = require('./transform');
const { read, utils, write } = require('xlsx');

describe('Spreadsheet transform service', () => {
  it('removes and redacts columns from an xlsx workbook while preserving sheet metadata', async () => {
    const workbook = utils.book_new();
    const worksheet = utils.aoa_to_sheet([
      ['Merged Header', '', '', ''],
      ['Employee', 'Email', 'RunwayMonths', 'Department'],
      ['Alice', 'alice@example.com', 18, 'Finance'],
      ['Bob', 'bob@example.com', 24, 'Operations'],
    ]);
    worksheet['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 3 } }];
    worksheet['!cols'] = [{ wch: 18 }, { wch: 28 }, { wch: 16 }, { wch: 20 }];
    worksheet['!autofilter'] = { ref: 'A2:D4' };
    utils.book_append_sheet(workbook, worksheet, 'Runway');

    const untouchedSheet = utils.aoa_to_sheet([
      ['Scenario', 'Cash'],
      ['Base', 10_000_000],
    ]);
    utils.book_append_sheet(workbook, untouchedSheet, 'Scenarios');

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
      ['Merged Header', '', ''],
      ['Employee', 'Email', 'Department'],
      ['Alice', '[REDACTED]', 'Finance'],
      ['Bob', '[REDACTED]', 'Operations'],
    ]);
    expect(outputWorkbook.Sheets.Runway['!merges']).toEqual([{ s: { r: 0, c: 0 }, e: { r: 0, c: 2 } }]);
    expect(outputWorkbook.Sheets.Runway['!autofilter'].ref).toBe('A2:C4');
    expect(outputWorkbook.Sheets.Scenarios['A2'].v).toBe('Base');
    expect(result.filename).toBe('runway-transformed.xlsx');
    expect(result.summary.matchedColumns.remove).toContain('RunwayMonths');
    expect(result.summary.matchedColumns.redact).toContain('Email');
  });

  it('keeps selected columns in csv output for a single sheet', async () => {
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

  it('redacts values in place without removing workbook structure when no columns are dropped', async () => {
    const workbook = utils.book_new();
    const worksheet = utils.aoa_to_sheet([
      ['Employee', 'Email', 'Department'],
      ['Alice', 'alice@example.com', 'Finance'],
    ]);
    worksheet['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }];
    utils.book_append_sheet(workbook, worksheet, 'Runway');

    const inputBuffer = Buffer.from(write(workbook, { type: 'buffer', bookType: 'xlsx' }));
    const result = await transformSpreadsheetBuffer({
      buffer: inputBuffer,
      sourceFilename: 'runway.xlsx',
      redactColumns: ['Email'],
      outputFormat: 'xlsx',
    });

    const outputWorkbook = read(result.buffer, { type: 'buffer' });
    expect(outputWorkbook.Sheets.Runway['B2'].v).toBe('[REDACTED]');
    expect(outputWorkbook.Sheets.Runway['!merges']).toEqual([{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }]);
  });

  it('recognizes supported spreadsheet MIME types', () => {
    expect(isSpreadsheetTransformable('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')).toBe(true);
    expect(isSpreadsheetTransformable('text/csv')).toBe(true);
    expect(isSpreadsheetTransformable('application/pdf')).toBe(false);
  });
});
