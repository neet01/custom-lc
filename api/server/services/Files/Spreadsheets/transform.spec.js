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

  it('supports advanced row, cell, calculation, formula, and ordering operations', async () => {
    const workbook = utils.book_new();
    const worksheet = utils.aoa_to_sheet([
      ['Employee', 'Revenue', 'Expense', 'Department'],
      ['Alice', 200, 50, 'Sales'],
      ['Bob', 150, 90, 'Sales'],
    ]);
    utils.book_append_sheet(workbook, worksheet, 'Pipeline');

    const inputBuffer = Buffer.from(write(workbook, { type: 'buffer', bookType: 'xlsx' }));
    const result = await transformSpreadsheetBuffer({
      buffer: inputBuffer,
      sourceFilename: 'pipeline.xlsx',
      outputFormat: 'xlsx',
      operations: [
        {
          type: 'add_column',
          sheetName: 'Pipeline',
          columnName: 'Net',
          expression: '{{Revenue}} - {{Expense}}',
        },
        {
          type: 'add_column',
          sheetName: 'Pipeline',
          columnName: 'NetPct',
          expression: 'round(({{Revenue}} - {{Expense}}) / {{Revenue}}, 2)',
          formula: '=ROUND(({{Revenue}}-{{Expense}})/{{Revenue}}, 2)',
        },
        {
          type: 'update_cells',
          sheetName: 'Pipeline',
          columnName: 'Department',
          rowMatch: { Employee: 'Bob' },
          value: 'Operations',
        },
        {
          type: 'add_row',
          sheetName: 'Pipeline',
          values: {
            Employee: 'Cara',
            Revenue: 400,
            Expense: 75,
            Department: 'Sales',
          },
        },
        {
          type: 'update_cells',
          sheetName: 'Pipeline',
          columnName: 'Net',
          rowMatch: { Employee: 'Cara' },
          expression: '{{Revenue}} - {{Expense}}',
        },
        {
          type: 'update_cells',
          sheetName: 'Pipeline',
          columnName: 'NetPct',
          rowMatch: { Employee: 'Cara' },
          expression: 'round(({{Revenue}} - {{Expense}}) / {{Revenue}}, 2)',
          formula: '=ROUND(({{Revenue}}-{{Expense}})/{{Revenue}}, 2)',
        },
        {
          type: 'sort_rows',
          sheetName: 'Pipeline',
          columnName: 'Net',
          direction: 'desc',
          numeric: true,
        },
        {
          type: 'reorder_rows',
          sheetName: 'Pipeline',
          orderedRowNumbers: [2, 1],
          appendRemaining: true,
        },
      ],
    });

    const outputWorkbook = read(result.buffer, { type: 'buffer' });
    const outputRows = utils.sheet_to_json(outputWorkbook.Sheets.Pipeline, {
      header: 1,
      raw: true,
      defval: '',
    });

    expect(outputRows).toEqual([
      ['Employee', 'Revenue', 'Expense', 'Department', 'Net', 'NetPct'],
      ['Alice', 200, 50, 'Sales', 150, 0.75],
      ['Cara', 400, 75, 'Sales', 325, 0.81],
      ['Bob', 150, 90, 'Operations', 60, 0.4],
    ]);
    expect(outputWorkbook.Sheets.Pipeline.F2.f).toBe('ROUND((B2-C2)/B2, 2)');
    expect(outputWorkbook.Sheets.Pipeline.F3.f).toBe('ROUND((B3-C3)/B3, 2)');
    expect(result.summary.operationsApplied.map((operation) => operation.type)).toContain('add_column');
    expect(result.summary.operationsApplied.map((operation) => operation.type)).toContain('sort_rows');
  });

  it('supports merging sheets and splitting a sheet by column value', async () => {
    const workbook = utils.book_new();
    utils.book_append_sheet(
      workbook,
      utils.aoa_to_sheet([
        ['Owner', 'Region', 'Amount'],
        ['Alice', 'East', 100],
        ['Bob', 'West', 90],
      ]),
      'Q1',
    );
    utils.book_append_sheet(
      workbook,
      utils.aoa_to_sheet([
        ['Owner', 'Region', 'Amount'],
        ['Cara', 'East', 120],
        ['Dan', 'West', 80],
      ]),
      'Q2',
    );

    const inputBuffer = Buffer.from(write(workbook, { type: 'buffer', bookType: 'xlsx' }));
    const result = await transformSpreadsheetBuffer({
      buffer: inputBuffer,
      sourceFilename: 'regional.xlsx',
      outputFormat: 'xlsx',
      operations: [
        {
          type: 'merge_sheets',
          sourceSheets: ['Q1', 'Q2'],
          outputSheetName: 'Combined',
          preserveSourceSheets: false,
        },
        {
          type: 'split_sheet',
          sourceSheetName: 'Combined',
          byColumn: 'Region',
          outputSheetPrefix: 'Region',
          preserveSourceSheet: true,
        },
      ],
    });

    const outputWorkbook = read(result.buffer, { type: 'buffer' });
    expect(outputWorkbook.SheetNames).toEqual(['Combined', 'Region East', 'Region West']);

    const combinedRows = utils.sheet_to_json(outputWorkbook.Sheets.Combined, {
      header: 1,
      raw: true,
      defval: '',
    });
    expect(combinedRows).toEqual([
      ['Source Sheet', 'Owner', 'Region', 'Amount'],
      ['Q1', 'Alice', 'East', 100],
      ['Q1', 'Bob', 'West', 90],
      ['Q2', 'Cara', 'East', 120],
      ['Q2', 'Dan', 'West', 80],
    ]);

    const eastRows = utils.sheet_to_json(outputWorkbook.Sheets['Region East'], {
      header: 1,
      raw: true,
      defval: '',
    });
    expect(eastRows).toEqual([
      ['Source Sheet', 'Owner', 'Region', 'Amount'],
      ['Q1', 'Alice', 'East', 100],
      ['Q2', 'Cara', 'East', 120],
    ]);
    expect(result.summary.operationsApplied.map((operation) => operation.type)).toEqual(
      expect.arrayContaining(['merge_sheets', 'split_sheet']),
    );
  });

  it('preserves existing cell formatting for transformed workbooks', async () => {
    const workbook = utils.book_new();
    const worksheet = utils.aoa_to_sheet([
      ['Employee', 'Revenue', 'Expense'],
      ['Alice', 2000, 500],
      ['Bob', 1500, 400],
    ]);
    worksheet.B1.s = { font: { bold: true, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '1F4E78' } } };
    worksheet.B2.s = { numFmt: '$#,##0.00', fill: { fgColor: { rgb: 'FFF2CC' } } };
    worksheet.B2.z = '$#,##0.00';
    worksheet.C2.s = { numFmt: '$#,##0.00', fill: { fgColor: { rgb: 'FCE4D6' } } };
    worksheet.C2.z = '$#,##0.00';
    utils.book_append_sheet(workbook, worksheet, 'Styled');

    const inputBuffer = Buffer.from(
      write(workbook, { type: 'buffer', bookType: 'xlsx', cellStyles: true }),
    );
    const result = await transformSpreadsheetBuffer({
      buffer: inputBuffer,
      sourceFilename: 'styled.xlsx',
      outputFormat: 'xlsx',
      operations: [
        {
          type: 'add_column',
          sheetName: 'Styled',
          columnName: 'Net',
          expression: '{{Revenue}} - {{Expense}}',
        },
      ],
    });

    const outputWorkbook = read(result.buffer, {
      type: 'buffer',
      cellStyles: true,
      cellNF: true,
    });
    expect(outputWorkbook.Sheets.Styled.B1.s).toBeTruthy();
    expect(outputWorkbook.Sheets.Styled.B2.s).toBeTruthy();
    expect(outputWorkbook.Sheets.Styled.B2.z).toBe('$#,##0.00');
    expect(outputWorkbook.Sheets.Styled.D2.s).toBeTruthy();
  });

  it('sanitizes Excel-like expressions and avoids failing when xlsx formulas can still be written', async () => {
    const workbook = utils.book_new();
    const worksheet = utils.aoa_to_sheet([
      ['Employee', 'Revenue', 'Expense', 'Status'],
      ['Alice', 2000, 500, ''],
      ['Bob', 900, 300, ''],
    ]);
    utils.book_append_sheet(workbook, worksheet, 'Sanitized');

    const inputBuffer = Buffer.from(write(workbook, { type: 'buffer', bookType: 'xlsx' }));
    const result = await transformSpreadsheetBuffer({
      buffer: inputBuffer,
      sourceFilename: 'sanitized.xlsx',
      outputFormat: 'xlsx',
      operations: [
        {
          type: 'update_cells',
          sheetName: 'Sanitized',
          columnName: 'Status',
          rowMatch: { Employee: 'Alice' },
          expression: '=IF({{Revenue}} >= 1000, "Healthy", "Watch")',
          formula: '=IF({{Revenue}} >= 1000, "Healthy", "Watch")',
        },
        {
          type: 'add_column',
          sheetName: 'Sanitized',
          columnName: 'MarginPct',
          expression: '=ROUND(({{Revenue}}-{{Expense}})/{{Revenue}}*100%, 2)',
          formula: '=ROUND(({{Revenue}}-{{Expense}})/{{Revenue}}, 2)',
        },
      ],
    });

    const outputWorkbook = read(result.buffer, { type: 'buffer', cellNF: true });
    const outputRows = utils.sheet_to_json(outputWorkbook.Sheets.Sanitized, {
      header: 1,
      raw: true,
      defval: '',
    });

    expect(outputRows).toEqual([
      ['Employee', 'Revenue', 'Expense', 'Status', 'MarginPct'],
      ['Alice', 2000, 500, 'Healthy', 0.75],
      ['Bob', 900, 300, '', 0.67],
    ]);
    expect(outputWorkbook.Sheets.Sanitized.E2.f).toBe('ROUND((B2-C2)/B2, 2)');
  });

  it('recognizes supported spreadsheet MIME types', () => {
    expect(isSpreadsheetTransformable('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')).toBe(true);
    expect(isSpreadsheetTransformable('text/csv')).toBe(true);
    expect(isSpreadsheetTransformable('application/pdf')).toBe(false);
  });
});
