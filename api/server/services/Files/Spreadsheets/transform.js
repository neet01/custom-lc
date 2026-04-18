const path = require('path');
const { excelMimeTypes } = require('librechat-data-provider');

const CSV_MIME_TYPES = new Set(['text/csv', 'application/csv']);
const SUPPORTED_SPREADSHEET_MIME_TYPES = new Set([
  ...CSV_MIME_TYPES,
  'application/vnd.oasis.opendocument.spreadsheet',
]);

const normalizeColumnName = (value) => String(value ?? '').trim().toLowerCase();

const unique = (values) => [...new Set(values)];

function isSpreadsheetTransformable(mimetype) {
  return Boolean(mimetype) && (excelMimeTypes.test(mimetype) || SUPPORTED_SPREADSHEET_MIME_TYPES.has(mimetype));
}

function normalizeColumnsInput(columns) {
  if (!Array.isArray(columns)) {
    return [];
  }

  return unique(
    columns
      .map((value) => String(value ?? '').trim())
      .filter(Boolean),
  );
}

function normalizeSheetNames(sheetNames) {
  if (!Array.isArray(sheetNames)) {
    return [];
  }

  return unique(
    sheetNames
      .map((value) => String(value ?? '').trim())
      .filter(Boolean),
  );
}

function getDefaultOutputFormat(sourceFilename) {
  return path.extname(sourceFilename).toLowerCase() === '.csv' ? 'csv' : 'xlsx';
}

function buildOutputFilename(sourceFilename, outputFormat) {
  const parsed = path.parse(sourceFilename || 'spreadsheet');
  const safeBase = parsed.name || 'spreadsheet';
  return `${safeBase}-transformed.${outputFormat}`;
}

async function inspectSpreadsheetBuffer({
  buffer,
  sourceFilename,
  maxPreviewRows = 5,
}) {
  if (!Buffer.isBuffer(buffer) || buffer.length === 0) {
    throw new Error('Spreadsheet buffer is required');
  }

  const previewRowCount = Math.min(Math.max(Number(maxPreviewRows) || 5, 1), 10);
  const { read, utils } = require('xlsx');
  const workbook = read(buffer, { type: 'buffer', cellDates: true });
  const workbookSheetNames = workbook.SheetNames ?? [];

  if (workbookSheetNames.length === 0) {
    throw new Error('Spreadsheet does not contain any sheets');
  }

  const sheets = workbookSheetNames.map((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    const rows = utils.sheet_to_json(worksheet, {
      header: 1,
      raw: false,
      defval: '',
    });

    const headerRow = Array.isArray(rows[0]) ? rows[0] : [];
    const columnNames = headerRow.map((value) => String(value ?? '').trim()).filter(Boolean);
    const previewRows = rows
      .slice(1, previewRowCount + 1)
      .map((row) =>
        columnNames.reduce((acc, columnName, columnIndex) => {
          acc[columnName] = Array.isArray(row) ? row[columnIndex] ?? '' : '';
          return acc;
        }, {}),
      );

    return {
      sheetName,
      rowCount: Math.max(rows.length - 1, 0),
      columnCount: columnNames.length,
      columns: columnNames,
      previewRows,
    };
  });

  return {
    filename: sourceFilename,
    sheetCount: sheets.length,
    sheets,
  };
}

async function transformSpreadsheetBuffer({
  buffer,
  sourceFilename,
  removeColumns = [],
  keepColumns = [],
  redactColumns = [],
  redactionText = '[REDACTED]',
  sheetNames = [],
  outputFormat,
}) {
  if (!Buffer.isBuffer(buffer) || buffer.length === 0) {
    throw new Error('Spreadsheet buffer is required');
  }

  const normalizedKeepColumns = normalizeColumnsInput(keepColumns);
  const normalizedRemoveColumns = normalizeColumnsInput(removeColumns);
  const normalizedRedactColumns = normalizeColumnsInput(redactColumns);
  const selectedSheetNames = normalizeSheetNames(sheetNames);

  if (
    normalizedKeepColumns.length === 0 &&
    normalizedRemoveColumns.length === 0 &&
    normalizedRedactColumns.length === 0
  ) {
    throw new Error('At least one spreadsheet transformation must be requested');
  }

  const resolvedOutputFormat = outputFormat ?? getDefaultOutputFormat(sourceFilename);
  if (resolvedOutputFormat !== 'xlsx' && resolvedOutputFormat !== 'csv') {
    throw new Error(`Unsupported output format: ${resolvedOutputFormat}`);
  }

  const { read, write, utils } = require('xlsx');
  const workbook = read(buffer, { type: 'buffer', cellDates: true });
  const workbookSheetNames = workbook.SheetNames ?? [];

  if (workbookSheetNames.length === 0) {
    throw new Error('Spreadsheet does not contain any sheets');
  }

  const targetSheetNames =
    selectedSheetNames.length > 0
      ? selectedSheetNames.filter((name) => workbookSheetNames.includes(name))
      : workbookSheetNames;

  if (selectedSheetNames.length > 0 && targetSheetNames.length === 0) {
    throw new Error('None of the requested sheet names were found in the spreadsheet');
  }

  if (resolvedOutputFormat === 'csv' && targetSheetNames.length !== 1) {
    throw new Error('CSV output requires exactly one sheet');
  }

  const matchedColumns = {
    keep: new Set(),
    remove: new Set(),
    redact: new Set(),
  };

  const transformedWorkbook = utils.book_new();
  const sheetSummaries = [];

  for (const sheetName of targetSheetNames) {
    const worksheet = workbook.Sheets[sheetName];
    const rows = utils.sheet_to_json(worksheet, {
      header: 1,
      raw: false,
      defval: '',
    });

    if (rows.length === 0) {
      const emptySheet = utils.aoa_to_sheet([]);
      utils.book_append_sheet(transformedWorkbook, emptySheet, sheetName);
      sheetSummaries.push({
        sheetName,
        rowCount: 0,
        originalColumnCount: 0,
        outputColumnCount: 0,
        removedColumns: [],
        keptColumns: [],
        redactedColumns: [],
      });
      continue;
    }

    const headerRow = Array.isArray(rows[0]) ? rows[0] : [];
    const normalizedHeaderMap = headerRow.map((header, index) => ({
      index,
      original: String(header ?? '').trim(),
      normalized: normalizeColumnName(header),
    }));

    let includedColumnIndexes = normalizedHeaderMap.map(({ index }) => index);
    if (normalizedKeepColumns.length > 0) {
      const keepSet = new Set(normalizedKeepColumns.map(normalizeColumnName));
      includedColumnIndexes = normalizedHeaderMap
        .filter(({ normalized }) => keepSet.has(normalized))
        .map(({ index }) => index);

      normalizedHeaderMap.forEach(({ original, normalized }) => {
        if (keepSet.has(normalized) && original) {
          matchedColumns.keep.add(original);
        }
      });
    }

    if (normalizedRemoveColumns.length > 0) {
      const removeSet = new Set(normalizedRemoveColumns.map(normalizeColumnName));
      includedColumnIndexes = includedColumnIndexes.filter((index) => {
        const header = normalizedHeaderMap[index];
        const shouldRemove = removeSet.has(header?.normalized);
        if (shouldRemove && header?.original) {
          matchedColumns.remove.add(header.original);
        }
        return !shouldRemove;
      });
    }

    const redactSet = new Set(normalizedRedactColumns.map(normalizeColumnName));
    const redactColumnIndexes = new Set(
      includedColumnIndexes.filter((index) => {
        const header = normalizedHeaderMap[index];
        const shouldRedact = redactSet.has(header?.normalized);
        if (shouldRedact && header?.original) {
          matchedColumns.redact.add(header.original);
        }
        return shouldRedact;
      }),
    );

    const transformedRows = rows.map((row, rowIndex) =>
      includedColumnIndexes.map((columnIndex) => {
        const value = Array.isArray(row) ? row[columnIndex] ?? '' : '';
        if (rowIndex > 0 && redactColumnIndexes.has(columnIndex) && value !== '') {
          return redactionText;
        }
        return value;
      }),
    );

    const transformedSheet = utils.aoa_to_sheet(transformedRows);
    utils.book_append_sheet(transformedWorkbook, transformedSheet, sheetName);

    const keptColumns = includedColumnIndexes
      .map((columnIndex) => normalizedHeaderMap[columnIndex]?.original)
      .filter(Boolean);

    const removedColumns = normalizedHeaderMap
      .filter(({ index }) => !includedColumnIndexes.includes(index))
      .map(({ original }) => original)
      .filter(Boolean);

    const redactedColumnsInSheet = includedColumnIndexes
      .filter((index) => redactColumnIndexes.has(index))
      .map((index) => normalizedHeaderMap[index]?.original)
      .filter(Boolean);

    sheetSummaries.push({
      sheetName,
      rowCount: Math.max(rows.length - 1, 0),
      originalColumnCount: normalizedHeaderMap.length,
      outputColumnCount: includedColumnIndexes.length,
      removedColumns,
      keptColumns,
      redactedColumns: redactedColumnsInSheet,
    });
  }

  const matchedColumnCount =
    matchedColumns.keep.size + matchedColumns.remove.size + matchedColumns.redact.size;

  if (matchedColumnCount === 0) {
    throw new Error('None of the requested columns were found in the spreadsheet');
  }

  let outputBuffer;
  let mimeType;
  if (resolvedOutputFormat === 'csv') {
    const firstSheet = transformedWorkbook.Sheets[targetSheetNames[0]];
    outputBuffer = Buffer.from(utils.sheet_to_csv(firstSheet), 'utf8');
    mimeType = 'text/csv';
  } else {
    outputBuffer = Buffer.from(write(transformedWorkbook, { type: 'buffer', bookType: 'xlsx' }));
    mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  }

  return {
    buffer: outputBuffer,
    bytes: outputBuffer.length,
    mimeType,
    filename: buildOutputFilename(sourceFilename, resolvedOutputFormat),
    summary: {
      outputFormat: resolvedOutputFormat,
      sheetCount: targetSheetNames.length,
      sheets: sheetSummaries,
      matchedColumns: {
        keep: [...matchedColumns.keep],
        remove: [...matchedColumns.remove],
        redact: [...matchedColumns.redact],
      },
    },
  };
}

module.exports = {
  inspectSpreadsheetBuffer,
  isSpreadsheetTransformable,
  transformSpreadsheetBuffer,
};
