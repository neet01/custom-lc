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

function cloneCell(cell) {
  return cell ? JSON.parse(JSON.stringify(cell)) : cell;
}

function isWorksheetMetaKey(key) {
  return key.startsWith('!');
}

function getCellDisplayValue(cell) {
  if (!cell) {
    return '';
  }

  if (cell.w != null) {
    return String(cell.w);
  }

  if (cell.v == null) {
    return '';
  }

  return String(cell.v);
}

function detectHeaderRowIndex({ worksheet, utils, range }) {
  const maxProbeRow = Math.min(range.s.r + 9, range.e.r);
  let bestRowIndex = range.s.r;
  let bestScore = -1;

  for (let rowIndex = range.s.r; rowIndex <= maxProbeRow; rowIndex += 1) {
    let nonEmptyCells = 0;
    for (let columnIndex = range.s.c; columnIndex <= range.e.c; columnIndex += 1) {
      const cell = worksheet[utils.encode_cell({ r: rowIndex, c: columnIndex })];
      if (getCellDisplayValue(cell).trim() !== '') {
        nonEmptyCells += 1;
      }
    }

    if (nonEmptyCells > bestScore) {
      bestScore = nonEmptyCells;
      bestRowIndex = rowIndex;
    }
  }

  return bestRowIndex;
}

function setRedactedCellValue(cell, redactionText) {
  const nextCell = cloneCell(cell) ?? {};
  nextCell.t = 's';
  nextCell.v = redactionText;
  nextCell.w = redactionText;
  delete nextCell.f;
  delete nextCell.r;
  delete nextCell.h;
  return nextCell;
}

function remapWorksheet({
  worksheet,
  utils,
  headerRowIndex,
  includedColumnIndexes,
  redactedColumnIndexes,
  redactionText,
}) {
  const range = utils.decode_range(worksheet['!ref']);
  const columnIndexMap = new Map(includedColumnIndexes.map((columnIndex, nextIndex) => [columnIndex, nextIndex]));
  const nextSheet = {};

  for (const [key, value] of Object.entries(worksheet)) {
    if (isWorksheetMetaKey(key)) {
      continue;
    }

    const address = utils.decode_cell(key);
    const nextColumnIndex = columnIndexMap.get(address.c);
    if (nextColumnIndex == null) {
      continue;
    }

    const nextAddress = utils.encode_cell({ r: address.r, c: nextColumnIndex });
    const nextCell =
      address.r > headerRowIndex &&
      redactedColumnIndexes.has(address.c) &&
      getCellDisplayValue(value) !== ''
        ? setRedactedCellValue(value, redactionText)
        : cloneCell(value);
    nextSheet[nextAddress] = nextCell;
  }

  nextSheet['!ref'] = utils.encode_range({
    s: { r: range.s.r, c: 0 },
    e: { r: range.e.r, c: Math.max(includedColumnIndexes.length - 1, 0) },
  });

  if (Array.isArray(worksheet['!merges'])) {
    const nextMerges = worksheet['!merges']
      .map((merge) => {
        const mappedColumns = [];
        for (let columnIndex = merge.s.c; columnIndex <= merge.e.c; columnIndex += 1) {
          const mapped = columnIndexMap.get(columnIndex);
          if (mapped != null) {
            mappedColumns.push(mapped);
          }
        }

        if (mappedColumns.length === 0) {
          return null;
        }

        return {
          s: { r: merge.s.r, c: Math.min(...mappedColumns) },
          e: { r: merge.e.r, c: Math.max(...mappedColumns) },
        };
      })
      .filter(Boolean);

    if (nextMerges.length > 0) {
      nextSheet['!merges'] = nextMerges;
    }
  }

  if (Array.isArray(worksheet['!cols'])) {
    nextSheet['!cols'] = includedColumnIndexes.map((columnIndex) =>
      worksheet['!cols'][columnIndex] ? { ...worksheet['!cols'][columnIndex] } : {},
    );
  }

  if (Array.isArray(worksheet['!rows'])) {
    nextSheet['!rows'] = worksheet['!rows'].map((row) => ({ ...row }));
  }

  if (worksheet['!autofilter']?.ref) {
    const autoFilterRange = utils.decode_range(worksheet['!autofilter'].ref);
    const mappedColumns = [];
    for (let columnIndex = autoFilterRange.s.c; columnIndex <= autoFilterRange.e.c; columnIndex += 1) {
      const mapped = columnIndexMap.get(columnIndex);
      if (mapped != null) {
        mappedColumns.push(mapped);
      }
    }

    if (mappedColumns.length > 0) {
      nextSheet['!autofilter'] = {
        ...worksheet['!autofilter'],
        ref: utils.encode_range({
          s: { r: autoFilterRange.s.r, c: Math.min(...mappedColumns) },
          e: { r: autoFilterRange.e.r, c: Math.max(...mappedColumns) },
        }),
      };
    }
  }

  return nextSheet;
}

function getHeaderMetadata({ worksheet, utils, range }) {
  const headers = [];
  for (let columnIndex = range.s.c; columnIndex <= range.e.c; columnIndex += 1) {
    const headerCell = worksheet[utils.encode_cell({ r: range.s.r, c: columnIndex })];
    headers.push({
      index: columnIndex,
      original: getCellDisplayValue(headerCell).trim(),
      normalized: normalizeColumnName(getCellDisplayValue(headerCell)),
    });
  }
  return headers;
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
  const workbook = read(buffer, { type: 'buffer', cellDates: true, cellStyles: true, cellNF: true });
  const workbookSheetNames = workbook.SheetNames ?? [];

  if (workbookSheetNames.length === 0) {
    throw new Error('Spreadsheet does not contain any sheets');
  }

  const sheets = workbookSheetNames.map((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    const range = worksheet?.['!ref'] ? utils.decode_range(worksheet['!ref']) : null;
    if (!range) {
      return {
        sheetName,
        rowCount: 0,
        columnCount: 0,
        columns: [],
        previewRows: [],
      };
    }

    const rows = utils.sheet_to_json(worksheet, {
      header: 1,
      raw: false,
      defval: '',
    });
    const headerRowIndex = detectHeaderRowIndex({ worksheet, utils, range });
    const relativeHeaderIndex = headerRowIndex - range.s.r;
    const headerRow = Array.isArray(rows[relativeHeaderIndex]) ? rows[relativeHeaderIndex] : [];
    const columnNames = headerRow.map((value) => String(value ?? '').trim()).filter(Boolean);
    const previewRows = rows
      .slice(relativeHeaderIndex + 1, relativeHeaderIndex + 1 + previewRowCount)
      .map((row) =>
        columnNames.reduce((acc, columnName, columnIndex) => {
          acc[columnName] = Array.isArray(row) ? row[columnIndex] ?? '' : '';
          return acc;
        }, {}),
      );

    return {
      sheetName,
      rowCount: Math.max(rows.length - (relativeHeaderIndex + 1), 0),
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
  const workbook = read(buffer, {
    type: 'buffer',
    cellDates: true,
    cellStyles: true,
    cellNF: true,
    bookVBA: true,
  });
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
  const sheetSummaries = [];

  for (const sheetName of targetSheetNames) {
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet?.['!ref']) {
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

    const range = utils.decode_range(worksheet['!ref']);
    const headerRowIndex = detectHeaderRowIndex({ worksheet, utils, range });
    const normalizedHeaderMap = getHeaderMetadata({
      worksheet,
      utils,
      range: { ...range, s: { ...range.s, r: headerRowIndex }, e: range.e },
    });
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
        const header = normalizedHeaderMap.find((item) => item.index === index);
        const shouldRemove = removeSet.has(header?.normalized);
        if (shouldRemove && header?.original) {
          matchedColumns.remove.add(header.original);
        }
        return !shouldRemove;
      });
    }

    const redactSet = new Set(normalizedRedactColumns.map(normalizeColumnName));
    const redactedColumnIndexes = new Set(
      includedColumnIndexes.filter((index) => {
        const header = normalizedHeaderMap.find((item) => item.index === index);
        const shouldRedact = redactSet.has(header?.normalized);
        if (shouldRedact && header?.original) {
          matchedColumns.redact.add(header.original);
        }
        return shouldRedact;
      }),
    );

    if (includedColumnIndexes.length === 0) {
      throw new Error(`All columns were removed from sheet "${sheetName}"`);
    }

    const transformedSheet =
      remapWorksheet({
        worksheet,
        utils,
        headerRowIndex,
        includedColumnIndexes,
        redactedColumnIndexes,
        redactionText,
      });

    workbook.Sheets[sheetName] = transformedSheet;

    const keptColumns = includedColumnIndexes
      .map((columnIndex) => normalizedHeaderMap.find((item) => item.index === columnIndex)?.original)
      .filter(Boolean);

    const removedColumns = normalizedHeaderMap
      .filter(({ index }) => !includedColumnIndexes.includes(index))
      .map(({ original }) => original)
      .filter(Boolean);

    const redactedColumnsInSheet = includedColumnIndexes
      .filter((index) => redactedColumnIndexes.has(index))
      .map((index) => normalizedHeaderMap.find((item) => item.index === index)?.original)
      .filter(Boolean);

    sheetSummaries.push({
      sheetName,
      rowCount: Math.max(range.e.r - headerRowIndex, 0),
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
    const firstSheet = workbook.Sheets[targetSheetNames[0]];
    outputBuffer = Buffer.from(utils.sheet_to_csv(firstSheet), 'utf8');
    mimeType = 'text/csv';
  } else {
    outputBuffer = Buffer.from(
      write(workbook, {
        type: 'buffer',
        bookType: 'xlsx',
        cellStyles: true,
        bookVBA: true,
      }),
    );
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
