const path = require('path');
const { create, all } = require('mathjs');
const { excelMimeTypes } = require('librechat-data-provider');

const math = create(all, {});

const CSV_MIME_TYPES = new Set(['text/csv', 'application/csv']);
const SUPPORTED_SPREADSHEET_MIME_TYPES = new Set([
  ...CSV_MIME_TYPES,
  'application/vnd.oasis.opendocument.spreadsheet',
]);

const FORMULA_FUNCTION_MAP = new Map([
  ['SUM', 'sum'],
  ['AVERAGE', 'mean'],
  ['MIN', 'min'],
  ['MAX', 'max'],
  ['ABS', 'abs'],
  ['ROUND', 'round'],
  ['CEILING', 'ceil'],
  ['FLOOR', 'floor'],
]);

const SHEET_OPERATION_TYPES = new Set([
  'add_column',
  'add_row',
  'update_cells',
  'sort_rows',
  'reorder_rows',
  'merge_sheets',
  'split_sheet',
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

function normalizeOperationsInput(operations) {
  if (!Array.isArray(operations)) {
    return [];
  }

  return operations
    .filter((operation) => operation && typeof operation === 'object')
    .map((operation) => ({
      ...operation,
      type: String(operation.type ?? '').trim(),
    }))
    .filter((operation) => SHEET_OPERATION_TYPES.has(operation.type));
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

function cloneArrayValue(value) {
  return Array.isArray(value) ? JSON.parse(JSON.stringify(value)) : [];
}

function cloneObjectValue(value) {
  return value && typeof value === 'object' ? JSON.parse(JSON.stringify(value)) : {};
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

function getHeaderColumnPlan({
  worksheet,
  utils,
  range,
  normalizedKeepColumns,
  normalizedRemoveColumns,
  normalizedRedactColumns,
}) {
  const headerRowIndex = detectHeaderRowIndex({ worksheet, utils, range });
  const normalizedHeaderMap = getHeaderMetadata({
    worksheet,
    utils,
    range: { ...range, s: { ...range.s, r: headerRowIndex }, e: range.e },
  });

  let includedColumnIndexes = normalizedHeaderMap.map(({ index }) => index);
  const matchedColumns = {
    keep: [],
    remove: [],
    redact: [],
  };

  if (normalizedKeepColumns.length > 0) {
    const keepSet = new Set(normalizedKeepColumns.map(normalizeColumnName));
    includedColumnIndexes = normalizedHeaderMap
      .filter(({ normalized }) => keepSet.has(normalized))
      .map(({ index }) => index);

    normalizedHeaderMap.forEach(({ original, normalized }) => {
      if (keepSet.has(normalized) && original) {
        matchedColumns.keep.push(original);
      }
    });
  }

  if (normalizedRemoveColumns.length > 0) {
    const removeSet = new Set(normalizedRemoveColumns.map(normalizeColumnName));
    includedColumnIndexes = includedColumnIndexes.filter((index) => {
      const header = normalizedHeaderMap.find((item) => item.index === index);
      const shouldRemove = removeSet.has(header?.normalized);
      if (shouldRemove && header?.original) {
        matchedColumns.remove.push(header.original);
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
        matchedColumns.redact.push(header.original);
      }
      return shouldRedact;
    }),
  );

  return {
    headerRowIndex,
    normalizedHeaderMap,
    includedColumnIndexes,
    redactedColumnIndexes,
    matchedColumns,
  };
}

function normalizeScalarValue(value) {
  if (value == null) {
    return '';
  }

  if (value instanceof Date) {
    return value.toISOString();
  }

  return String(value).trim();
}

function coerceMaybeNumber(value) {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : null;
  }

  if (typeof value !== 'string') {
    return null;
  }

  const trimmed = value.trim();
  if (trimmed === '') {
    return null;
  }

  const parsed = Number(trimmed.replace(/,/g, ''));
  return Number.isFinite(parsed) ? parsed : null;
}

function sanitizeFormulaFunctionNames(formula) {
  let nextFormula = formula;
  for (const [excelName, mathName] of FORMULA_FUNCTION_MAP.entries()) {
    nextFormula = nextFormula.replace(new RegExp(`\\b${excelName}\\s*\\(`, 'gi'), `${mathName}(`);
  }
  return nextFormula;
}

function buildExpressionScope(template, rowObject) {
  const scope = { row_number: Number(rowObject.__row_number) || 0 };
  let variableIndex = 0;
  const expression = String(template).replace(/{{\s*([^}]+?)\s*}}/g, (_match, rawColumnName) => {
    const columnName = String(rawColumnName).trim();
    const key = `col_${variableIndex}`;
    variableIndex += 1;

    const directValue = rowObject[columnName];
    const numericValue = coerceMaybeNumber(directValue);
    scope[key] = numericValue != null ? numericValue : directValue ?? '';
    return key;
  });

  return {
    expression,
    scope,
  };
}

function evaluateExpressionTemplate(template, rowObject) {
  const { expression, scope } = buildExpressionScope(template, rowObject);
  const normalizedExpression = sanitizeFormulaFunctionNames(String(expression).trim());
  return math.evaluate(normalizedExpression, scope);
}

function buildFormulaFromTemplate(template, headers, rowNumber, utils) {
  const formula = String(template ?? '').trim();
  if (!formula) {
    return null;
  }

  const withoutEquals = formula.startsWith('=') ? formula.slice(1) : formula;
  return withoutEquals.replace(/{{\s*([^}]+?)\s*}}/g, (_match, rawColumnName) => {
    const columnIndex = findColumnIndex(headers, rawColumnName);
    if (columnIndex === -1) {
      throw new Error(`Formula references unknown column "${rawColumnName}"`);
    }
    return utils.encode_cell({ r: rowNumber - 1, c: columnIndex }).replace(/\d+$/, `${rowNumber}`);
  });
}

function buildSheetSummary({
  sheetName,
  rowCount,
  originalColumnCount,
  outputColumnCount,
  removedColumns = [],
  keptColumns = [],
  redactedColumns = [],
}) {
  return {
    sheetName,
    rowCount,
    originalColumnCount,
    outputColumnCount,
    removedColumns,
    keptColumns,
    redactedColumns,
  };
}

function padRow(row, length) {
  const nextRow = Array.isArray(row) ? [...row] : [];
  while (nextRow.length < length) {
    nextRow.push('');
  }
  return nextRow;
}

function trimTrailingEmptyCells(row) {
  const nextRow = [...row];
  while (nextRow.length > 0 && normalizeScalarValue(nextRow[nextRow.length - 1]) === '') {
    nextRow.pop();
  }
  return nextRow;
}

function toWorksheetModel({ sheetName, worksheet, utils }) {
  if (!worksheet?.['!ref']) {
    return {
      sheetName,
      prefixRows: [],
      prefixRowMeta: [],
      headers: [],
      headerRowMeta: {},
      dataRows: [],
      dataRowMeta: [],
      cols: [],
      merges: [],
      autofilter: null,
      cellFormulas: new Map(),
    };
  }

  const range = utils.decode_range(worksheet['!ref']);
  const headerRowIndex = detectHeaderRowIndex({ worksheet, utils, range });
  const rows = utils.sheet_to_json(worksheet, {
    header: 1,
    raw: true,
    defval: '',
    blankrows: false,
  });

  const relativeHeaderIndex = headerRowIndex - range.s.r;
  const prefixRows = rows
    .slice(0, relativeHeaderIndex)
    .map((row) => trimTrailingEmptyCells(Array.isArray(row) ? row : []));
  const headers = trimTrailingEmptyCells(
    Array.isArray(rows[relativeHeaderIndex]) ? rows[relativeHeaderIndex].map((value) => String(value ?? '').trim()) : [],
  );
  const dataRows = rows
    .slice(relativeHeaderIndex + 1)
    .map((row) => padRow(Array.isArray(row) ? row : [], headers.length));

  const rowMeta = cloneArrayValue(worksheet['!rows']);
  const headerMetaIndex = prefixRows.length;

  return {
    sheetName,
    prefixRows: prefixRows.map((row) => padRow(row, headers.length)),
    prefixRowMeta: rowMeta.slice(0, prefixRows.length).map(cloneObjectValue),
    headers,
    headerRowMeta: cloneObjectValue(rowMeta[headerMetaIndex]),
    dataRows,
    dataRowMeta: rowMeta
      .slice(headerMetaIndex + 1, headerMetaIndex + 1 + dataRows.length)
      .map(cloneObjectValue),
    cols: cloneArrayValue(worksheet['!cols']).map(cloneObjectValue),
    merges: cloneArrayValue(worksheet['!merges']),
    autofilter: worksheet['!autofilter'] ? cloneObjectValue(worksheet['!autofilter']) : null,
    cellFormulas: new Map(),
  };
}

function cloneWorksheetModel(model, sheetNameOverride) {
  const nextModel = {
    sheetName: sheetNameOverride ?? model.sheetName,
    prefixRows: model.prefixRows.map((row) => [...row]),
    prefixRowMeta: model.prefixRowMeta.map(cloneObjectValue),
    headers: [...model.headers],
    headerRowMeta: cloneObjectValue(model.headerRowMeta),
    dataRows: model.dataRows.map((row) => [...row]),
    dataRowMeta: model.dataRowMeta.map(cloneObjectValue),
    cols: model.cols.map(cloneObjectValue),
    merges: cloneArrayValue(model.merges),
    autofilter: model.autofilter ? cloneObjectValue(model.autofilter) : null,
    cellFormulas: new Map(model.cellFormulas),
  };

  return nextModel;
}

function getFormulaKey(dataRowIndex, columnIndex) {
  return `${dataRowIndex}:${columnIndex}`;
}

function setFormulaCell(model, dataRowIndex, columnIndex, formula) {
  model.cellFormulas.set(getFormulaKey(dataRowIndex, columnIndex), formula);
}

function clearFormulaCell(model, dataRowIndex, columnIndex) {
  model.cellFormulas.delete(getFormulaKey(dataRowIndex, columnIndex));
}

function getFormulaEntriesForRow(model, dataRowIndex) {
  const entries = [];
  for (const [key, formula] of model.cellFormulas.entries()) {
    const [rowIndexRaw, columnIndexRaw] = key.split(':');
    const rowIndex = Number(rowIndexRaw);
    if (rowIndex !== dataRowIndex) {
      continue;
    }
    entries.push({
      columnIndex: Number(columnIndexRaw),
      formula,
    });
  }
  return entries;
}

function rebaseFormulaRowNumbers(formula, fromRowNumber, toRowNumber) {
  return String(formula).replace(/(\$?[A-Z]{1,3}\$?)(\d+)/g, (match, columnPart, rowPart) => {
    const numericRow = Number(rowPart);
    if (numericRow !== fromRowNumber) {
      return match;
    }
    return `${columnPart}${toRowNumber}`;
  });
}

function shiftFormulaColumns(model, insertIndex) {
  const nextEntries = [];
  for (const [key, formula] of model.cellFormulas.entries()) {
    const [rowIndexRaw, columnIndexRaw] = key.split(':');
    const rowIndex = Number(rowIndexRaw);
    const columnIndex = Number(columnIndexRaw);
    nextEntries.push([
      getFormulaKey(rowIndex, columnIndex >= insertIndex ? columnIndex + 1 : columnIndex),
      formula,
    ]);
  }
  model.cellFormulas = new Map(nextEntries);
}

function getExcelDataRowNumber(model, dataRowIndex) {
  return model.prefixRows.length + 2 + dataRowIndex;
}

function buildRowObject(model, row, dataRowIndex) {
  const rowObject = { __row_number: dataRowIndex + 1 };
  model.headers.forEach((header, columnIndex) => {
    if (header) {
      rowObject[header] = row[columnIndex] ?? '';
    }
  });
  return rowObject;
}

function findColumnIndex(headers, columnName) {
  const normalizedTarget = normalizeColumnName(columnName);
  return headers.findIndex((header) => normalizeColumnName(header) === normalizedTarget);
}

function resolveColumnIndex(headers, columnName, contextMessage = 'column') {
  const columnIndex = findColumnIndex(headers, columnName);
  if (columnIndex === -1) {
    throw new Error(`Could not find ${contextMessage} "${columnName}"`);
  }
  return columnIndex;
}

function getRowTargetIndexes(model, operation) {
  if (Number.isInteger(operation.rowNumber)) {
    if (operation.rowNumber < 1 || operation.rowNumber > model.dataRows.length) {
      throw new Error(
        `Row number ${operation.rowNumber} is out of range for sheet "${model.sheetName}"`,
      );
    }
    return [operation.rowNumber - 1];
  }

  if (operation.rowMatch && typeof operation.rowMatch === 'object') {
    const entries = Object.entries(operation.rowMatch);
    return model.dataRows
      .map((row, dataRowIndex) => ({ row, dataRowIndex }))
      .filter(({ row, dataRowIndex }) => {
        const rowObject = buildRowObject(model, row, dataRowIndex);
        return entries.every(([columnName, expectedValue]) => {
          const columnIndex = findColumnIndex(model.headers, columnName);
          if (columnIndex === -1) {
            return false;
          }

          return normalizeScalarValue(rowObject[model.headers[columnIndex]]) === normalizeScalarValue(expectedValue);
        });
      })
      .map(({ dataRowIndex }) => dataRowIndex);
  }

  return model.dataRows.map((_row, index) => index);
}

function resolveInsertIndex({ length, operation, beforeColumnIndex, afterColumnIndex }) {
  if (beforeColumnIndex != null) {
    return beforeColumnIndex;
  }

  if (afterColumnIndex != null) {
    return afterColumnIndex + 1;
  }

  if (Number.isInteger(operation.index)) {
    return Math.max(0, Math.min(length, operation.index));
  }

  if (operation.position === 'start') {
    return 0;
  }

  return length;
}

function shiftMergesForColumnInsert(merges, insertIndex) {
  return merges.map((merge) => {
    const nextMerge = cloneObjectValue(merge);
    if (nextMerge.s.c >= insertIndex) {
      nextMerge.s.c += 1;
    }
    if (nextMerge.e.c >= insertIndex) {
      nextMerge.e.c += 1;
    }
    return nextMerge;
  });
}

function removeDataRegionMerges(model) {
  const headerRowIndex = model.prefixRows.length;
  model.merges = model.merges.filter((merge) => merge.e.r <= headerRowIndex);
}

function buildCalculatedCellValue({
  operation,
  rowObject,
  headers,
  rowNumber,
  outputFormat,
  utils,
}) {
  const hasDirectValue = Object.prototype.hasOwnProperty.call(operation, 'value');
  const hasDefaultValue = Object.prototype.hasOwnProperty.call(operation, 'defaultValue');
  const hasExpression = typeof operation.expression === 'string' && operation.expression.trim() !== '';
  const hasFormula = typeof operation.formula === 'string' && operation.formula.trim() !== '';

  let value;
  if (hasDirectValue) {
    value = operation.value;
  } else if (hasDefaultValue) {
    value = operation.defaultValue;
  }

  if (hasExpression) {
    try {
      value = evaluateExpressionTemplate(operation.expression, rowObject);
    } catch (error) {
      throw new Error(`Failed to evaluate expression "${operation.expression}": ${error.message}`);
    }
  }

  let formula = null;
  if (hasFormula) {
    formula = buildFormulaFromTemplate(operation.formula, headers, rowNumber, utils);

    if (!hasExpression) {
      try {
        value = evaluateExpressionTemplate(String(operation.formula).replace(/^=/, ''), rowObject);
      } catch (_error) {
        if (outputFormat === 'csv') {
          throw new Error(
            'Formula-only spreadsheet operations require xlsx output or an accompanying expression for CSV exports',
          );
        }

        if (value == null) {
          value = '';
        }
      }
    }
  }

  return { value, formula };
}

function applyAddColumnOperation({ model, operation, outputFormat, utils }) {
  const columnName = String(operation.columnName ?? '').trim();
  if (!columnName) {
    throw new Error('add_column operations require columnName');
  }

  if (findColumnIndex(model.headers, columnName) !== -1) {
    throw new Error(`Sheet "${model.sheetName}" already contains column "${columnName}"`);
  }

  const beforeColumnIndex =
    operation.beforeColumn != null
      ? resolveColumnIndex(model.headers, operation.beforeColumn, 'beforeColumn')
      : null;
  const afterColumnIndex =
    operation.afterColumn != null
      ? resolveColumnIndex(model.headers, operation.afterColumn, 'afterColumn')
      : null;
  const insertIndex = resolveInsertIndex({
    length: model.headers.length,
    operation,
    beforeColumnIndex,
    afterColumnIndex,
  });

  shiftFormulaColumns(model, insertIndex);
  model.headers.splice(insertIndex, 0, columnName);
  model.cols.splice(insertIndex, 0, {});
  model.prefixRows = model.prefixRows.map((row) => {
    const nextRow = padRow(row, model.headers.length - 1);
    nextRow.splice(insertIndex, 0, '');
    return nextRow;
  });
  model.merges = shiftMergesForColumnInsert(model.merges, insertIndex);

  model.dataRows = model.dataRows.map((row, dataRowIndex) => {
    const nextRow = padRow(row, model.headers.length - 1);
    const rowObject = buildRowObject(
      {
        ...model,
        headers: model.headers.filter((_header, headerIndex) => headerIndex !== insertIndex),
      },
      row,
      dataRowIndex,
    );
    const rowNumber = getExcelDataRowNumber(model, dataRowIndex);
    const { value, formula } = buildCalculatedCellValue({
      operation,
      rowObject,
      headers: model.headers,
      rowNumber,
      outputFormat,
      utils,
    });

    nextRow.splice(insertIndex, 0, value ?? '');
    if (formula) {
      setFormulaCell(model, dataRowIndex, insertIndex, formula);
    }
    return nextRow;
  });
}

function applyUpdateCellsOperation({ model, operation, outputFormat, utils }) {
  const columnName = String(operation.columnName ?? '').trim();
  if (!columnName) {
    throw new Error('update_cells operations require columnName');
  }

  const columnIndex = resolveColumnIndex(model.headers, columnName, 'column');
  const rowIndexes = getRowTargetIndexes(model, operation);
  if (rowIndexes.length === 0) {
    throw new Error(`No rows matched update_cells operation on sheet "${model.sheetName}"`);
  }

  rowIndexes.forEach((dataRowIndex) => {
    const row = padRow(model.dataRows[dataRowIndex], model.headers.length);
    const rowObject = buildRowObject(model, row, dataRowIndex);
    const rowNumber = getExcelDataRowNumber(model, dataRowIndex);
    const { value, formula } = buildCalculatedCellValue({
      operation,
      rowObject,
      headers: model.headers,
      rowNumber,
      outputFormat,
      utils,
    });

    row[columnIndex] = value ?? '';
    model.dataRows[dataRowIndex] = row;

    if (formula) {
      setFormulaCell(model, dataRowIndex, columnIndex, formula);
    } else {
      clearFormulaCell(model, dataRowIndex, columnIndex);
    }
  });
}

function applyAddRowOperation({ model, operation }) {
  if (!operation.values || typeof operation.values !== 'object') {
    throw new Error('add_row operations require a values object');
  }

  const nextRow = new Array(model.headers.length).fill('');
  for (const [columnName, value] of Object.entries(operation.values)) {
    const columnIndex = resolveColumnIndex(model.headers, columnName, 'row value column');
    nextRow[columnIndex] = value ?? '';
  }

  let insertIndex = model.dataRows.length;
  if (Number.isInteger(operation.index)) {
    insertIndex = Math.max(0, Math.min(model.dataRows.length, operation.index - 1));
  } else if (operation.position === 'start') {
    insertIndex = 0;
  }

  model.dataRows.splice(insertIndex, 0, nextRow);
  model.dataRowMeta.splice(insertIndex, 0, {});

  if (model.cellFormulas.size > 0) {
    const nextEntries = [];
    for (const [key, formula] of model.cellFormulas.entries()) {
      const [rowIndexRaw, columnIndexRaw] = key.split(':');
      const rowIndex = Number(rowIndexRaw);
      const columnIndex = Number(columnIndexRaw);
      const nextRowIndex = rowIndex >= insertIndex ? rowIndex + 1 : rowIndex;
      const oldRowNumber = getExcelDataRowNumber(model, rowIndex);
      const nextRowNumber = getExcelDataRowNumber(model, nextRowIndex);
      nextEntries.push([
        getFormulaKey(nextRowIndex, columnIndex),
        rowIndex >= insertIndex
          ? rebaseFormulaRowNumbers(formula, oldRowNumber, nextRowNumber)
          : formula,
      ]);
    }
    model.cellFormulas = new Map(nextEntries);
  }
}

function normalizeSortColumns(operation) {
  if (Array.isArray(operation.columns) && operation.columns.length > 0) {
    return operation.columns;
  }

  if (operation.columnName) {
    return [
      {
        columnName: operation.columnName,
        direction: operation.direction,
        numeric: operation.numeric,
      },
    ];
  }

  throw new Error('sort_rows operations require columnName or columns');
}

function compareSortValues(leftValue, rightValue, sortSpec) {
  const leftBlank = normalizeScalarValue(leftValue) === '';
  const rightBlank = normalizeScalarValue(rightValue) === '';
  if (leftBlank && rightBlank) {
    return 0;
  }
  if (leftBlank) {
    return 1;
  }
  if (rightBlank) {
    return -1;
  }

  const leftNumeric = coerceMaybeNumber(leftValue);
  const rightNumeric = coerceMaybeNumber(rightValue);
  if (sortSpec.numeric || (leftNumeric != null && rightNumeric != null)) {
    if (leftNumeric === rightNumeric) {
      return 0;
    }
    return leftNumeric < rightNumeric ? -1 : 1;
  }

  return normalizeScalarValue(leftValue).localeCompare(normalizeScalarValue(rightValue), undefined, {
    numeric: true,
    sensitivity: 'base',
  });
}

function applySortRowsOperation({ model, operation }) {
  const sortSpecs = normalizeSortColumns(operation).map((sortSpec) => ({
    ...sortSpec,
    direction: sortSpec.direction === 'desc' ? 'desc' : 'asc',
    columnIndex: resolveColumnIndex(model.headers, sortSpec.columnName, 'sort column'),
  }));

  const decoratedRows = model.dataRows.map((row, dataRowIndex) => ({
    row: [...row],
    rowMeta: cloneObjectValue(model.dataRowMeta[dataRowIndex]),
    formulas: getFormulaEntriesForRow(model, dataRowIndex),
    originalIndex: dataRowIndex,
  }));

  decoratedRows.sort((left, right) => {
    for (const sortSpec of sortSpecs) {
      const comparison = compareSortValues(
        left.row[sortSpec.columnIndex],
        right.row[sortSpec.columnIndex],
        sortSpec,
      );

      if (comparison !== 0) {
        return sortSpec.direction === 'desc' ? comparison * -1 : comparison;
      }
    }

    return left.originalIndex - right.originalIndex;
  });

  model.dataRows = decoratedRows.map((entry) => entry.row);
  model.dataRowMeta = decoratedRows.map((entry) => entry.rowMeta);
  model.cellFormulas = new Map();
  decoratedRows.forEach((entry, newRowIndex) => {
    const oldRowNumber = getExcelDataRowNumber(model, entry.originalIndex);
    const newRowNumber = getExcelDataRowNumber(model, newRowIndex);
    entry.formulas.forEach(({ columnIndex, formula }) => {
      setFormulaCell(
        model,
        newRowIndex,
        columnIndex,
        rebaseFormulaRowNumbers(formula, oldRowNumber, newRowNumber),
      );
    });
  });
  removeDataRegionMerges(model);
}

function applyReorderRowsOperation({ model, operation }) {
  if (!Array.isArray(operation.orderedRowNumbers) || operation.orderedRowNumbers.length === 0) {
    throw new Error('reorder_rows operations require orderedRowNumbers');
  }

  const requestedIndexes = unique(
    operation.orderedRowNumbers
      .map((value) => Number(value))
      .filter((value) => Number.isInteger(value) && value >= 1),
  );

  if (requestedIndexes.length === 0) {
    throw new Error('reorder_rows orderedRowNumbers must contain positive integers');
  }

  for (const rowNumber of requestedIndexes) {
    if (rowNumber > model.dataRows.length) {
      throw new Error(`Row number ${rowNumber} is out of range for sheet "${model.sheetName}"`);
    }
  }

  const requestedSet = new Set(requestedIndexes.map((value) => value - 1));
  const orderedEntries = requestedIndexes.map((value) => ({
    row: [...model.dataRows[value - 1]],
    rowMeta: cloneObjectValue(model.dataRowMeta[value - 1]),
    formulas: getFormulaEntriesForRow(model, value - 1),
    originalIndex: value - 1,
  }));

  const appendRemaining = operation.appendRemaining !== false;
  const remainingEntries = appendRemaining
    ? model.dataRows
        .map((row, dataRowIndex) => ({
          row: [...row],
          rowMeta: cloneObjectValue(model.dataRowMeta[dataRowIndex]),
          formulas: getFormulaEntriesForRow(model, dataRowIndex),
          dataRowIndex,
          originalIndex: dataRowIndex,
        }))
        .filter((entry) => !requestedSet.has(entry.dataRowIndex))
    : [];

  model.dataRows = orderedEntries.concat(remainingEntries).map((entry) => entry.row);
  model.dataRowMeta = orderedEntries.concat(remainingEntries).map((entry) => entry.rowMeta);
  const reorderedEntries = orderedEntries.concat(remainingEntries);
  model.cellFormulas = new Map();
  reorderedEntries.forEach((entry, newRowIndex) => {
    const oldRowNumber = getExcelDataRowNumber(model, entry.originalIndex);
    const newRowNumber = getExcelDataRowNumber(model, newRowIndex);
    entry.formulas.forEach(({ columnIndex, formula }) => {
      setFormulaCell(
        model,
        newRowIndex,
        columnIndex,
        rebaseFormulaRowNumbers(formula, oldRowNumber, newRowNumber),
      );
    });
  });
  removeDataRegionMerges(model);
}

function sanitizeSheetName(value, fallback = 'Sheet') {
  const normalized = String(value ?? '').replace(/[\\/?*\[\]:]/g, ' ').trim() || fallback;
  return normalized.slice(0, 31);
}

function ensureUniqueSheetName(sheetOrder, desiredName) {
  const usedNames = new Set(sheetOrder);
  if (!usedNames.has(desiredName)) {
    return desiredName;
  }

  let suffix = 2;
  while (suffix < 1000) {
    const candidate = sanitizeSheetName(`${desiredName.slice(0, 27)} ${suffix}`, desiredName);
    if (!usedNames.has(candidate)) {
      return candidate;
    }
    suffix += 1;
  }

  throw new Error(`Could not generate a unique sheet name from "${desiredName}"`);
}

function buildMergedSheetModel({
  sourceModels,
  outputSheetName,
  includeSourceSheetColumn,
}) {
  const mergedHeaders = [];
  const seenHeaders = new Set();

  if (includeSourceSheetColumn) {
    mergedHeaders.push('Source Sheet');
    seenHeaders.add(normalizeColumnName('Source Sheet'));
  }

  for (const model of sourceModels) {
    for (const header of model.headers) {
      const normalized = normalizeColumnName(header);
      if (!header || seenHeaders.has(normalized)) {
        continue;
      }
      seenHeaders.add(normalized);
      mergedHeaders.push(header);
    }
  }

  const dataRows = [];
  for (const model of sourceModels) {
    for (const row of model.dataRows) {
      const rowObject = buildRowObject(model, row, 0);
      const nextRow = mergedHeaders.map((header) => {
        if (header === 'Source Sheet') {
          return model.sheetName;
        }
        return rowObject[header] ?? '';
      });
      dataRows.push(nextRow);
    }
  }

  return {
    sheetName: outputSheetName,
    prefixRows: [],
    prefixRowMeta: [],
    headers: mergedHeaders,
    headerRowMeta: {},
    dataRows,
    dataRowMeta: dataRows.map(() => ({})),
    cols: mergedHeaders.map(() => ({})),
    merges: [],
    autofilter: null,
    cellFormulas: new Map(),
  };
}

function applyMergeSheetsOperation({ sheetModels, sheetOrder, operation, operationSummaries }) {
  const sourceSheets = normalizeSheetNames(operation.sourceSheets);
  if (sourceSheets.length < 2) {
    throw new Error('merge_sheets operations require at least two sourceSheets');
  }

  const sourceModels = sourceSheets.map((sheetName) => {
    const model = sheetModels.get(sheetName);
    if (!model) {
      throw new Error(`merge_sheets could not find source sheet "${sheetName}"`);
    }
    return model;
  });

  const desiredName = sanitizeSheetName(
    operation.outputSheetName || `${sourceSheets[0]} Merged`,
    'Merged Sheet',
  );
  const collidesWithOtherSheet =
    sheetModels.has(desiredName) && !sourceSheets.includes(desiredName);
  if (collidesWithOtherSheet) {
    throw new Error(`merge_sheets output sheet "${desiredName}" already exists`);
  }

  const outputSheetName = sheetModels.has(desiredName)
    ? desiredName
    : ensureUniqueSheetName(sheetOrder, desiredName);

  const mergedModel = buildMergedSheetModel({
    sourceModels,
    outputSheetName,
    includeSourceSheetColumn: operation.includeSourceSheetColumn !== false,
  });

  sheetModels.set(outputSheetName, mergedModel);
  if (!sheetOrder.includes(outputSheetName)) {
    sheetOrder.push(outputSheetName);
  }

  if (operation.preserveSourceSheets === false) {
    for (const sourceSheet of sourceSheets) {
      if (sourceSheet === outputSheetName) {
        continue;
      }
      sheetModels.delete(sourceSheet);
      const existingIndex = sheetOrder.indexOf(sourceSheet);
      if (existingIndex !== -1) {
        sheetOrder.splice(existingIndex, 1);
      }
    }
  }

  operationSummaries.push({
    type: 'merge_sheets',
    sourceSheets,
    outputSheetName,
  });
}

function buildSplitSheetName({ sourceSheet, outputPrefix, groupValue }) {
  const safeGroupName = sanitizeSheetName(groupValue || 'Blank', 'Blank');
  const prefix = String(outputPrefix ?? `${sourceSheet} -`).trim();
  return sanitizeSheetName(`${prefix} ${safeGroupName}`, `${sourceSheet} Split`);
}

function applySplitSheetOperation({ sheetModels, sheetOrder, operation, operationSummaries }) {
  const sourceSheetName = String(operation.sourceSheetName ?? operation.sheetName ?? '').trim();
  if (!sourceSheetName) {
    throw new Error('split_sheet operations require sourceSheetName');
  }

  const sourceModel = sheetModels.get(sourceSheetName);
  if (!sourceModel) {
    throw new Error(`split_sheet could not find source sheet "${sourceSheetName}"`);
  }

  const byColumn = String(operation.byColumn ?? '').trim();
  if (!byColumn) {
    throw new Error('split_sheet operations require byColumn');
  }

  const splitColumnIndex = resolveColumnIndex(sourceModel.headers, byColumn, 'split column');
  const groupedRows = new Map();

  sourceModel.dataRows.forEach((row, dataRowIndex) => {
    const groupValue = normalizeScalarValue(row[splitColumnIndex]) || 'Blank';
    if (!groupedRows.has(groupValue)) {
      groupedRows.set(groupValue, []);
    }
    groupedRows.get(groupValue).push({
      row: [...row],
      rowMeta: cloneObjectValue(sourceModel.dataRowMeta[dataRowIndex]),
    });
  });

  const createdSheets = [];
  for (const [groupValue, rows] of groupedRows.entries()) {
    const desiredName = buildSplitSheetName({
      sourceSheet: sourceSheetName,
      outputPrefix: operation.outputSheetPrefix,
      groupValue,
    });
    const outputSheetName = ensureUniqueSheetName(sheetOrder, desiredName);
    const splitModel = cloneWorksheetModel(sourceModel, outputSheetName);
    splitModel.dataRows = rows.map((entry) => entry.row);
    splitModel.dataRowMeta = rows.map((entry) => entry.rowMeta);
    splitModel.cellFormulas = new Map();
    sheetModels.set(outputSheetName, splitModel);
    sheetOrder.push(outputSheetName);
    createdSheets.push(outputSheetName);
  }

  if (operation.preserveSourceSheet === false) {
    sheetModels.delete(sourceSheetName);
    const existingIndex = sheetOrder.indexOf(sourceSheetName);
    if (existingIndex !== -1) {
      sheetOrder.splice(existingIndex, 1);
    }
  }

  operationSummaries.push({
    type: 'split_sheet',
    sourceSheetName,
    byColumn,
    createdSheets,
  });
}

function buildWorksheetFromModel({ model, utils }) {
  const width = model.headers.length;
  const aoa = [
    ...model.prefixRows.map((row) => padRow(row, width)),
    padRow(model.headers, width),
    ...model.dataRows.map((row) => padRow(row, width)),
  ];

  const worksheet = utils.aoa_to_sheet(aoa.length > 0 ? aoa : [[]], {
    cellDates: true,
  });

  if (model.merges.length > 0) {
    worksheet['!merges'] = cloneArrayValue(model.merges);
  }

  if (model.cols.length > 0) {
    worksheet['!cols'] = model.cols.map(cloneObjectValue);
  }

  const rowMeta = [
    ...model.prefixRowMeta.map(cloneObjectValue),
    cloneObjectValue(model.headerRowMeta),
    ...model.dataRowMeta.map(cloneObjectValue),
  ].filter((row, index) => row && Object.keys(row).length > 0 && index < aoa.length);

  if (rowMeta.length > 0) {
    worksheet['!rows'] = rowMeta;
  }

  if (width > 0 && aoa.length > model.prefixRows.length + 1) {
    const headerRowNumber = model.prefixRows.length + 1;
    worksheet['!autofilter'] = {
      ref: utils.encode_range({
        s: { r: headerRowNumber - 1, c: 0 },
        e: { r: aoa.length - 1, c: width - 1 },
      }),
    };
  } else if (model.autofilter?.ref) {
    worksheet['!autofilter'] = cloneObjectValue(model.autofilter);
  }

  for (const [key, formula] of model.cellFormulas.entries()) {
    const [dataRowIndexRaw, columnIndexRaw] = key.split(':');
    const dataRowIndex = Number(dataRowIndexRaw);
    const columnIndex = Number(columnIndexRaw);
    const rowNumber = getExcelDataRowNumber(model, dataRowIndex);
    const address = utils.encode_cell({ r: rowNumber - 1, c: columnIndex });
    const cell = worksheet[address] ?? { t: 'n', v: '' };
    const cachedValue = model.dataRows[dataRowIndex]?.[columnIndex];
    if (cachedValue != null && cachedValue !== '') {
      cell.v = cachedValue;
      cell.t = typeof cachedValue === 'number' ? 'n' : 's';
      cell.w = String(cachedValue);
    }
    cell.f = formula;
    worksheet[address] = cell;
  }

  return worksheet;
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
  operations = [],
}) {
  if (!Buffer.isBuffer(buffer) || buffer.length === 0) {
    throw new Error('Spreadsheet buffer is required');
  }

  const normalizedKeepColumns = normalizeColumnsInput(keepColumns);
  const normalizedRemoveColumns = normalizeColumnsInput(removeColumns);
  const normalizedRedactColumns = normalizeColumnsInput(redactColumns);
  const selectedSheetNames = normalizeSheetNames(sheetNames);
  const normalizedOperations = normalizeOperationsInput(operations);

  const hasLegacyTransforms =
    normalizedKeepColumns.length > 0 ||
    normalizedRemoveColumns.length > 0 ||
    normalizedRedactColumns.length > 0;

  if (!hasLegacyTransforms && normalizedOperations.length === 0) {
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

  const matchedColumns = {
    keep: new Set(),
    remove: new Set(),
    redact: new Set(),
  };
  const sheetSummaries = [];
  const sheetModels = new Map();
  const sheetOrder = [...workbookSheetNames];

  for (const sheetName of workbookSheetNames) {
    let worksheet = workbook.Sheets[sheetName];

    if (targetSheetNames.includes(sheetName) && worksheet?.['!ref']) {
      const range = utils.decode_range(worksheet['!ref']);
      const {
        headerRowIndex,
        normalizedHeaderMap,
        includedColumnIndexes,
        redactedColumnIndexes,
        matchedColumns: sheetMatchedColumns,
      } = getHeaderColumnPlan({
        worksheet,
        utils,
        range,
        normalizedKeepColumns,
        normalizedRemoveColumns,
        normalizedRedactColumns,
      });

      if (hasLegacyTransforms) {
        if (includedColumnIndexes.length === 0) {
          throw new Error(`All columns were removed from sheet "${sheetName}"`);
        }

        worksheet = remapWorksheet({
          worksheet,
          utils,
          headerRowIndex,
          includedColumnIndexes,
          redactedColumnIndexes,
          redactionText,
        });

        sheetMatchedColumns.keep.forEach((column) => matchedColumns.keep.add(column));
        sheetMatchedColumns.remove.forEach((column) => matchedColumns.remove.add(column));
        sheetMatchedColumns.redact.forEach((column) => matchedColumns.redact.add(column));

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

        sheetSummaries.push(
          buildSheetSummary({
            sheetName,
            rowCount: Math.max(range.e.r - headerRowIndex, 0),
            originalColumnCount: normalizedHeaderMap.length,
            outputColumnCount: includedColumnIndexes.length,
            removedColumns,
            keptColumns,
            redactedColumns: redactedColumnsInSheet,
          }),
        );
      } else {
        sheetSummaries.push(
          buildSheetSummary({
            sheetName,
            rowCount: Math.max(range.e.r - headerRowIndex, 0),
            originalColumnCount: normalizedHeaderMap.length,
            outputColumnCount: normalizedHeaderMap.length,
          }),
        );
      }
    }

    sheetModels.set(sheetName, toWorksheetModel({ sheetName, worksheet, utils }));
  }

  if (hasLegacyTransforms) {
    const matchedColumnCount =
      matchedColumns.keep.size + matchedColumns.remove.size + matchedColumns.redact.size;
    if (matchedColumnCount === 0) {
      throw new Error('None of the requested columns were found in the spreadsheet');
    }
  }

  const operationSummaries = [];
  for (const operation of normalizedOperations) {
    if (operation.type === 'merge_sheets') {
      applyMergeSheetsOperation({ sheetModels, sheetOrder, operation, operationSummaries });
      continue;
    }

    if (operation.type === 'split_sheet') {
      applySplitSheetOperation({ sheetModels, sheetOrder, operation, operationSummaries });
      continue;
    }

    const explicitSheetName = String(operation.sheetName ?? '').trim();
    const operationSheetNames = explicitSheetName
      ? [explicitSheetName]
      : targetSheetNames;

    if (operationSheetNames.length === 0) {
      throw new Error(`Operation "${operation.type}" did not resolve to any sheet`);
    }

    for (const sheetName of operationSheetNames) {
      const model = sheetModels.get(sheetName);
      if (!model) {
        throw new Error(`Operation "${operation.type}" could not find sheet "${sheetName}"`);
      }

      if (operation.type === 'add_column') {
        applyAddColumnOperation({ model, operation, outputFormat: resolvedOutputFormat, utils });
      } else if (operation.type === 'update_cells') {
        applyUpdateCellsOperation({ model, operation, outputFormat: resolvedOutputFormat, utils });
      } else if (operation.type === 'add_row') {
        applyAddRowOperation({ model, operation });
      } else if (operation.type === 'sort_rows') {
        applySortRowsOperation({ model, operation });
      } else if (operation.type === 'reorder_rows') {
        applyReorderRowsOperation({ model, operation });
      }

      operationSummaries.push({
        type: operation.type,
        sheetName,
      });
    }
  }

  const finalWorkbook = { ...workbook, Sheets: {}, SheetNames: [] };
  for (const sheetName of sheetOrder) {
    const model = sheetModels.get(sheetName);
    if (!model) {
      continue;
    }
    finalWorkbook.SheetNames.push(sheetName);
    finalWorkbook.Sheets[sheetName] = buildWorksheetFromModel({ model, utils });
  }

  if (resolvedOutputFormat === 'csv' && finalWorkbook.SheetNames.length !== 1) {
    throw new Error('CSV output requires the transformed workbook to contain exactly one sheet');
  }

  let outputBuffer;
  let mimeType;
  if (resolvedOutputFormat === 'csv') {
    const firstSheet = finalWorkbook.Sheets[finalWorkbook.SheetNames[0]];
    outputBuffer = Buffer.from(utils.sheet_to_csv(firstSheet), 'utf8');
    mimeType = 'text/csv';
  } else {
    outputBuffer = Buffer.from(
      write(finalWorkbook, {
        type: 'buffer',
        bookType: 'xlsx',
        cellStyles: true,
        bookVBA: true,
      }),
    );
    mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  }

  const finalSheetSummaries = finalWorkbook.SheetNames.map((sheetName) => {
    const model = sheetModels.get(sheetName);
    return {
      sheetName,
      rowCount: model.dataRows.length,
      outputColumnCount: model.headers.length,
      columns: model.headers.filter(Boolean),
    };
  });

  return {
    buffer: outputBuffer,
    bytes: outputBuffer.length,
    mimeType,
    filename: buildOutputFilename(sourceFilename, resolvedOutputFormat),
    summary: {
      outputFormat: resolvedOutputFormat,
      sheetCount: finalWorkbook.SheetNames.length,
      sheets: sheetSummaries,
      finalSheets: finalSheetSummaries,
      operationsApplied: operationSummaries,
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
