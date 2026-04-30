const { tool } = require('@langchain/core/tools');
const { isSpreadsheetTransformable } = require('~/server/services/Files/Spreadsheets/transform');
const {
  inspectSpreadsheetFile,
  transformSpreadsheetFile,
} = require('~/server/services/Files/Spreadsheets/service');
const { filterFilesByAgentAccess } = require('~/server/services/Files/permissions');
const { getConvoFiles, getFiles } = require('~/models');

const SPREADSHEET_TOOL_NAME = 'spreadsheet_transform';

const spreadsheetTransformJsonSchema = {
  type: 'object',
  properties: {
    action: {
      type: 'string',
      enum: ['inspect', 'transform'],
      description:
        'Use "inspect" to understand workbook sheets, columns, and preview rows. Use "transform" to create a new downloadable spreadsheet file.',
    },
    file_id: {
      type: 'string',
      description:
        'The exact LibreChat file_id for the spreadsheet to inspect or transform. Prefer this when available.',
    },
    file_name: {
      type: 'string',
      description:
        'The spreadsheet filename when file_id is not known. Must match one of the attached spreadsheet files.',
    },
    keepColumns: {
      type: 'array',
      items: { type: 'string' },
      description:
        'For transform: keep only these column headers, preserving order from the source sheet.',
    },
    removeColumns: {
      type: 'array',
      items: { type: 'string' },
      description: 'For transform: remove these column headers from the output workbook.',
    },
    redactColumns: {
      type: 'array',
      items: { type: 'string' },
      description:
        'For transform: replace non-empty cell values in these columns with a redaction label.',
    },
    redactionText: {
      type: 'string',
      description: 'For transform: custom replacement text for redacted cells.',
    },
    sheetNames: {
      type: 'array',
      items: { type: 'string' },
      description:
        'Optional list of sheet names to target. Leave empty to inspect or transform every sheet.',
    },
    outputFormat: {
      type: 'string',
      enum: ['xlsx', 'csv'],
      description:
        'For transform: output workbook format. CSV requires exactly one selected sheet.',
    },
    maxPreviewRows: {
      type: 'integer',
      minimum: 1,
      maximum: 10,
      description:
        'For inspect: number of data rows to preview from each sheet. Defaults to 5 and maxes out at 10.',
    },
    operations: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          type: {
            type: 'string',
            enum: [
              'add_column',
              'add_row',
              'update_cells',
              'add_totals_row',
              'sort_rows',
              'reorder_rows',
              'merge_sheets',
              'split_sheet',
            ],
          },
          sheetName: {
            type: 'string',
            description:
              'Optional target sheet for row or column operations. If omitted, the operation applies to the selected sheets.',
          },
          columnName: {
            type: 'string',
            description:
              'Column header to create, update, or sort by, depending on the operation type.',
          },
          beforeColumn: {
            type: 'string',
            description: 'For add_column: insert before this existing column.',
          },
          afterColumn: {
            type: 'string',
            description: 'For add_column: insert after this existing column.',
          },
          position: {
            type: 'string',
            enum: ['start', 'end'],
            description:
              'For add_column or add_row: coarse insert position when before/after/index are not provided.',
          },
          index: {
            type: 'integer',
            minimum: 0,
            description:
              'Optional zero-based insert position for add_column, or one-based insertion index for add_row.',
          },
          defaultValue: {
            type: ['string', 'number', 'boolean', 'null'],
            description:
              'For add_column: default scalar value to populate when expression or formula are not used.',
          },
          value: {
            type: ['string', 'number', 'boolean', 'null'],
            description:
              'For update_cells: replacement scalar value when expression or formula are not used.',
          },
          expression: {
            type: 'string',
            description:
              'Calculation expression using {{Column Name}} placeholders, for example "{{Revenue}} - {{Expense}}".',
          },
          formula: {
            type: 'string',
            description:
              'Excel-style formula template using {{Column Name}} placeholders, for example "=SUM({{Q1}}, {{Q2}})". Best used with xlsx output.',
          },
          values: {
            type: 'object',
            additionalProperties: {
              type: ['string', 'number', 'boolean', 'null'],
            },
            description:
              'For add_row: object mapping column headers to scalar values for the new row.',
          },
          inheritFormulas: {
            type: 'boolean',
            description:
              'For add_row: when true, blank columns inherit translated formulas from the template row above. Defaults to true.',
          },
          rowNumber: {
            type: 'integer',
            minimum: 1,
            description:
              'For update_cells: one-based data row number under the header row to target.',
          },
          rowMatch: {
            type: 'object',
            additionalProperties: {
              type: ['string', 'number', 'boolean', 'null'],
            },
            description:
              'For update_cells: match rows by exact column values before applying the update.',
          },
          columns: {
            type: 'array',
            items: {
              type: 'object',
              properties: {
                columnName: { type: 'string' },
                direction: { type: 'string', enum: ['asc', 'desc'] },
                numeric: { type: 'boolean' },
                function: {
                  type: 'string',
                  enum: ['sum', 'average', 'min', 'max', 'count', 'counta'],
                },
              },
              required: ['columnName'],
            },
            description:
              'For sort_rows: multi-column sort specification in priority order. For add_totals_row: columns to aggregate, optionally with a function.',
          },
          direction: {
            type: 'string',
            enum: ['asc', 'desc'],
            description:
              'For sort_rows with a single columnName: sort direction, defaulting to asc.',
          },
          numeric: {
            type: 'boolean',
            description:
              'For sort_rows: force numeric comparison for the target sort column.',
          },
          labelColumn: {
            type: 'string',
            description:
              'For add_totals_row: the column header that should receive the totals label. Defaults to the first header column.',
          },
          label: {
            type: 'string',
            description:
              'For add_totals_row: label to write into the totals row, such as "Total", "Average", or "Runway Summary".',
          },
          orderedRowNumbers: {
            type: 'array',
            items: {
              type: 'integer',
              minimum: 1,
            },
            description:
              'For reorder_rows: explicit one-based row order for the selected data rows.',
          },
          appendRemaining: {
            type: 'boolean',
            description:
              'For reorder_rows: append rows not listed in orderedRowNumbers after the explicit ordering. Defaults to true.',
          },
          sourceSheets: {
            type: 'array',
            items: { type: 'string' },
            description:
              'For merge_sheets: source sheet names to combine into one sheet.',
          },
          outputSheetName: {
            type: 'string',
            description:
              'For merge_sheets: output sheet name for the combined result.',
          },
          includeSourceSheetColumn: {
            type: 'boolean',
            description:
              'For merge_sheets: include a Source Sheet column in the merged output. Defaults to true.',
          },
          preserveSourceSheets: {
            type: 'boolean',
            description:
              'For merge_sheets: keep original sheets after creating the merged sheet. Defaults to true.',
          },
          sourceSheetName: {
            type: 'string',
            description:
              'For split_sheet: the source sheet to split into multiple sheets.',
          },
          byColumn: {
            type: 'string',
            description:
              'For split_sheet: split rows into separate sheets by the distinct values in this column.',
          },
          outputSheetPrefix: {
            type: 'string',
            description:
              'For split_sheet: prefix used when naming the generated sheets.',
          },
          preserveSourceSheet: {
            type: 'boolean',
            description:
              'For split_sheet: keep the original source sheet after splitting. Defaults to true.',
          },
        },
        required: ['type'],
      },
      description:
        'Advanced transformation operations to add columns or rows, update cells, calculate values, add totals rows, reorder rows, sort rows, merge sheets, or split a sheet into multiple sheets.',
    },
  },
  required: ['action'],
};

function uniqueFiles(files) {
  const seen = new Set();
  return files.filter((file) => {
    if (!file?.file_id || seen.has(file.file_id)) {
      return false;
    }
    seen.add(file.file_id);
    return true;
  });
}

function registerGeneratedFile(files, req, file) {
  if (!file?.file_id) {
    return;
  }

  if (!files.some((candidate) => candidate.file_id === file.file_id)) {
    files.push(file);
  }

  if (!req?.body) {
    return;
  }

  if (!Array.isArray(req.body.files)) {
    req.body.files = [];
  }

  if (!req.body.files.some((candidate) => candidate?.file_id === file.file_id)) {
    req.body.files.push(file);
  }
}

function formatSpreadsheetContext(files) {
  if (!files.length) {
    return `- Note: The ${SPREADSHEET_TOOL_NAME} tool is available, but no spreadsheet files are attached. Ask the user to upload a spreadsheet before using it.`;
  }

  const lines = [`- Note: Use the ${SPREADSHEET_TOOL_NAME} tool for these attached spreadsheets:`];
  for (const file of files) {
    lines.push(`\t- ${file.filename} (file_id: ${file.file_id})`);
  }
  return lines.join('\n');
}

async function primeFiles({ req, agentId }) {
  const requestFiles = Array.isArray(req?.body?.files)
    ? req.body.files.filter((file) => isSpreadsheetTransformable(file?.type))
    : [];

  let conversationFiles = [];
  const conversationId = req?.body?.conversationId;
  if (conversationId) {
    const fileIds = (await getConvoFiles(conversationId)) ?? [];
    if (fileIds.length > 0) {
      const allFiles = (await getFiles({ file_id: { $in: fileIds } }, null, { text: 0 })) ?? [];
      const spreadsheetFiles = allFiles.filter((file) => isSpreadsheetTransformable(file?.type));
      if (req?.user?.id && agentId) {
        conversationFiles = await filterFilesByAgentAccess({
          files: spreadsheetFiles,
          userId: req.user.id,
          role: req.user.role,
          agentId,
        });
      } else {
        conversationFiles = spreadsheetFiles;
      }
    }
  }

  const files = uniqueFiles(requestFiles.concat(conversationFiles));
  return {
    files,
    toolContext: formatSpreadsheetContext(files),
  };
}

function resolveTargetFile(files, { file_id, file_name }) {
  if (file_id) {
    const file = files.find((candidate) => candidate.file_id === file_id);
    if (file) {
      return file;
    }
    throw new Error(`Spreadsheet file_id "${file_id}" is not available in this conversation`);
  }

  if (file_name) {
    const normalized = String(file_name).trim().toLowerCase();
    const file = files.find((candidate) => candidate.filename?.trim().toLowerCase() === normalized);
    if (file) {
      return file;
    }
    throw new Error(`Spreadsheet file "${file_name}" is not available in this conversation`);
  }

  if (files.length === 1) {
    return files[0];
  }

  const availableFiles = files.map((file) => `"${file.filename}"`).join(', ');
  throw new Error(
    `Multiple spreadsheets are attached. Specify file_id or file_name. Available files: ${availableFiles}`,
  );
}

function formatInspectionResult(file, inspection) {
  const lines = [
    `Workbook "${file.filename}" contains ${inspection.sheetCount} sheet(s).`,
  ];

  for (const sheet of inspection.sheets) {
    lines.push(
      `Sheet "${sheet.sheetName}": ${sheet.rowCount} data row(s), ${sheet.columnCount} column(s): ${sheet.columns.join(', ') || 'no headers detected'}.`,
    );
    if (sheet.previewRows.length > 0) {
      lines.push(`Preview rows: ${JSON.stringify(sheet.previewRows)}`);
    }
  }

  return lines.join('\n');
}

async function createSpreadsheetTool({ req, res, files }) {
  return tool(
    async ({
      action,
      file_id,
      file_name,
      removeColumns,
      keepColumns,
      redactColumns,
      redactionText,
      sheetNames,
      outputFormat,
      maxPreviewRows,
      operations,
    }) => {
      if (files.length === 0) {
        return [
          'No spreadsheet files are attached. Ask the user to upload a spreadsheet first.',
          undefined,
        ];
      }

      const sourceFile = resolveTargetFile(files, { file_id, file_name });

      if (action === 'inspect') {
        const inspection = await inspectSpreadsheetFile({
          req,
          res,
          sourceFile,
          maxPreviewRows,
        });

        return [
          formatInspectionResult(sourceFile, inspection),
          {
            [SPREADSHEET_TOOL_NAME]: {
              sourceFileId: sourceFile.file_id,
              inspection,
            },
          },
        ];
      }

      const transformed = await transformSpreadsheetFile({
        req,
        res,
        sourceFile,
        removeColumns,
        keepColumns,
        redactColumns,
        redactionText,
        sheetNames,
        outputFormat,
        operations,
        conversationId: req?.body?.conversationId,
      });

      registerGeneratedFile(files, req, transformed.file);

      return [
        `Created "${transformed.file.filename}" from "${sourceFile.filename}".`,
        {
          files: [transformed.file],
          [SPREADSHEET_TOOL_NAME]: {
            sourceFileId: sourceFile.file_id,
            outputFileId: transformed.file.file_id,
            summary: transformed.summary,
          },
        },
      ];
    },
    {
      name: SPREADSHEET_TOOL_NAME,
      responseFormat: 'content_and_artifact',
      description: `Inspect attached spreadsheets and create downloadable spreadsheet exports directly in chat. Use "${SPREADSHEET_TOOL_NAME}" when the user wants to inspect workbook structure, clean columns, add or update rows and cells, calculate values, sort data, or merge and split sheets. For uncertain workbooks, inspect first and transform second.`,
      schema: spreadsheetTransformJsonSchema,
    },
  );
}

module.exports = {
  SPREADSHEET_TOOL_NAME,
  spreadsheetTransformJsonSchema,
  createSpreadsheetTool,
  formatSpreadsheetContext,
  primeFiles,
};
