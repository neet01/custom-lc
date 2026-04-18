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
        conversationId: req?.body?.conversationId,
      });

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
      description: `Inspect attached spreadsheets and create downloadable spreadsheet exports directly in chat. Use "${SPREADSHEET_TOOL_NAME}" when the user wants to understand workbook structure, inspect sheet columns, remove columns, keep only selected columns, or redact sensitive spreadsheet data. For uncertain workbooks, inspect first and transform second.`,
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
