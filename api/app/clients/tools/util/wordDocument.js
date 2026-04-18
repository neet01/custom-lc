const { tool } = require('@langchain/core/tools');
const {
  isWordDocumentTransformable,
} = require('~/server/services/Files/WordDocuments/transform');
const {
  inspectWordDocumentFile,
  transformWordDocumentFile,
} = require('~/server/services/Files/WordDocuments/service');
const { filterFilesByAgentAccess } = require('~/server/services/Files/permissions');
const { getConvoFiles, getFiles } = require('~/models');

const WORD_DOCUMENT_TOOL_NAME = 'word_document_transform';

const wordDocumentTransformJsonSchema = {
  type: 'object',
  properties: {
    action: {
      type: 'string',
      enum: ['inspect', 'transform'],
      description:
        'Use "inspect" to understand the attached Word document. Use "transform" to create a new downloadable .docx file in chat.',
    },
    file_id: {
      type: 'string',
      description:
        'The exact LibreChat file_id for the Word document to inspect or transform. Prefer this when available.',
    },
    file_name: {
      type: 'string',
      description:
        'The Word document filename when file_id is not known. Must match one of the attached .docx files.',
    },
    replaceText: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          find: { type: 'string' },
          replace: { type: 'string' },
        },
        required: ['find', 'replace'],
      },
      description:
        'For transform: exact text replacements to apply to the document body after extraction.',
    },
    redactPhrases: {
      type: 'array',
      items: { type: 'string' },
      description:
        'For transform: exact phrases to replace with the redaction text throughout the document.',
    },
    redactionText: {
      type: 'string',
      description: 'For transform: replacement text used when redacting phrases.',
    },
    prependText: {
      type: 'string',
      description: 'For transform: text to insert at the beginning of the generated document.',
    },
    appendText: {
      type: 'string',
      description: 'For transform: text to append to the end of the generated document.',
    },
    replacementText: {
      type: 'string',
      description:
        'For transform: full replacement body for the new document. Use this when rewriting the entire document.',
    },
    outputFilename: {
      type: 'string',
      description:
        'Optional output filename for the generated .docx file. The .docx extension is enforced automatically.',
    },
    maxPreviewParagraphs: {
      type: 'integer',
      minimum: 1,
      maximum: 10,
      description:
        'For inspect: number of paragraphs to preview from the document. Defaults to 5 and maxes out at 10.',
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

function formatWordDocumentContext(files) {
  if (!files.length) {
    return `- Note: The ${WORD_DOCUMENT_TOOL_NAME} tool is available, but no .docx files are attached. Ask the user to upload a Word document before using it.`;
  }

  const lines = [`- Note: Use the ${WORD_DOCUMENT_TOOL_NAME} tool for these attached Word documents:`];
  for (const file of files) {
    lines.push(`\t- ${file.filename} (file_id: ${file.file_id})`);
  }
  return lines.join('\n');
}

async function primeFiles({ req, agentId }) {
  const requestFiles = Array.isArray(req?.body?.files)
    ? req.body.files.filter((file) => isWordDocumentTransformable(file?.type))
    : [];

  let conversationFiles = [];
  const conversationId = req?.body?.conversationId;
  if (conversationId) {
    const fileIds = (await getConvoFiles(conversationId)) ?? [];
    if (fileIds.length > 0) {
      const allFiles = (await getFiles({ file_id: { $in: fileIds } }, null, { text: 0 })) ?? [];
      const wordFiles = allFiles.filter((file) => isWordDocumentTransformable(file?.type));
      if (req?.user?.id && agentId) {
        conversationFiles = await filterFilesByAgentAccess({
          files: wordFiles,
          userId: req.user.id,
          role: req.user.role,
          agentId,
        });
      } else {
        conversationFiles = wordFiles;
      }
    }
  }

  const files = uniqueFiles(requestFiles.concat(conversationFiles));
  return {
    files,
    toolContext: formatWordDocumentContext(files),
  };
}

function resolveTargetFile(files, { file_id, file_name }) {
  if (file_id) {
    const file = files.find((candidate) => candidate.file_id === file_id);
    if (file) {
      return file;
    }
    throw new Error(`Word document file_id "${file_id}" is not available in this conversation`);
  }

  if (file_name) {
    const normalized = String(file_name).trim().toLowerCase();
    const file = files.find((candidate) => candidate.filename?.trim().toLowerCase() === normalized);
    if (file) {
      return file;
    }
    throw new Error(`Word document "${file_name}" is not available in this conversation`);
  }

  if (files.length === 1) {
    return files[0];
  }

  const availableFiles = files.map((file) => `"${file.filename}"`).join(', ');
  throw new Error(
    `Multiple Word documents are attached. Specify file_id or file_name. Available files: ${availableFiles}`,
  );
}

function formatInspectionResult(file, inspection) {
  const lines = [
    `Document "${file.filename}" contains ${inspection.paragraphCount} paragraph(s) and ${inspection.wordCount} word(s).`,
  ];

  if (inspection.previewParagraphs.length > 0) {
    lines.push(`Preview paragraphs: ${JSON.stringify(inspection.previewParagraphs)}`);
  }

  return lines.join('\n');
}

async function createWordDocumentTool({ req, res, files }) {
  return tool(
    async ({
      action,
      file_id,
      file_name,
      replaceText,
      redactPhrases,
      redactionText,
      prependText,
      appendText,
      replacementText,
      outputFilename,
      maxPreviewParagraphs,
    }) => {
      if (files.length === 0) {
        return [
          'No Word documents are attached. Ask the user to upload a .docx file first.',
          undefined,
        ];
      }

      const sourceFile = resolveTargetFile(files, { file_id, file_name });

      if (action === 'inspect') {
        const inspection = await inspectWordDocumentFile({
          req,
          res,
          sourceFile,
          maxPreviewParagraphs,
        });

        return [
          formatInspectionResult(sourceFile, inspection),
          {
            [WORD_DOCUMENT_TOOL_NAME]: {
              sourceFileId: sourceFile.file_id,
              inspection,
            },
          },
        ];
      }

      const transformed = await transformWordDocumentFile({
        req,
        res,
        sourceFile,
        replaceText,
        redactPhrases,
        redactionText,
        prependText,
        appendText,
        replacementText,
        outputFilename,
        conversationId: req?.body?.conversationId,
      });

      return [
        `Created "${transformed.file.filename}" from "${sourceFile.filename}".`,
        {
          files: [transformed.file],
          [WORD_DOCUMENT_TOOL_NAME]: {
            sourceFileId: sourceFile.file_id,
            outputFileId: transformed.file.file_id,
            summary: transformed.summary,
          },
        },
      ];
    },
    {
      name: WORD_DOCUMENT_TOOL_NAME,
      description:
        'Inspect attached .docx files and create downloadable Word documents directly in chat. Use this to preview a document, redact phrases, replace text, or generate a rewritten .docx file.',
      schema: wordDocumentTransformJsonSchema,
      responseFormat: 'content_and_artifact',
    },
  );
}

module.exports = {
  WORD_DOCUMENT_TOOL_NAME,
  createWordDocumentTool,
  primeFiles,
};
