const path = require('path');
const JSZip = require('jszip');
const mammoth = require('mammoth');
const { XMLParser, XMLBuilder, XMLValidator } = require('fast-xml-parser');

const DOCX_MIME_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const REQUIRED_DOCX_PARTS = ['[Content_Types].xml', '_rels/.rels', 'word/document.xml'];

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  preserveOrder: true,
  trimValues: false,
  processEntities: false,
});

const xmlBuilder = new XMLBuilder({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  preserveOrder: true,
  suppressEmptyNode: false,
  format: false,
  processEntities: false,
});

function isWordDocumentTransformable(mimetype) {
  return mimetype === DOCX_MIME_TYPE;
}

function normalizeTextInput(value) {
  if (typeof value !== 'string') {
    return undefined;
  }

  return sanitizeXmlText(value).replace(/\r\n/g, '\n');
}

function normalizeStringList(values) {
  if (!Array.isArray(values)) {
    return [];
  }

  return [...new Set(values.map((value) => String(value ?? '').trim()).filter(Boolean))];
}

function normalizeReplacementOperations(replaceText) {
  if (!Array.isArray(replaceText)) {
    return [];
  }

  return replaceText
    .map((operation) => ({
      find: String(operation?.find ?? '').trim(),
      replace: typeof operation?.replace === 'string' ? sanitizeXmlText(operation.replace) : '',
    }))
    .filter((operation) => operation.find.length > 0);
}

function sanitizeXmlText(value) {
  return Array.from(String(value ?? ''))
    .filter((char) => {
      const codePoint = char.codePointAt(0);
      return (
        codePoint === 0x09 ||
        codePoint === 0x0a ||
        codePoint === 0x0d ||
        (codePoint >= 0x20 && codePoint <= 0xd7ff) ||
        (codePoint >= 0xe000 && codePoint <= 0xfffd) ||
        (codePoint >= 0x10000 && codePoint <= 0x10ffff)
      );
    })
    .join('');
}

function parseParagraphs(text) {
  return String(text ?? '')
    .replace(/\r\n/g, '\n')
    .split(/\n+/)
    .map((paragraph) => paragraph.trim())
    .filter(Boolean);
}

function countWords(text) {
  return String(text ?? '').trim().match(/\S+/g)?.length ?? 0;
}

function countOccurrences(text, searchValue) {
  if (!searchValue) {
    return 0;
  }

  return String(text).split(searchValue).length - 1;
}

function buildOutputFilename(sourceFilename, outputFilename) {
  if (typeof outputFilename === 'string' && outputFilename.trim()) {
    const parsed = path.parse(outputFilename.trim());
    const safeBase = parsed.name || 'document';
    return `${safeBase}.docx`;
  }

  const parsed = path.parse(sourceFilename || 'document');
  const safeBase = parsed.name || 'document';
  return `${safeBase}-transformed.docx`;
}

async function extractRawTextFromDocxBuffer(buffer) {
  const { value } = await mammoth.extractRawText({ buffer });
  return value ?? '';
}

function cloneNode(node) {
  return JSON.parse(JSON.stringify(node));
}

function getNodeChildren(node, key) {
  return node?.[key] ?? null;
}

function getTextValue(textNode) {
  if (!textNode) {
    return '';
  }

  const textEntries = getNodeChildren(textNode, 'w:t');
  if (!Array.isArray(textEntries) || textEntries.length === 0) {
    return '';
  }

  return String(textEntries[0]['#text'] ?? '');
}

function collectNodeText(nodes) {
  if (!Array.isArray(nodes)) {
    return '';
  }

  let text = '';
  for (const node of nodes) {
    if (!node || typeof node !== 'object') {
      continue;
    }

    if (node['w:t']) {
      text += getTextValue(node);
      continue;
    }

    if (node['w:tab']) {
      text += '\t';
      continue;
    }

    if (node['w:br'] || node['w:cr']) {
      text += '\n';
      continue;
    }

    for (const key of Object.keys(node)) {
      if (key === ':@') {
        continue;
      }

      const children = getNodeChildren(node, key);
      if (Array.isArray(children)) {
        text += collectNodeText(children);
      }
    }
  }

  return text;
}

function getParagraphText(paragraphNode) {
  return collectNodeText(getNodeChildren(paragraphNode, 'w:p'));
}

function createTextNodes(text) {
  const normalized = sanitizeXmlText(text).replace(/\r\n/g, '\n');
  const parts = normalized.split('\n');
  const nodes = [];

  for (let i = 0; i < parts.length; i += 1) {
    nodes.push({
      'w:t': [{ '#text': parts[i] }],
      ':@': { '@_xml:space': 'preserve' },
    });

    if (i < parts.length - 1) {
      nodes.push({ 'w:br': [] });
    }
  }

  return nodes;
}

function getParagraphProperties(paragraphChildren) {
  return paragraphChildren.filter((child) => child['w:pPr']).map(cloneNode);
}

function getFirstRunProperties(paragraphChildren) {
  for (const child of paragraphChildren) {
    const runChildren = getNodeChildren(child, 'w:r');
    if (!Array.isArray(runChildren)) {
      continue;
    }

    const runProperties = runChildren.find((runChild) => runChild['w:rPr']);
    if (runProperties) {
      return cloneNode(runProperties);
    }
  }

  return null;
}

function createRunNode(text, runProperties) {
  const runChildren = [];
  if (runProperties) {
    runChildren.push(cloneNode(runProperties));
  }
  runChildren.push(...createTextNodes(text));
  return { 'w:r': runChildren };
}

function replaceParagraphText(paragraphNode, text) {
  const paragraphChildren = getNodeChildren(paragraphNode, 'w:p');
  const nextChildren = [
    ...getParagraphProperties(paragraphChildren),
    createRunNode(text, getFirstRunProperties(paragraphChildren)),
  ];
  paragraphNode['w:p'] = nextChildren;
}

function createParagraphNode(text) {
  return {
    'w:p': [createRunNode(text, null)],
  };
}

function normalizeParagraphInsertions(text) {
  if (!text?.trim()) {
    return [];
  }

  return String(text)
    .replace(/\r\n/g, '\n')
    .split('\n')
    .map((paragraph) => createParagraphNode(paragraph));
}

function findEntry(nodes, key) {
  return Array.isArray(nodes) ? nodes.find((node) => node[key]) : undefined;
}

function transformParagraphs(nodes, state) {
  if (!Array.isArray(nodes)) {
    return;
  }

  for (const node of nodes) {
    if (!node || typeof node !== 'object') {
      continue;
    }

    if (node['w:p']) {
      const originalText = getParagraphText(node);
      let updatedText = originalText;

      for (const operation of state.replacements) {
        const occurrences = countOccurrences(updatedText, operation.find);
        if (occurrences > 0) {
          updatedText = updatedText.split(operation.find).join(operation.replace);
          operation.occurrences += occurrences;
        }
      }

      for (const operation of state.redactions) {
        const occurrences = countOccurrences(updatedText, operation.phrase);
        if (occurrences > 0) {
          updatedText = updatedText.split(operation.phrase).join(state.redactionText);
          operation.occurrences += occurrences;
        }
      }

      if (updatedText !== originalText) {
        replaceParagraphText(node, updatedText);
        state.changed = true;
      }

      continue;
    }

    for (const key of Object.keys(node)) {
      if (key === ':@') {
        continue;
      }

      const children = getNodeChildren(node, key);
      if (Array.isArray(children)) {
        transformParagraphs(children, state);
      }
    }
  }
}

function replaceBodyWithParagraphs(bodyEntry, paragraphs) {
  const bodyChildren = getNodeChildren(bodyEntry, 'w:body');
  const sectionProps = bodyChildren.filter((child) => child['w:sectPr']).map(cloneNode);
  bodyEntry['w:body'] = [...paragraphs, ...sectionProps];
}

function insertParagraphsIntoBody(bodyEntry, paragraphs, position) {
  if (!paragraphs.length) {
    return;
  }

  const bodyChildren = getNodeChildren(bodyEntry, 'w:body');
  const sectionIndex = bodyChildren.findIndex((child) => child['w:sectPr']);
  const insertIndex =
    position === 'prepend'
      ? 0
      : sectionIndex >= 0
        ? sectionIndex
        : bodyChildren.length;

  bodyChildren.splice(insertIndex, 0, ...paragraphs);
}

async function inspectWordDocumentBuffer({
  buffer,
  sourceFilename,
  maxPreviewParagraphs = 5,
}) {
  if (!Buffer.isBuffer(buffer) || buffer.length === 0) {
    throw new Error('Word document buffer is required');
  }

  const previewCount = Math.min(Math.max(Number(maxPreviewParagraphs) || 5, 1), 10);
  const rawText = await extractRawTextFromDocxBuffer(buffer);
  const paragraphs = parseParagraphs(rawText);

  if (paragraphs.length === 0) {
    throw new Error('No text found in Word document');
  }

  return {
    filename: sourceFilename,
    paragraphCount: paragraphs.length,
    wordCount: countWords(rawText),
    previewParagraphs: paragraphs.slice(0, previewCount),
  };
}

function formatXmlValidationError(validation) {
  if (validation === true) {
    return undefined;
  }

  const error = validation?.err;
  if (!error) {
    return 'XML is invalid';
  }

  const location = [
    error.line ? `line ${error.line}` : null,
    error.col ? `column ${error.col}` : null,
  ]
    .filter(Boolean)
    .join(', ');
  return [error.msg, location].filter(Boolean).join(' at ');
}

async function validateXmlPart(zip, partName) {
  const xml = await zip.file(partName)?.async('string');
  if (!xml) {
    throw new Error(`Generated Word document is missing required part: ${partName}`);
  }

  const validation = XMLValidator.validate(xml);
  const validationError = formatXmlValidationError(validation);
  if (validationError) {
    throw new Error(`Generated Word document contains invalid XML in ${partName}: ${validationError}`);
  }

  return xml;
}

function validateDocxEntryNames(zip) {
  for (const entryName of Object.keys(zip.files)) {
    const pathSegments = entryName.split('/');
    if (entryName.startsWith('/') || pathSegments.includes('..') || entryName.includes('\0')) {
      throw new Error(`Generated Word document contains an unsafe package entry: ${entryName}`);
    }
  }
}

async function validateWordDocumentBuffer({ buffer }) {
  if (!Buffer.isBuffer(buffer) || buffer.length === 0) {
    throw new Error('Generated Word document buffer is empty');
  }

  let zip;
  try {
    zip = await JSZip.loadAsync(buffer);
  } catch (error) {
    throw new Error(`Generated Word document is not a readable DOCX package: ${error.message}`);
  }

  validateDocxEntryNames(zip);

  for (const partName of REQUIRED_DOCX_PARTS) {
    if (!zip.file(partName)) {
      throw new Error(`Generated Word document is missing required part: ${partName}`);
    }
  }

  await validateXmlPart(zip, '[Content_Types].xml');
  await validateXmlPart(zip, '_rels/.rels');
  const documentXml = await validateXmlPart(zip, 'word/document.xml');

  const documentTree = xmlParser.parse(documentXml);
  const documentEntry = findEntry(documentTree, 'w:document');
  const bodyEntry = findEntry(getNodeChildren(documentEntry, 'w:document'), 'w:body');
  if (!documentEntry || !bodyEntry) {
    throw new Error('Generated Word document is missing a valid document body');
  }

  try {
    await extractRawTextFromDocxBuffer(buffer);
  } catch (error) {
    throw new Error(`Generated Word document failed readback validation: ${error.message}`);
  }

  return {
    checkedParts: REQUIRED_DOCX_PARTS,
    readable: true,
  };
}

async function transformWordDocumentBuffer({
  buffer,
  sourceFilename,
  replaceText = [],
  redactPhrases = [],
  redactionText = '[REDACTED]',
  prependText,
  appendText,
  replacementText,
  outputFilename,
}) {
  if (!Buffer.isBuffer(buffer) || buffer.length === 0) {
    throw new Error('Word document buffer is required');
  }

  const normalizedReplacementText = normalizeTextInput(replacementText);
  const normalizedPrependText = normalizeTextInput(prependText);
  const normalizedAppendText = normalizeTextInput(appendText);
  const normalizedReplaceText = normalizeReplacementOperations(replaceText);
  const normalizedRedactPhrases = normalizeStringList(redactPhrases);
  const normalizedRedactionText = sanitizeXmlText(redactionText);

  if (
    normalizedReplacementText === undefined &&
    normalizedReplaceText.length === 0 &&
    normalizedRedactPhrases.length === 0 &&
    !normalizedPrependText?.trim() &&
    !normalizedAppendText?.trim()
  ) {
    throw new Error('At least one Word document transformation must be requested');
  }

  const zip = await JSZip.loadAsync(buffer);
  const documentXml = await zip.file('word/document.xml')?.async('string');
  if (!documentXml) {
    throw new Error('Word document is missing word/document.xml');
  }

  const documentTree = xmlParser.parse(documentXml);
  const documentEntry = findEntry(documentTree, 'w:document');
  const bodyEntry = findEntry(getNodeChildren(documentEntry, 'w:document'), 'w:body');

  if (!documentEntry || !bodyEntry) {
    throw new Error('Word document structure is invalid');
  }

  const state = {
    changed: false,
    redactionText: normalizedRedactionText,
    replacements: normalizedReplaceText.map((operation) => ({
      ...operation,
      occurrences: 0,
    })),
    redactions: normalizedRedactPhrases.map((phrase) => ({
      phrase,
      occurrences: 0,
    })),
  };

  if (normalizedReplacementText !== undefined) {
    replaceBodyWithParagraphs(bodyEntry, normalizeParagraphInsertions(normalizedReplacementText));
    state.changed = true;
  } else {
    transformParagraphs(getNodeChildren(bodyEntry, 'w:body'), state);
  }

  if (normalizedPrependText?.trim()) {
    insertParagraphsIntoBody(
      bodyEntry,
      normalizeParagraphInsertions(normalizedPrependText),
      'prepend',
    );
    state.changed = true;
  }

  if (normalizedAppendText?.trim()) {
    insertParagraphsIntoBody(
      bodyEntry,
      normalizeParagraphInsertions(normalizedAppendText),
      'append',
    );
    state.changed = true;
  }

  if (!state.changed) {
    throw new Error('Requested transformations did not match any text in the Word document');
  }

  zip.file('word/document.xml', xmlBuilder.build(documentTree));
  const outputBuffer = await zip.generateAsync({ type: 'nodebuffer' });
  const validation = await validateWordDocumentBuffer({ buffer: outputBuffer });
  const outputText = await extractRawTextFromDocxBuffer(outputBuffer);
  const outputParagraphs = parseParagraphs(outputText);

  return {
    buffer: outputBuffer,
    bytes: outputBuffer.length,
    mimeType: DOCX_MIME_TYPE,
    filename: buildOutputFilename(sourceFilename, outputFilename),
    summary: {
      paragraphCount: outputParagraphs.length,
      wordCount: countWords(outputText),
      replacements: state.replacements,
      redactions: state.redactions,
      usedReplacementText: normalizedReplacementText !== undefined,
      prependedText: Boolean(normalizedPrependText?.trim()),
      appendedText: Boolean(normalizedAppendText?.trim()),
      validation,
    },
  };
}

module.exports = {
  DOCX_MIME_TYPE,
  inspectWordDocumentBuffer,
  isWordDocumentTransformable,
  transformWordDocumentBuffer,
  validateWordDocumentBuffer,
};
