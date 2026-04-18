const path = require('path');
const JSZip = require('jszip');
const mammoth = require('mammoth');

const DOCX_MIME_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

function isWordDocumentTransformable(mimetype) {
  return mimetype === DOCX_MIME_TYPE;
}

function normalizeTextInput(value) {
  if (typeof value !== 'string') {
    return undefined;
  }

  return value.replace(/\r\n/g, '\n');
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
      replace: typeof operation?.replace === 'string' ? operation.replace : '',
    }))
    .filter((operation) => operation.find.length > 0);
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

function xmlEscape(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function createParagraphXml(paragraph) {
  if (!paragraph) {
    return '<w:p/>';
  }

  return `<w:p><w:r><w:t xml:space="preserve">${xmlEscape(paragraph)}</w:t></w:r></w:p>`;
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

async function buildDocxBufferFromText(text) {
  const paragraphs = String(text ?? '')
    .replace(/\r\n/g, '\n')
    .split('\n');
  const safeParagraphs = paragraphs.length > 0 ? paragraphs : [''];
  const bodyXml = safeParagraphs.map(createParagraphXml).join('');
  const now = new Date().toISOString();

  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`,
  );
  zip.folder('_rels').file(
    '.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`,
  );
  zip.folder('docProps').file(
    'core.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>LibreChat Document</dc:title>
  <dc:creator>LibreChat</dc:creator>
  <cp:lastModifiedBy>LibreChat</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
</cp:coreProperties>`,
  );
  zip.folder('docProps').file(
    'app.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>LibreChat</Application>
</Properties>`,
  );
  zip.folder('word').folder('_rels').file(
    'document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`,
  );
  zip.folder('word').file(
    'document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${bodyXml}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`,
  );

  return zip.generateAsync({ type: 'nodebuffer' });
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

  if (
    normalizedReplacementText === undefined &&
    normalizedReplaceText.length === 0 &&
    normalizedRedactPhrases.length === 0 &&
    !normalizedPrependText?.trim() &&
    !normalizedAppendText?.trim()
  ) {
    throw new Error('At least one Word document transformation must be requested');
  }

  const sourceText = await extractRawTextFromDocxBuffer(buffer);
  if (!sourceText.trim() && normalizedReplacementText === undefined) {
    throw new Error('No text found in Word document');
  }

  let outputText = normalizedReplacementText !== undefined ? normalizedReplacementText : sourceText;

  const replacementSummary = [];
  for (const operation of normalizedReplaceText) {
    const occurrences = countOccurrences(outputText, operation.find);
    if (occurrences > 0) {
      outputText = outputText.split(operation.find).join(operation.replace);
    }
    replacementSummary.push({
      find: operation.find,
      replace: operation.replace,
      occurrences,
    });
  }

  const redactionSummary = [];
  for (const phrase of normalizedRedactPhrases) {
    const occurrences = countOccurrences(outputText, phrase);
    if (occurrences > 0) {
      outputText = outputText.split(phrase).join(redactionText);
    }
    redactionSummary.push({
      phrase,
      occurrences,
    });
  }

  if (normalizedPrependText?.trim()) {
    outputText = `${normalizedPrependText}\n\n${outputText}`;
  }

  if (normalizedAppendText?.trim()) {
    outputText = `${outputText}\n\n${normalizedAppendText}`;
  }

  const changed =
    normalizedReplacementText !== undefined ||
    replacementSummary.some((item) => item.occurrences > 0) ||
    redactionSummary.some((item) => item.occurrences > 0) ||
    Boolean(normalizedPrependText?.trim()) ||
    Boolean(normalizedAppendText?.trim());

  if (!changed) {
    throw new Error('Requested transformations did not match any text in the Word document');
  }

  const outputBuffer = await buildDocxBufferFromText(outputText);
  const outputParagraphs = parseParagraphs(outputText);

  return {
    buffer: outputBuffer,
    bytes: outputBuffer.length,
    mimeType: DOCX_MIME_TYPE,
    filename: buildOutputFilename(sourceFilename, outputFilename),
    summary: {
      paragraphCount: outputParagraphs.length,
      wordCount: countWords(outputText),
      replacements: replacementSummary,
      redactions: redactionSummary,
      usedReplacementText: normalizedReplacementText !== undefined,
      prependedText: Boolean(normalizedPrependText?.trim()),
      appendedText: Boolean(normalizedAppendText?.trim()),
    },
  };
}

module.exports = {
  DOCX_MIME_TYPE,
  inspectWordDocumentBuffer,
  isWordDocumentTransformable,
  transformWordDocumentBuffer,
};
