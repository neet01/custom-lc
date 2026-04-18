const JSZip = require('jszip');
const mammoth = require('mammoth');
const {
  DOCX_MIME_TYPE,
  inspectWordDocumentBuffer,
  isWordDocumentTransformable,
  transformWordDocumentBuffer,
} = require('./transform');

async function createDocxBufferFromText(text) {
  const zip = new JSZip();
  const bodyXml = String(text)
    .split('\n')
    .map((paragraph) => {
      if (!paragraph) {
        return '<w:p/>';
      }
      const escaped = paragraph
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
      return `<w:p><w:r><w:t xml:space="preserve">${escaped}</w:t></w:r></w:p>`;
    })
    .join('');

  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`,
  );
  zip.folder('_rels').file(
    '.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`,
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
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>`,
  );

  return zip.generateAsync({ type: 'nodebuffer' });
}

describe('Word document transform service', () => {
  it('inspects a docx buffer and previews paragraphs', async () => {
    const inputBuffer = await createDocxBufferFromText(
      ['Quarterly update', '', 'Revenue grew 18 percent year over year.', 'Next steps follow.'].join('\n'),
    );

    const inspection = await inspectWordDocumentBuffer({
      buffer: inputBuffer,
      sourceFilename: 'update.docx',
      maxPreviewParagraphs: 2,
    });

    expect(inspection.filename).toBe('update.docx');
    expect(inspection.paragraphCount).toBe(3);
    expect(inspection.previewParagraphs).toEqual([
      'Quarterly update',
      'Revenue grew 18 percent year over year.',
    ]);
  });

  it('rewrites and redacts a docx buffer into a new docx file', async () => {
    const inputBuffer = await createDocxBufferFromText(
      ['Budget memo', 'Revenue is 500.', 'Contact CFO at cfo@example.com.'].join('\n'),
    );

    const result = await transformWordDocumentBuffer({
      buffer: inputBuffer,
      sourceFilename: 'memo.docx',
      replaceText: [{ find: '500', replace: '650' }],
      redactPhrases: ['cfo@example.com'],
      appendText: 'Prepared for the finance team.',
    });

    expect(result.mimeType).toBe(DOCX_MIME_TYPE);
    expect(result.filename).toBe('memo-transformed.docx');
    expect(result.summary.replacements[0].occurrences).toBe(1);
    expect(result.summary.redactions[0].occurrences).toBe(1);

    const extracted = await mammoth.extractRawText({ buffer: result.buffer });
    expect(extracted.value).toContain('Revenue is 650.');
    expect(extracted.value).toContain('[REDACTED]');
    expect(extracted.value).toContain('Prepared for the finance team.');
  });

  it('recognizes supported Word document MIME types', () => {
    expect(isWordDocumentTransformable(DOCX_MIME_TYPE)).toBe(true);
    expect(isWordDocumentTransformable('application/pdf')).toBe(false);
  });
});
