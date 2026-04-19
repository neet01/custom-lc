const JSZip = require('jszip');
const mammoth = require('mammoth');
const {
  DOCX_MIME_TYPE,
  inspectWordDocumentBuffer,
  isWordDocumentTransformable,
  transformWordDocumentBuffer,
} = require('./transform');

async function createStructuredDocxBuffer() {
  const zip = new JSZip();
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
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:b/>
        </w:rPr>
        <w:t xml:space="preserve">Budget memo</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t xml:space="preserve">Revenue is 500.</w:t>
      </w:r>
    </w:p>
    <w:tbl>
      <w:tr>
        <w:tc>
          <w:p>
            <w:r>
              <w:t xml:space="preserve">Cell text</w:t>
            </w:r>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
    </w:sectPr>
  </w:body>
</w:document>`,
  );

  return zip.generateAsync({ type: 'nodebuffer' });
}

describe('Word document transform service', () => {
  it('inspects a docx buffer and previews paragraphs', async () => {
    const inputBuffer = await createStructuredDocxBuffer();

    const inspection = await inspectWordDocumentBuffer({
      buffer: inputBuffer,
      sourceFilename: 'update.docx',
      maxPreviewParagraphs: 2,
    });

    expect(inspection.filename).toBe('update.docx');
    expect(inspection.paragraphCount).toBe(3);
    expect(inspection.previewParagraphs).toEqual(['Budget memo', 'Revenue is 500.']);
  });

  it('preserves document structure while updating text and adding a blurb', async () => {
    const inputBuffer = await createStructuredDocxBuffer();

    const result = await transformWordDocumentBuffer({
      buffer: inputBuffer,
      sourceFilename: 'memo.docx',
      replaceText: [{ find: '500', replace: '650' }],
      appendText: 'Prepared for the finance team.',
    });

    expect(result.mimeType).toBe(DOCX_MIME_TYPE);
    expect(result.filename).toBe('memo-transformed.docx');
    expect(result.summary.replacements[0].occurrences).toBe(1);
    expect(result.summary.appendedText).toBe(true);

    const extracted = await mammoth.extractRawText({ buffer: result.buffer });
    expect(extracted.value).toContain('Revenue is 650.');
    expect(extracted.value).toContain('Prepared for the finance team.');
    expect(extracted.value).toContain('Cell text');

    const outputZip = await JSZip.loadAsync(result.buffer);
    const documentXml = await outputZip.file('word/document.xml').async('string');
    expect(documentXml).toContain('<w:tbl>');
    expect(documentXml).toContain('w:pStyle w:val="Heading1"');
    expect(documentXml).toContain('<w:b>');
  });

  it('supports full replacement text while preserving section properties', async () => {
    const inputBuffer = await createStructuredDocxBuffer();

    const result = await transformWordDocumentBuffer({
      buffer: inputBuffer,
      sourceFilename: 'memo.docx',
      replacementText: 'Completely rewritten memo',
    });

    const extracted = await mammoth.extractRawText({ buffer: result.buffer });
    expect(extracted.value).toContain('Completely rewritten memo');

    const outputZip = await JSZip.loadAsync(result.buffer);
    const documentXml = await outputZip.file('word/document.xml').async('string');
    expect(documentXml).toContain('<w:sectPr>');
  });

  it('recognizes supported Word document MIME types', () => {
    expect(isWordDocumentTransformable(DOCX_MIME_TYPE)).toBe(true);
    expect(isWordDocumentTransformable('application/pdf')).toBe(false);
  });
});
