import fs from 'fs';
import path from 'path';
import JSZip from 'jszip';
import { validateTemplate } from '../index';

const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
async function document(body: string, header?: string) {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    '<Types><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
  );
  zip.file(
    'word/document.xml',
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`
  );
  if (header)
    zip.file(
      'word/header1.xml',
      `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${header}</w:hdr>`
    );
  return zip.generateAsync({ type: 'nodebuffer' });
}

describe('validateTemplate', () => {
  it('accepts nested blocks without executing JavaScript', async () => {
    const template = await document(
      paragraph(
        '+++EXEC throw new Error("must not run")++++++FOR item IN missingData++++++IF unknown()++++++INS $item.name++++++END-IF++++++END-FOR item+++'
      )
    );
    expect(await validateTemplate(template)).toEqual({
      valid: true,
      diagnostics: [],
    });
  });

  it('collects errors with table positions and opening-block locations', async () => {
    const template = await document(
      paragraph('+++END-IF+++') +
        '<w:tbl><w:tr><w:tc>' +
        paragraph('+++FOR item IN items+++') +
        '</w:tc></w:tr></w:tbl>'
    );
    const result = await validateTemplate(template);
    expect(result.valid).toBe(false);
    expect(result.diagnostics).toMatchObject([
      {
        code: 'UNEXPECTED_END',
        location: { part: 'word/document.xml', paragraph: 1 },
      },
      {
        code: 'UNCLOSED_BLOCK',
        command: 'FOR item IN items',
        location: {
          part: 'word/document.xml',
          paragraph: 2,
          table: 1,
          row: 1,
          cell: 1,
        },
      },
    ]);
  });

  it('does not match blocks across document parts', async () => {
    const result = await validateTemplate(
      await document(paragraph('+++IF true+++'), paragraph('+++END-IF+++'))
    );
    expect(result.diagnostics.map(d => [d.code, d.location.part])).toEqual([
      ['UNCLOSED_BLOCK', 'word/document.xml'],
      ['UNEXPECTED_END', 'word/header1.xml'],
    ]);
  });

  it('resolves aliases and custom delimiters', async () => {
    const template = await document(
      paragraph('{{ALIAS begin FOR x IN items}}{{*begin}}{{END-FOR x}}')
    );
    expect((await validateTemplate(template, ['{{', '}}'])).valid).toBe(true);
  });

  it('reports malformed commands, unknown aliases, and mismatched nesting', async () => {
    const template = await document(
      paragraph(
        '+++ALIAS broken++++++*missing++++++FOR nope++++++IF ++++++FOR x IN items++++++IF true++++++END-FOR x++++++END-IF++++++END-FOR x+++'
      )
    );
    const result = await validateTemplate(template);
    expect(result.diagnostics.map(d => d.code)).toEqual([
      'INVALID_COMMAND',
      'INVALID_COMMAND',
      'INVALID_COMMAND',
      'INVALID_COMMAND',
      'UNEXPECTED_END',
    ]);
  });

  it('reports unfinished commands', async () => {
    const result = await validateTemplate(
      await document(paragraph('+++INS name'))
    );
    expect(result.diagnostics).toMatchObject([
      { code: 'UNCLOSED_COMMAND', location: { part: 'word/document.xml' } },
    ]);
  });

  it('accepts tags and delimiters split across Word runs', async () => {
    const template = await document(
      '<w:p><w:r><w:t>+</w:t></w:r><w:r><w:t>++IF true++</w:t></w:r><w:r><w:t>+</w:t></w:r></w:p>' +
        paragraph('+++END-IF+++')
    );
    expect((await validateTemplate(template)).valid).toBe(true);
  });

  it.each(['for1inline.docx', 'if2.docx', 'insertInHeaderAndFooter.docx'])(
    'accepts existing Word fixture %s',
    async name => {
      const template = await fs.promises.readFile(
        path.join(__dirname, 'fixtures', name)
      );
      expect(await validateTemplate(template)).toEqual({
        valid: true,
        diagnostics: [],
      });
    }
  );

  it('rejects empty delimiters and corrupt packages', async () => {
    await expect(
      validateTemplate(await document(''), ['', '}'])
    ).rejects.toThrow('must not be empty');
    await expect(validateTemplate(Buffer.from('invalid'))).rejects.toThrow();
  });
});
