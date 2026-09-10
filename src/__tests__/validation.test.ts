import fs from 'fs';
import path from 'path';
import JSZip from 'jszip';
import { validateTemplate, formatValidationReport, DataSchema } from '../index';

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
    expect(await validateTemplate(template)).toMatchObject({
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
      expect(await validateTemplate(template)).toMatchObject({
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

describe('data preflight and reports', () => {
  it('checks nested paths and array indexes, preserving falsy and null values', async () => {
    const template = await document(
      paragraph(
        '+++customer.name++++++items[0].price++++++enabled++++++empty++++++nil+++'
      )
    );
    expect(
      await validateTemplate(template, {
        data: {
          customer: { name: '' },
          items: [{ price: 0 }],
          enabled: false,
          empty: undefined,
          nil: null,
        },
      })
    ).toMatchObject({
      valid: true,
      diagnostics: [],
      coverage: { checked: 5, skipped: 0 },
    });
  });

  it('finds multiple missing paths with context and report positions', async () => {
    const template = await document(
      paragraph('Customer address: +++customer.address+++') +
        '<w:tbl><w:tr><w:tc>' +
        paragraph('+++items[1].price+++') +
        '</w:tc></w:tr></w:tbl>'
    );
    const result = await validateTemplate(template, {
      data: { customer: {}, items: [{ price: 1 }] },
    });
    expect(result.valid).toBe(false);
    expect(result.diagnostics).toMatchObject([
      {
        code: 'MISSING_FIELD',
        severity: 'error',
        excerpt: 'Customer address: +++customer.address+++',
      },
      { code: 'MISSING_FIELD', location: { table: 1, row: 1, cell: 1 } },
    ]);
    expect(formatValidationReport(result)).toContain('table 1, row 1, cell 1');
    expect(formatValidationReport(result)).toContain(
      'Context: Customer address:'
    );
  });

  it('does not execute getters or complex expressions', async () => {
    const getter = jest.fn(() => 'secret');
    const data = Object.defineProperty({}, 'secret', { get: getter });
    const template = await document(
      paragraph('+++secret++++++fetchSomething()+++')
    );
    const result = await validateTemplate(template, { data });
    expect(result.valid).toBe(true);
    expect(result.diagnostics.map(d => d.code)).toEqual([
      'UNCHECKED_EXPRESSION',
      'UNCHECKED_EXPRESSION',
    ]);
    expect(getter).not.toHaveBeenCalled();
  });

  it('treats conditional and loop fields conservatively', async () => {
    const template = await document(
      paragraph(
        '+++FOR item IN items++++++$item.name++++++END-FOR item++++++IF enabled++++++optional.value++++++END-IF+++'
      )
    );
    const result = await validateTemplate(template, {
      data: { items: [], enabled: false },
    });
    expect(result.valid).toBe(true);
    expect(result.diagnostics).toHaveLength(2);
    expect(result.diagnostics.every(d => d.severity === 'warning')).toBe(true);
  });

  it('warns after EXEC and keeps context independent across parts', async () => {
    const template = await document(
      paragraph('+++EXEC foo = 1++++++foo+++'),
      paragraph('+++missing+++')
    );
    const result = await validateTemplate(template, { data: {} });
    expect(result.diagnostics.map(d => d.code)).toEqual([
      'UNCHECKED_EXPRESSION',
      'UNCHECKED_EXPRESSION',
      'MISSING_FIELD',
    ]);
  });

  it('checks expanded aliases and custom delimiters using the options object', async () => {
    const template = await document(
      paragraph('{{ALIAS address INS customer.address}}{{*address}}')
    );
    const result = await validateTemplate(template, {
      cmdDelimiter: ['{{', '}}'],
      data: { customer: {} },
    });
    expect(result.diagnostics).toMatchObject([
      { code: 'MISSING_FIELD', command: 'INS customer.address' },
    ]);
  });

  it('does not check fields unless data is explicitly supplied', async () => {
    const template = await document(paragraph('+++unknown+++'));
    expect((await validateTemplate(template, {})).diagnostics).toEqual([]);
    expect((await validateTemplate(template, { data: undefined })).valid).toBe(
      false
    );
  });

  it('handles null intermediates and excludes inherited properties', async () => {
    const template = await document(
      paragraph('+++customer.address++++++toString+++')
    );
    const result = await validateTemplate(template, {
      data: { customer: null },
    });
    expect(result.diagnostics.map(d => d.code)).toEqual([
      'MISSING_FIELD',
      'MISSING_FIELD',
    ]);
  });

  it('bounds long excerpts around the command and formats legacy diagnostics', async () => {
    const template = await document(
      paragraph('Before '.repeat(100) + '+++missing+++' + ' After'.repeat(100))
    );
    const result = await validateTemplate(template, { data: {} });
    expect(result.diagnostics[0].excerpt!.length).toBeLessThanOrEqual(242);
    expect(result.diagnostics[0].excerpt).toContain('missing');
    expect(formatValidationReport({ valid: true, diagnostics: [] })).toContain(
      'No issues found'
    );
    expect(
      formatValidationReport({
        valid: false,
        diagnostics: [
          {
            code: 'UNCLOSED_COMMAND',
            message: 'Unclosed',
            command: 'INS x',
            location: { part: 'word/document.xml' },
          },
        ],
      })
    ).toContain('[error]');
  });
});

it('does not report missing fields after expressions that may change data', async () => {
  const template = await document(
    paragraph('+++initialize()++++++created.name+++')
  );
  const result = await validateTemplate(template, { data: {} });
  expect(result.valid).toBe(true);
  expect(result.diagnostics.map(d => d.code)).toEqual([
    'UNCHECKED_EXPRESSION',
    'UNCHECKED_EXPRESSION',
  ]);
});

describe('loop inspection and coverage', () => {
  it('reports the actual missing item path and counts repeated checks', async () => {
    const template = await document(
      paragraph('+++FOR item IN items++++++$item.price++++++END-FOR item+++')
    );
    const result = await validateTemplate(template, {
      data: { items: [{ price: 0 }, { price: 2 }, {}] },
    });
    expect(result.coverage).toEqual({ checked: 4, skipped: 0 });
    expect(result.diagnostics).toMatchObject([
      {
        code: 'MISSING_FIELD',
        dataPath: 'items[2].price',
        iterations: [{ variable: 'item', index: 2 }],
      },
    ]);
    expect(formatValidationReport(result)).toContain('item #3');
    expect(formatValidationReport(result)).toContain('4 checked, 0 skipped');
  });

  it('resolves nested loops and restores outer bindings', async () => {
    const template = await document(
      paragraph(
        '+++FOR group IN groups++++++FOR item IN $group.items++++++$item.name++++++END-FOR item++++++$group.title++++++END-FOR group++++++footer+++'
      )
    );
    const result = await validateTemplate(template, {
      data: {
        groups: [
          { title: 'A', items: [{ name: 'ok' }, {}] },
          { title: 'B', items: [{ name: 'ok' }] },
        ],
        footer: '',
      },
    });
    expect(result.coverage).toEqual({ checked: 9, skipped: 0 });
    expect(result.diagnostics).toMatchObject([
      {
        dataPath: 'groups[0].items[1].name',
        iterations: [
          { variable: 'group', index: 0 },
          { variable: 'item', index: 1 },
        ],
      },
    ]);
  });

  it('restores a shadowed loop binding and checks aliases', async () => {
    const template = await document(
      paragraph(
        '+++ALIAS label INS $item.name++++++FOR item IN items++++++FOR item IN $item.children++++++*label++++++END-FOR item++++++*label++++++END-FOR item+++'
      )
    );
    const result = await validateTemplate(template, {
      data: { items: [{ name: 'parent', children: [{}] }] },
    });
    expect(result.diagnostics).toHaveLength(1);
    expect(result.diagnostics[0].dataPath).toBe('items[0].children[0].name');
  });

  it('distinguishes invalid, missing, empty and dynamic sources', async () => {
    const template = await document(
      paragraph('+++FOR item IN items++++++$item.name++++++END-FOR item+++')
    );
    expect(
      (await validateTemplate(template, { data: { items: 3 } })).diagnostics[0]
        .code
    ).toBe('INVALID_LOOP_DATA');
    expect(
      (await validateTemplate(template, { data: {} })).diagnostics[0].code
    ).toBe('MISSING_FIELD');
    const empty = await validateTemplate(template, { data: { items: [] } });
    expect(empty.coverage).toEqual({ checked: 1, skipped: 1 });
    expect(empty.diagnostics[0].message).toContain('Empty loop');
    const dynamic = await validateTemplate(
      await document(
        paragraph(
          '+++FOR item IN getItems()++++++$item.name++++++END-FOR item+++'
        )
      ),
      { data: {} }
    );
    expect(dynamic.coverage).toEqual({ checked: 0, skipped: 2 });
  });

  it('does not invoke accessor array elements', async () => {
    const getter = jest.fn(() => ({ name: 'secret' }));
    const items: unknown[] = [];
    Object.defineProperty(items, '0', { get: getter });
    const template = await document(
      paragraph('+++FOR item IN items++++++$item.name++++++END-FOR item+++')
    );
    const result = await validateTemplate(template, { data: { items } });
    expect(result.coverage).toEqual({ checked: 1, skipped: 1 });
    expect(getter).not.toHaveBeenCalled();
  });

  it('limits expansion and reports skipped checks', async () => {
    const template = await document(
      paragraph('+++FOR item IN items++++++$item.name++++++END-FOR item+++')
    );
    const result = await validateTemplate(template, {
      data: { items: [{ name: 'A' }, { name: 'B' }, {}] },
      maxLoopItems: 1,
    });
    expect(result.coverage).toEqual({ checked: 2, skipped: 1 });
    expect(result.diagnostics[0].message).toContain('limit');
    await expect(
      validateTemplate(template, { maxLoopItems: 0 })
    ).rejects.toThrow('positive safe integer');
  });

  it('aggregates part coverage without leaking loop scope', async () => {
    const template = await document(
      paragraph('+++FOR item IN items++++++$item.name++++++END-FOR item+++'),
      paragraph('+++title+++')
    );
    const result = await validateTemplate(template, {
      data: { items: [{ name: '' }], title: '' },
    });
    expect(result.coverage).toEqual({ checked: 3, skipped: 0 });
    expect(result.diagnostics).toEqual([]);
    expect((await validateTemplate(template)).coverage).toBeUndefined();
  });
});

describe('schema validation and check status', () => {
  const schema: DataSchema = {
    type: 'object',
    required: true,
    properties: {
      customer: {
        type: 'object',
        required: true,
        properties: {
          name: { type: 'string', required: true, minLength: 1, maxLength: 8 },
        },
      },
      items: {
        type: 'array',
        required: true,
        minItems: 1,
        maxItems: 4,
        items: {
          type: 'object',
          properties: {
            price: { type: 'number', required: true, minimum: 0, maximum: 100 },
          },
        },
      },
      address: { type: 'string', nullable: true },
    },
  };

  it('accepts valid data and separates schema checks from field checks', async () => {
    const template = await document(paragraph('+++customer.name+++'));
    const result = await validateTemplate(template, {
      data: {
        customer: { name: 'Alice' },
        items: [{ price: 0 }],
        address: null,
      },
      schema,
    });
    expect(result.valid).toBe(true);
    expect(result.coverage).toEqual({ checked: 1, skipped: 0 });
    expect(result.schemaCoverage).toEqual({ checked: 7, skipped: 0 });
    expect(result.checks).toEqual({
      structure: 'valid',
      data: 'valid',
      schema: 'valid',
      coverage: 'complete',
    });
    expect(formatValidationReport(result)).toContain(
      'Schema checks: 7 checked'
    );
  });

  it('associates nested item errors with template positions and iterations', async () => {
    const template = await document(
      '<w:tbl><w:tr><w:tc>' +
        paragraph(
          '+++FOR item IN items++++++$item.price++++++END-FOR item+++'
        ) +
        '</w:tc></w:tr></w:tbl>'
    );
    const result = await validateTemplate(template, {
      data: {
        customer: { name: '' },
        items: [{ price: 2 }, { price: 'SECRET_VALUE' }],
      },
      schema,
    });
    expect(result.diagnostics).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          source: 'schema',
          dataPath: 'items[1].price',
          iterations: [{ variable: 'item', index: 1 }],
          location: expect.objectContaining({ table: 1, row: 1, cell: 1 }),
        }),
        expect.objectContaining({
          source: 'schema',
          dataPath: 'customer.name',
          location: { part: 'data' },
        }),
      ])
    );
    expect(formatValidationReport(result)).not.toContain('SECRET_VALUE');
    expect(result.checks).toMatchObject({
      structure: 'valid',
      schema: 'invalid',
      data: 'invalid',
    });
  });

  it('enforces required and nullability independently', async () => {
    const template = await document('');
    const result = await validateTemplate(template, {
      data: { present: null, empty: '', optional: undefined },
      schema: {
        type: 'object',
        properties: {
          missing: { type: 'string', required: true },
          present: { type: 'string' },
          empty: { type: 'string', required: true },
          optional: { type: 'string' },
          absentNullable: { type: 'string', nullable: true, required: true },
        },
      },
    });
    expect(result.diagnostics.map(issue => issue.dataPath)).toEqual([
      'missing',
      'present',
      'absentNullable',
    ]);
  });

  it.each([
    [{ type: 'number', minimum: 0 }, -1],
    [{ type: 'number', maximum: 2 }, 3],
    [{ type: 'number' }, Infinity],
    [{ type: 'number' }, NaN],
    [{ type: 'integer' }, 1.2],
    [{ type: 'boolean' }, 'true'],
    [{ type: 'string', maxLength: 1 }, 'AB'],
    [{ type: 'array', minItems: 1 }, []],
    [{ type: 'array', maxItems: 1 }, [1, 2]],
    [{ type: 'object' }, []],
    [{ type: 'array' }, {}],
  ])('rejects type/range violation %j with %j', async (rule, data) => {
    const result = await validateTemplate(await document(''), {
      data,
      schema: rule as DataSchema,
    });
    expect(result.checks!.schema).toBe('invalid');
    expect(result.diagnostics[0].code).toBe('SCHEMA_VIOLATION');
  });

  it('counts string code points and checks integer, boolean and optional array items', async () => {
    const result = await validateTemplate(await document(''), {
      data: { name: '😀', n: 2, enabled: false, list: [null] },
      schema: {
        type: 'object',
        properties: {
          name: { type: 'string', minLength: 1, maxLength: 1 },
          n: { type: 'integer' },
          enabled: { type: 'boolean' },
          list: { type: 'array' },
        },
      },
    });
    expect(result.valid).toBe(true);
  });

  it('does not invoke property or item getters and exposes incomplete status', async () => {
    const getter = jest.fn(() => 'secret');
    const data = Object.defineProperty({}, 'name', { get: getter });
    const result = await validateTemplate(await document(''), {
      data,
      schema: { type: 'object', properties: { name: { type: 'string' } } },
    });
    expect(getter).not.toHaveBeenCalled();
    expect(result.valid).toBe(true);
    expect(result.checks).toMatchObject({
      data: 'incomplete',
      schema: 'incomplete',
      coverage: 'partial',
    });
    const list: unknown[] = [];
    Object.defineProperty(list, '0', { get: getter });
    const arrayResult = await validateTemplate(await document(''), {
      data: list,
      schema: { type: 'array', items: { type: 'string' } },
    });
    expect(arrayResult.diagnostics[0].code).toBe('UNCHECKED_SCHEMA');
    expect(getter).not.toHaveBeenCalled();
  });

  it('limits schema traversal and preserves known failures', async () => {
    const result = await validateTemplate(await document(''), {
      data: ['wrong', 1, 2],
      schema: { type: 'array', items: { type: 'number' } },
      maxSchemaChecks: 2,
    });
    expect(result.schemaCoverage).toEqual({ checked: 2, skipped: 1 });
    expect(result.checks).toMatchObject({
      schema: 'invalid',
      coverage: 'partial',
    });
  });

  it('keeps structure, missing fields and skipped expressions distinct', async () => {
    const template = await document(
      paragraph('+++END-IF++++++missing++++++run()+++')
    );
    const result = await validateTemplate(template, {
      data: {},
      schema: { type: 'object' },
    });
    expect(result.checks).toEqual({
      structure: 'invalid',
      data: 'invalid',
      schema: 'valid',
      coverage: 'partial',
    });
    expect((await validateTemplate(await document(''))).checks).toEqual({
      structure: 'valid',
      data: 'not-requested',
      schema: 'not-requested',
      coverage: 'not-requested',
    });
  });

  it('checks sparse arrays and inherited properties as absent', async () => {
    const data = Object.create({ name: 'inherited' });
    data.items = new Array(1);
    const result = await validateTemplate(await document(''), {
      data,
      schema: {
        type: 'object',
        properties: {
          name: { type: 'string', required: true },
          items: { type: 'array', items: { type: 'number', required: true } },
        },
      },
    });
    expect(result.diagnostics.map(issue => issue.dataPath)).toEqual([
      'name',
      'items[0]',
    ]);
  });

  it.each([
    { type: 'date' },
    { type: 'string', pattern: 'x' },
    { type: 'number', minimum: 2, maximum: 1 },
    { type: 'string', minLength: -1 },
    { type: 'array', minItems: 0.5 },
    { type: 'boolean', required: [] },
    { type: 'string', nullable: 1 },
    { type: 'number', minimum: Infinity },
    { type: 'array', items: undefined },
  ])('rejects invalid schema configuration %j', async rule => {
    await expect(
      validateTemplate(await document(''), {
        data: {},
        schema: rule as DataSchema,
      })
    ).rejects.toThrow();
  });

  it('rejects cyclic/accessor schemas, schema without data, and invalid budgets', async () => {
    const cyclic: DataSchema = { type: 'array' };
    cyclic.items = cyclic;
    const getter = jest.fn(() => 'string');
    const accessor = Object.defineProperty({}, 'type', { get: getter });
    const template = await document('');
    await expect(
      validateTemplate(template, { data: [], schema: cyclic })
    ).rejects.toThrow('cyclic');
    await expect(
      validateTemplate(template, { data: {}, schema: accessor as DataSchema })
    ).rejects.toThrow('accessors');
    expect(getter).not.toHaveBeenCalled();
    await expect(validateTemplate(template, { schema })).rejects.toThrow(
      'explicit data'
    );
    await expect(
      validateTemplate(template, { data: {}, maxSchemaChecks: 0 })
    ).rejects.toThrow('positive safe integer');
  });
});

it('maps schema violations to repeated/header references and nearest source locations', async () => {
  const template = await document(
    paragraph('+++FOR item IN items++++++END-FOR item+++'),
    paragraph('+++name+++')
  );
  const result = await validateTemplate(template, {
    data: { items: [{ price: -1 }], name: 42 },
    schema: {
      type: 'object',
      properties: {
        items: {
          type: 'array',
          items: {
            type: 'object',
            properties: { price: { type: 'number', minimum: 0 } },
          },
        },
        name: { type: 'string' },
      },
    },
  });
  expect(result.diagnostics).toMatchObject([
    {
      source: 'schema',
      dataPath: 'items[0].price',
      command: 'FOR item IN items',
      location: { part: 'word/document.xml' },
    },
    {
      source: 'schema',
      dataPath: 'name',
      location: { part: 'word/header1.xml' },
    },
  ]);
});

it('allows reused schema definitions and quotes unusual data keys', async () => {
  const child: DataSchema = { type: 'string', required: true };
  const result = await validateTemplate(await document(''), {
    data: {},
    schema: { type: 'object', properties: { 'a.b': child, other: child } },
  });
  expect(result.diagnostics.map(issue => issue.dataPath)).toEqual([
    '["a.b"]',
    'other',
  ]);
});
