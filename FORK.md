# Qitao-Chen/docx-templates

This MIT-licensed fork preserves the upstream history and attribution. Existing
`createReport`, `listCommands`, and `getMetadata` APIs remain available.

## Install

Install the built package from the GitHub release, retaining existing imports:

```sh
npm install docx-templates@https://github.com/Qitao-Chen/docx-templates/releases/download/v4.16.0-qitao.2/qitao-chen-docx-templates-4.16.0-qitao.2.tgz
```

The package name is `@qitao-chen/docx-templates`; the command above aliases it to
`docx-templates`. This release is distributed on GitHub, not the npm registry.
Use a pinned release in applications. The release includes compiled Node code,
TypeScript declarations, and the browser bundle.

## Structural preflight

```ts
import createReport, { validateTemplate } from 'docx-templates';

const result = await validateTemplate(template, ['{{', '}}']);
if (!result.valid) {
  for (const issue of result.diagnostics) {
    console.error(issue.code, issue.message, issue.location);
  }
} else {
  const output = await createReport({
    template,
    cmdDelimiter: ['{{', '}}'],
    data,
  });
}
```

`validateTemplate(template, delimiter?)` checks FOR/IF nesting, matching loop
names, missing closing commands, malformed FOR/IF/ALIAS declarations, unknown
aliases, and unfinished command delimiters. It never evaluates template
JavaScript, invokes data resolvers, or fetches images. Commands split across
Word text runs use the same preprocessing as rendering.

The result is `{ valid, diagnostics }`. Each diagnostic contains `code`,
`message`, `command`, and `location`. Locations include the package `part`
(such as `word/document.xml` or `word/header1.xml`), and, where available,
one-based paragraph/table numbers within that part and row/cell numbers within
the nearest table. Unfinished delimiters currently report the part only.
Block stacks and aliases are scoped to each document part.

Structural validation is not a guarantee of successful rendering. It does
not check JavaScript syntax, runtime errors, layout, schema validity,
or all renderer-specific restrictions. Invalid ZIP/XML packages reject the
promise instead of returning command diagnostics. Rendering still executes
JavaScript according to the upstream behavior and security model.

## Data checks and readable reports (qitao.2)

```ts
import { validateTemplate, formatValidationReport } from 'docx-templates';

const result = await validateTemplate(template, {
  cmdDelimiter: ['{{', '}}'],
  data: { customer: {}, items: [{ price: 10 }] },
});
console.log(formatValidationReport(result));
```

The original `validateTemplate(template, delimiter?)` signature remains supported.
The options form opts into field checks only when it has its own `data` property.
Use plain data objects (for example parsed JSON), not proxies or custom runtime
objects. Property accessors are not invoked. Prototype/inherited properties are
not treated as data fields.

Supported references include `customer.address` and `items[0].price`. The checker
reports `MISSING_FIELD` when a path cannot be resolved. A present terminal value
of `null`, `undefined`, `false`, `0`, or an empty string is not a missing field;
a nullish intermediate value cannot resolve a deeper path. This checks existence,
not types, value suitability, or a JSON Schema. In particular, an intentionally
absent top-level IF field will be reported missing when data checking is enabled.

Complex expressions, constants, getters, conditional/loop body references, and
fields after EXEC or other potentially mutating JavaScript produce
`UNCHECKED_EXPRESSION` warnings instead of speculative missing-field errors.
No template JavaScript is executed. Loop item data is not checked in this version.
Alias expansions and each document part use the same checks. A warning does not
make `valid` false; callers requiring complete coverage should also inspect
warnings. `valid: true` means only that the requested checks found no errors.

Diagnostics now include `severity` and, when available, a bounded `excerpt` of
nearby template paragraph text. Excerpts contain template text, not input data;
template authors should still avoid putting secrets in template text.
`formatValidationReport(result)` renders these diagnostics as plain text for logs
or a CLI. It neither produces HTML nor changes the DOCX. For example:

```text
[error] MISSING_FIELD: Missing own field customer.address
  word/document.xml, paragraph 1
  INS customer.address
  Context: Address: {{INS customer.address}}
```

## Development and maintenance

```sh
yarn install --frozen-lockfile
yarn test
npm run compile
npm pack
```

Keep `upstream` pointing to `https://github.com/guigrpa/docx-templates.git` and
`origin` pointing to this fork. Review upstream changes before merging; run the
full test suite and build before tagging a release. Preserve the original MIT
license and author attribution. Fork versions use the `-qitao.N` suffix.

Future candidates: schema/type checks, annotated DOCX reports, additional Word
editing regression fixtures, native rich text, reusable document fragments, and
reusable compiled templates. These are not included in this release.
