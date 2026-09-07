# Qitao-Chen/docx-templates

This MIT-licensed fork preserves the upstream history and attribution. Existing
`createReport`, `listCommands`, and `getMetadata` APIs remain available.

## Install

Install the built package from the GitHub release, retaining existing imports:

```sh
npm install docx-templates@https://github.com/Qitao-Chen/docx-templates/releases/download/v4.16.0-qitao.1/qitao-chen-docx-templates-4.16.0-qitao.1.tgz
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

This is structural validation, not a guarantee of successful rendering. It does
not check JavaScript syntax, missing data, runtime errors, layout, schema validity,
or all renderer-specific restrictions. Invalid ZIP/XML packages reject the
promise instead of returning command diagnostics. Rendering still executes
JavaScript according to the upstream behavior and security model.

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

Future candidates: richer diagnostics and field/schema checks, additional Word
editing regression fixtures, native rich text, reusable document fragments, and
reusable compiled templates. These are not included in this release.
