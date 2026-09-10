# Qitao-Chen/docx-templates

This MIT-licensed fork preserves the upstream history and attribution. Existing
`createReport`, `listCommands`, and `getMetadata` APIs remain available.

## Install

Install the built package from the GitHub release, retaining existing imports:

```sh
npm install docx-templates@https://github.com/Qitao-Chen/docx-templates/releases/download/v4.16.0-qitao.4/qitao-chen-docx-templates-4.16.0-qitao.4.tgz
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

Complex expressions, constants, getters, conditional body references, and
fields after EXEC or other potentially mutating JavaScript produce
`UNCHECKED_EXPRESSION` warnings instead of speculative missing-field errors.
No template JavaScript is executed. Resolvable FOR sources are checked item by item as described below.
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

## Loop checks and coverage (qitao.3)

Simple FOR sources now support per-item checks, including nested loops and aliases.
For example, `FOR item IN items` followed by `INS $item.price` can report
`dataPath: "items[2].price"` and `iterations: [{ variable: "item", index: 2 }]`.
Indexes in structured diagnostics are zero-based; the text report displays
one-based item numbers. Document locations still refer to the template, not the
expanded output rows. Loop bindings are restored after END-FOR.

With `data` supplied, the result adds `coverage: { checked, skipped }`. These are
field-check attempts, not unique tags: three loop items produce three checks for
each field. Missing fields count as checked because their absence was determined.
FOR source expressions and EXEC skips are included. A skipped body with no
inspectable items contributes one placeholder attempt per expression; skipped
counts are not estimates of an unknown number of runtime iterations. Without
`data`, coverage is omitted for compatibility. `formatValidationReport` includes
the counts even when there are no errors.

Non-array resolved sources produce `INVALID_LOOP_DATA`. Empty, missing, dynamic,
sparse or accessor-based sources/items leave body references unchecked with an
explicit warning. Conditional bodies remain unchecked. Checks use the supplied
data snapshot and never execute JavaScript; they do not simulate mutations across
iterations or evaluate conditions. Supply plain data, inspect warnings, and do not
treat `valid: true` as a guarantee of successful rendering.

`maxLoopItems` limits total expanded items per document part (default 10000,
positive safe integer). Remaining body checks become warnings on reaching the
limit, rather than silently claiming complete coverage. For example:

```ts
const result = await validateTemplate(template, { data, maxLoopItems: 5000 });
console.log(result.coverage); // { checked: ..., skipped: ... }
console.log(formatValidationReport(result));
```

The GitHub test workflow runs tests, builds through the install lifecycle, packs
and installs the compiled package, then checks its public validation/report and
DOCX rendering APIs. It supports push, pull request and manual dispatch.

## Optional schema and separate statuses (qitao.4)

This is a **library-specific lightweight schema**, not JSON Schema. No external
validator dependency is required. Unknown keywords and malformed rules reject the
validation promise with a TypeError rather than being silently ignored.

```ts
import { validateTemplate, formatValidationReport, DataSchema } from 'docx-templates';

const schema: DataSchema = {
  type: 'object', required: true,
  properties: {
    customer: {
      type: 'object', required: true,
      properties: { name: { type: 'string', required: true, minLength: 1 } },
    },
    items: {
      type: 'array', required: true, minItems: 1,
      items: {
        type: 'object',
        properties: { price: { type: 'number', required: true, minimum: 0 } },
      },
    },
    address: { type: 'string', nullable: true },
  },
};
const result = await validateTemplate(template, { data, schema });
console.log(result.checks);
console.log(formatValidationReport(result));
```

Rules:

- All nodes: `type`, optional boolean `required` and `nullable` (both default false).
- `string`: optional `minLength`/`maxLength`, measured in Unicode code points.
- `number`/`integer`: optional inclusive `minimum`/`maximum`; nonfinite numbers fail.
- `boolean`: boolean values only, with no coercion.
- `object`: optional `properties`, a map of property names to schemas. Extra keys are allowed.
- `array`: optional `items`, `minItems` and `maxItems`. Without `items`, elements have no schema constraint.

Required means present and not undefined. It does not mean nonempty: use
`minLength: 1` or `minItems: 1` for that. Strings are not trimmed. A required
nullable field may contain null but may not be missing. An optional object that
is absent does not trigger its children's required rules. Sparse array elements
are treated as undefined. Inherited data properties are absent. There is no
coercion, defaulting, mutation, `$ref`, regex, union or format support.

Schemas check all declared data, even fields not referenced by the template.
`SCHEMA_VIOLATION` diagnostics carry `source: 'schema'` and the actual `dataPath`.
Exact template references supply document locations and loop indexes. If only an
ancestor path is referenced, its location is used (for example the FOR source);
otherwise `location.part` is `data`, not a claimed DOCX location. Repeated
references may produce multiple located diagnostics for the same violation.
Input values are not copied into reports.

`schema` requires an explicitly supplied `data` option. Plain data and schema
objects are expected; do not pass proxies. Schema accessors/cycles are rejected;
data accessors produce `UNCHECKED_SCHEMA` warnings without invoking getters.
`maxSchemaChecks` is a positive safe integer, default 10000, bounding schema-node
visits across the whole data object. Reaching the budget adds a warning for the
unvisited remainder. Schema definitions are limited to 10000 nodes and depth 100.
`schemaCoverage` counts schema-node visits separately from field checks; one
skipped count can represent an unvisited subtree/remainder, not its unknown size.

Every validation result now has an additive `checks` object:

- `structure`: `valid` or `invalid` for supported structural checks.
- `data`: `not-requested`, `valid`, `invalid`, or `incomplete`; combines field and schema checks.
- `schema`: `not-requested`, `valid`, `invalid`, or `incomplete` for schema checks alone.
- `coverage`: `not-requested`, `complete`, or `partial` for the requested data checks.

`invalid` takes precedence over `incomplete`; `coverage: partial` still reveals
skips alongside known errors. `complete` refers only to requested checks, not
full JavaScript execution, renderer restrictions or document layout. The existing
`valid` flag still means no reported errors; warnings alone do not make it false.
Existing calls remain supported and the formatter still accepts older result
objects without `checks`. Consumers comparing whole result objects should allow
these additive fields.

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

Future candidates: annotated DOCX reports, additional Word
editing regression fixtures, native rich text, reusable document fragments, and
reusable compiled templates. These are not included in this release.
