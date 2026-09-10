import type { DataSchema } from './schema';
import type { Context, Node } from './types';
import { getCommand, splitCommand } from './processTemplate';

export interface ValidationOptions {
  cmdDelimiter?: string | [string, string];
  /** Opt in to checking simple property paths against plain data. */
  data?: unknown;
  /** Maximum expanded loop items per document part (default 10000). */
  maxLoopItems?: number;
  schema?: DataSchema;
  /** Schema-node check budget, default 10000. */
  maxSchemaChecks?: number;
}

export interface TemplateLocation {
  /** DOCX package part, e.g. word/header1.xml. */
  part: string;
  /** One-based paragraph and table numbers within this part. */
  paragraph?: number;
  table?: number;
  /** One-based row and cell within the nearest table. */
  row?: number;
  cell?: number;
}

export interface TemplateDiagnostic {
  code:
    | 'SCHEMA_VIOLATION'
    | 'UNCHECKED_SCHEMA'
    | 'INVALID_COMMAND'
    | 'UNEXPECTED_END'
    | 'UNCLOSED_BLOCK'
    | 'UNCLOSED_COMMAND'
    | 'INVALID_LOOP_DATA'
    | 'MISSING_FIELD'
    | 'UNCHECKED_EXPRESSION';
  message: string;
  command: string;
  location: TemplateLocation;
  /** Absent on older diagnostics means error. */
  severity?: 'error' | 'warning';
  excerpt?: string;
  source?: 'schema';
  dataPath?: string;
  iterations?: { variable: string; index: number }[];
}

export interface ValidationCoverage {
  /** Field-check attempts, including repeated loop instances. */
  checked: number;
  skipped: number;
}

type Scope = {
  bindings: Record<string, { value: unknown; path: string }>;
  iterations: { variable: string; index: number }[];
  unavailable?: string;
};

export interface ValidationChecks {
  structure: 'valid' | 'invalid';
  data: 'not-requested' | 'valid' | 'invalid' | 'incomplete';
  schema: 'not-requested' | 'valid' | 'invalid' | 'incomplete';
  coverage: 'not-requested' | 'complete' | 'partial';
}

export interface ValidationResult {
  /** True when requested checks find no errors; inspect checks for incompleteness. */
  valid: boolean;
  diagnostics: TemplateDiagnostic[];
  /** Present when data checks were requested. Skipped checks may hide errors. */
  coverage?: ValidationCoverage;
  schemaCoverage?: ValidationCoverage;
  checks?: ValidationChecks;
}

export class TemplateValidator {
  readonly references: TemplateDiagnostic[] = [];
  readonly coverage: ValidationCoverage = { checked: 0, skipped: 0 };
  private scopes: Scope[] = [{ bindings: {}, iterations: [] }];
  private expandedItems = 0;
  private dynamicContext = false;
  private locations = new Map<Node, TemplateLocation>();
  private blocks: {
    type: string;
    name: string;
    raw: string;
    node: Node;
    scopes: Scope[];
  }[] = [];

  constructor(
    tree: Node,
    private part: string,
    private diagnostics: TemplateDiagnostic[],
    private options: ValidationOptions = {}
  ) {
    let paragraph = 0;
    let table = 0;
    const pending: { node: Node; location: TemplateLocation }[] = [
      { node: tree, location: { part } },
    ];
    while (pending.length) {
      const { node, location: inherited } = pending.pop()!;
      const location = { ...inherited };
      if (!node._fTextNode) {
        if (node._tag === 'w:p') location.paragraph = ++paragraph;
        if (node._tag === 'w:tbl') {
          location.table = ++table;
          delete location.row;
          delete location.cell;
        }
        if (node._tag === 'w:tr' || node._tag === 'w:tc') {
          const siblings = node._parent?._children.filter(
            sibling => !sibling._fTextNode && sibling._tag === node._tag
          );
          const position = siblings ? siblings.indexOf(node) + 1 : undefined;
          if (node._tag === 'w:tr') location.row = position;
          else location.cell = position;
        }
      }
      this.locations.set(node, location);
      for (let i = node._children.length - 1; i >= 0; i--) {
        pending.push({ node: node._children[i], location });
      }
    }
  }

  private add(
    code: TemplateDiagnostic['code'],
    message: string,
    command: string,
    node?: Node,
    severity: 'error' | 'warning' = 'error',
    scope?: Scope,
    dataPath?: string
  ) {
    this.diagnostics.push({
      code,
      severity,
      ...(scope ? { iterations: scope.iterations, dataPath } : {}),
      excerpt: node ? this.excerpt(node, command) : undefined,
      message,
      command,
      location: (node && this.locations.get(node)) || { part: this.part },
    });
  }

  private excerpt(node: Node, command: string): string | undefined {
    let paragraph: Node | undefined = node;
    while (paragraph && (paragraph._fTextNode || paragraph._tag !== 'w:p')) {
      paragraph = paragraph._parent || undefined;
    }
    if (!paragraph) return undefined;
    const pending: Node[] = [paragraph];
    let text = '';
    while (pending.length) {
      const current = pending.pop()!;
      if (
        current._fTextNode &&
        current._parent &&
        !current._parent._fTextNode &&
        current._parent._tag === 'w:t'
      ) {
        text += current._text;
      }
      for (let i = current._children.length - 1; i >= 0; i--)
        pending.push(current._children[i]);
    }
    text = text.replace(/\s+/g, ' ').trim();
    const anchor = node._fTextNode ? node._text : command;
    const index = text.indexOf(anchor.replace(/\s+/g, ' ').trim());
    const start = Math.max(0, index - 80);
    return (
      (start ? '…' : '') +
      text.slice(start, start + 240) +
      (text.length > start + 240 ? '…' : '')
    );
  }

  private checkField(
    expression: string,
    raw: string,
    node: Node,
    scope: Scope
  ): { value: unknown; path: string } | undefined {
    const unchecked = (message: string) => {
      this.coverage.skipped++;
      this.add('UNCHECKED_EXPRESSION', message, raw, node, 'warning', scope);
    };
    const isPath =
      /^[A-Za-z_$][\w$]*(?:\.[A-Za-z_$][\w$]*|\[(?:0|[1-9]\d*)\])*$/.test(
        expression
      );
    if (!isPath) this.dynamicContext = true;
    if (
      this.dynamicContext ||
      scope.unavailable ||
      this.blocks.some(block => block.type === 'IF')
    ) {
      unchecked(
        scope.unavailable ||
          'Expression depends on dynamic JavaScript or a conditional; not checked'
      );
      return;
    }
    if (/^(true|false|null|undefined|NaN|Infinity)$/.test(expression)) {
      unchecked('Only simple property paths can be checked statically');
      return;
    }
    const segments = expression.replace(/\[(\d+)\]/g, '.$1').split('.');
    const binding = Object.prototype.hasOwnProperty.call(
      scope.bindings,
      segments[0]
    )
      ? scope.bindings[segments[0]]
      : undefined;
    let path = expression;
    let value: unknown = this.options.data;
    if (binding) {
      const root = segments.shift()!;
      path = binding.path + expression.slice(root.length);
      value = binding.value;
    }
    if (this.options.schema)
      this.references.push({
        code: 'SCHEMA_VIOLATION',
        message: '',
        command: raw,
        location: this.locations.get(node) || { part: this.part },
        dataPath: path,
        iterations: scope.iterations,
        excerpt: this.excerpt(node, raw),
      });
    for (const segment of segments) {
      const descriptor =
        value != null &&
        (typeof value === 'object' || typeof value === 'string')
          ? Object.getOwnPropertyDescriptor(Object(value), segment)
          : undefined;
      if (!descriptor) {
        this.coverage.checked++;
        this.add(
          'MISSING_FIELD',
          `Missing own field ${path}`,
          raw,
          node,
          'error',
          scope,
          path
        );
        return;
      }
      if (!('value' in descriptor)) {
        this.dynamicContext = true;
        unchecked(`Field ${path} uses an accessor; not checked`);
        return;
      }
      value = descriptor.value;
    }
    this.coverage.checked++;
    return { value, path };
  }

  command(source: string, node: Node, ctx: Context) {
    let raw: string;
    try {
      raw = getCommand(source, ctx.shorthands, false);
    } catch (error) {
      this.add('INVALID_COMMAND', String(error), source.trim(), node);
      return;
    }
    const { cmdName: type, cmdRest: code } = splitCommand(raw);
    let loopScopes: Scope[] | undefined;
    if (Object.prototype.hasOwnProperty.call(this.options, 'data')) {
      if (type === 'EXEC') {
        this.dynamicContext = true;
        this.coverage.skipped += this.scopes.length;
        this.add(
          'UNCHECKED_EXPRESSION',
          'EXEC is not executed; subsequent fields cannot be checked reliably',
          raw,
          node,
          'warning'
        );
      } else if (
        ['INS', 'IMAGE', 'LINK', 'HTML', 'IF', 'FOR'].includes(type || '')
      ) {
        const match = type === 'FOR' ? /^(\S+)\s+IN\s+(.+)/i.exec(code) : null;
        if (type === 'FOR' && match) loopScopes = [];
        if (type !== 'FOR' || match) {
          for (const scope of this.scopes) {
            const resolved = this.checkField(
              match ? match[2] : code,
              raw,
              node,
              scope
            );
            if (loopScopes && match) {
              if (!resolved || !Array.isArray(resolved.value)) {
                if (resolved)
                  this.add(
                    'INVALID_LOOP_DATA',
                    `FOR source ${resolved.path} must be an array`,
                    raw,
                    node,
                    'error',
                    scope,
                    resolved.path
                  );
                loopScopes.push({
                  ...scope,
                  unavailable: 'Loop source could not be inspected',
                });
              } else if (!resolved.value.length) {
                loopScopes.push({
                  ...scope,
                  unavailable: 'Empty loop has no items to inspect',
                });
              } else {
                for (let index = 0; index < resolved.value.length; index++) {
                  if (
                    this.expandedItems >= (this.options.maxLoopItems ?? 10000)
                  ) {
                    loopScopes.push({
                      ...scope,
                      unavailable:
                        'Loop item inspection limit reached; remaining items not checked',
                    });
                    break;
                  }
                  this.expandedItems++;
                  const item = Object.getOwnPropertyDescriptor(
                    resolved.value,
                    String(index)
                  );
                  const path = `${resolved.path}[${index}]`;
                  loopScopes.push({
                    bindings: {
                      ...scope.bindings,
                      [`$${match[1]}`]: {
                        value: item && 'value' in item ? item.value : undefined,
                        path,
                      },
                    },
                    iterations: [
                      ...scope.iterations,
                      { variable: match[1], index },
                    ],
                    ...(!item || !('value' in item)
                      ? {
                          unavailable:
                            'Sparse or accessor loop item not checked',
                        }
                      : {}),
                  });
                }
              }
            }
          }
        }
      }
    }
    if (type === 'ALIAS') {
      const match = /^(\S+)\s+(.+)/.exec(code);
      if (match) ctx.shorthands[match[1]] = match[2];
      else this.add('INVALID_COMMAND', 'Invalid ALIAS command', raw, node);
    } else if (type === 'FOR' || type === 'IF') {
      const match = /^(\S+)\s+IN\s+(.+)/i.exec(code);
      if ((type === 'FOR' && !match) || (type === 'IF' && !code)) {
        this.add('INVALID_COMMAND', `Invalid ${type} command`, raw, node);
        return;
      }
      this.blocks.push({
        type,
        name: type === 'FOR' ? match![1] : '',
        raw,
        node,
        scopes: this.scopes,
      });
      if (loopScopes) this.scopes = loopScopes;
    } else if (type === 'END-FOR' || type === 'END-IF') {
      const block = this.blocks[this.blocks.length - 1];
      if (
        !block ||
        type !== `END-${block.type}` ||
        (type === 'END-FOR' && code !== block.name)
      ) {
        this.add(
          'UNEXPECTED_END',
          `Unexpected ${raw}; closing commands must match the innermost block`,
          raw,
          node
        );
      } else {
        this.scopes = block.scopes;
        this.blocks.pop();
      }
    }
  }

  finish(ctx: Context) {
    if (ctx.fCmd) {
      this.add(
        'UNCLOSED_COMMAND',
        'Command is missing its closing delimiter',
        ctx.cmd
      );
    }
    for (const block of this.blocks) {
      this.add(
        'UNCLOSED_BLOCK',
        `Missing END-${block.type}${block.name ? ` ${block.name}` : ''}`,
        block.raw,
        block.node
      );
    }
  }
}

/** Plain text suitable for CLI output and logs; includes no data values. */
export function formatValidationReport(result: ValidationResult): string {
  const statuses = result.checks
    ? `Structure: ${result.checks.structure}; data: ${result.checks.data}; schema: ${result.checks.schema}; coverage: ${result.checks.coverage}.\n`
    : '';
  const summary =
    statuses +
    (result.coverage
      ? `Field checks: ${result.coverage.checked} checked, ${result.coverage.skipped} skipped.\n\n`
      : '');
  const schemaSummary = result.schemaCoverage
    ? `Schema checks: ${result.schemaCoverage.checked} checked, ${result.schemaCoverage.skipped} skipped.\n\n`
    : '';
  if (!result.diagnostics.length)
    return summary + schemaSummary + 'No issues found by the requested checks.';
  return (
    summary +
    schemaSummary +
    result.diagnostics
      .map(issue => {
        const { part, paragraph, table, row, cell } = issue.location;
        const position = [
          part,
          paragraph && `paragraph ${paragraph}`,
          table && `table ${table}`,
          row && `row ${row}`,
          cell && `cell ${cell}`,
        ]
          .filter(Boolean)
          .join(', ');
        return (
          `[${issue.severity || 'error'}] ${issue.code}: ${
            issue.message
          }\n  ${position}\n  ${issue.command}` +
          (issue.iterations?.length
            ? `\n  Iterations: ${issue.iterations
                .map(item => `${item.variable} #${item.index + 1}`)
                .join(', ')}`
            : '') +
          (issue.excerpt ? `\n  Context: ${issue.excerpt}` : '')
        );
      })
      .join('\n\n')
  );
}
