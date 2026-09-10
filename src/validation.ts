import type { Context, Node } from './types';
import { getCommand, splitCommand } from './processTemplate';

export interface ValidationOptions {
  cmdDelimiter?: string | [string, string];
  /** Opt in to checking simple property paths against plain data. */
  data?: unknown;
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
    | 'INVALID_COMMAND'
    | 'UNEXPECTED_END'
    | 'UNCLOSED_BLOCK'
    | 'UNCLOSED_COMMAND'
    | 'MISSING_FIELD'
    | 'UNCHECKED_EXPRESSION';
  message: string;
  command: string;
  location: TemplateLocation;
  /** Absent on older diagnostics means error. */
  severity?: 'error' | 'warning';
  excerpt?: string;
}

export interface ValidationResult {
  /** Structural validity only; expressions are never evaluated. */
  valid: boolean;
  diagnostics: TemplateDiagnostic[];
}

export class TemplateValidator {
  private dynamicContext = false;
  private locations = new Map<Node, TemplateLocation>();
  private blocks: { type: string; name: string; raw: string; node: Node }[] =
    [];

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
    severity: 'error' | 'warning' = 'error'
  ) {
    this.diagnostics.push({
      code,
      severity,
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

  private checkField(expression: string, raw: string, node: Node) {
    const unchecked = (message: string) =>
      this.add('UNCHECKED_EXPRESSION', message, raw, node, 'warning');
    const isPath = /^[A-Za-z_$][\w$]*(?:\.[A-Za-z_$][\w$]*|\[\d+\])*$/.test(
      expression
    );
    if (!isPath) this.dynamicContext = true;
    if (this.dynamicContext || this.blocks.length) {
      unchecked(
        'Expression depends on dynamic JavaScript, a conditional, or a loop; not checked'
      );
      return;
    }
    // Deliberately accept only property paths. Never evaluate JS or invoke getters.
    if (
      !/^[A-Za-z_$][\w$]*(?:\.[A-Za-z_$][\w$]*|\[\d+\])*$/.test(expression) ||
      /^(true|false|null|undefined|NaN|Infinity)$/.test(expression)
    ) {
      unchecked('Only simple property paths can be checked statically');
      return;
    }
    const segments = expression.replace(/\[(\d+)\]/g, '.$1').split('.');
    let value: unknown = this.options.data;
    for (const segment of segments) {
      if (
        value == null ||
        (typeof value !== 'object' && typeof value !== 'string')
      ) {
        this.add(
          'MISSING_FIELD',
          `Cannot resolve field ${expression}`,
          raw,
          node
        );
        return;
      }
      const descriptor = Object.getOwnPropertyDescriptor(
        Object(value),
        segment
      );
      if (!descriptor) {
        this.add('MISSING_FIELD', `Missing own field ${expression}`, raw, node);
        return;
      }
      if (!('value' in descriptor)) {
        unchecked(`Field ${expression} uses an accessor; not checked`);
        return;
      }
      value = descriptor.value;
    }
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
    if (Object.prototype.hasOwnProperty.call(this.options, 'data')) {
      if (type === 'EXEC') {
        this.dynamicContext = true;
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
        if (type !== 'FOR' || match)
          this.checkField(match ? match[2] : code, raw, node);
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
      });
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
  if (!result.diagnostics.length)
    return 'No issues found by the requested checks.';
  return result.diagnostics
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
        (issue.excerpt ? `\n  Context: ${issue.excerpt}` : '')
      );
    })
    .join('\n\n');
}
