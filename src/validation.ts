import type { Context, Node } from './types';
import { getCommand, splitCommand } from './processTemplate';

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
    | 'UNCLOSED_COMMAND';
  message: string;
  command: string;
  location: TemplateLocation;
}

export interface ValidationResult {
  /** Structural validity only; expressions are never evaluated. */
  valid: boolean;
  diagnostics: TemplateDiagnostic[];
}

export class TemplateValidator {
  private locations = new Map<Node, TemplateLocation>();
  private blocks: { type: string; name: string; raw: string; node: Node }[] =
    [];

  constructor(
    tree: Node,
    private part: string,
    private diagnostics: TemplateDiagnostic[]
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
    node?: Node
  ) {
    this.diagnostics.push({
      code,
      message,
      command,
      location: (node && this.locations.get(node)) || { part: this.part },
    });
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
