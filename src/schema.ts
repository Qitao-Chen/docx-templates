import type { TemplateDiagnostic, ValidationCoverage } from './validation';

/** A small library-specific schema, not a JSON Schema implementation. */
export type DataSchema = {
  required?: boolean;
  nullable?: boolean;
} & (
  | { type: 'string'; minLength?: number; maxLength?: number }
  | { type: 'number' | 'integer'; minimum?: number; maximum?: number }
  | { type: 'boolean' }
  | { type: 'array'; items?: DataSchema; minItems?: number; maxItems?: number }
  | { type: 'object'; properties?: Record<string, DataSchema> }
);

// Read descriptors so configuration/data accessors are never invoked.
function entries(value: unknown): [string, unknown][] {
  if (
    !value ||
    typeof value !== 'object' ||
    Array.isArray(value) ||
    ![Object.prototype, null].includes(Object.getPrototypeOf(value))
  ) {
    throw new TypeError('Schema definitions must be plain objects');
  }
  return Object.getOwnPropertyNames(value).map(key => {
    const descriptor = Object.getOwnPropertyDescriptor(value, key)!;
    if (!('value' in descriptor))
      throw new TypeError('Schema accessors are not supported');
    return [key, descriptor.value];
  });
}

/** Reject invalid rules before any data/template processing. */
export function assertDataSchema(schema: DataSchema): void {
  const pending: { schema: unknown; ancestors: unknown[] }[] = [
    { schema, ancestors: [] },
  ];
  let count = 0;
  while (pending.length) {
    const current = pending.pop()!;
    if (
      ++count > 10000 ||
      current.ancestors.length > 100 ||
      current.ancestors.includes(current.schema)
    ) {
      throw new TypeError(
        'Schema is cyclic or exceeds the definition size/depth limit'
      );
    }
    const rules: Record<string, unknown> = Object.create(null);
    for (const [key, value] of entries(current.schema)) rules[key] = value;
    const allowed: Record<string, string[]> = {
      string: ['minLength', 'maxLength'],
      number: ['minimum', 'maximum'],
      integer: ['minimum', 'maximum'],
      boolean: [],
      array: ['items', 'minItems', 'maxItems'],
      object: ['properties'],
    };
    if (
      typeof rules.type !== 'string' ||
      !Object.prototype.hasOwnProperty.call(allowed, rules.type)
    ) {
      throw new TypeError(
        'Schema type must be string, number, integer, boolean, array or object'
      );
    }
    for (const [key, value] of Object.entries(rules)) {
      if (
        !['type', 'required', 'nullable', ...allowed[rules.type]].includes(key)
      ) {
        throw new TypeError(`Unsupported schema rule: ${key}`);
      }
      if (['required', 'nullable'].includes(key) && typeof value !== 'boolean')
        throw new TypeError(`${key} must be boolean`);
      if (
        ['minimum', 'maximum'].includes(key) &&
        (typeof value !== 'number' || !Number.isFinite(value))
      )
        throw new TypeError(`${key} must be finite`);
      if (
        ['minLength', 'maxLength', 'minItems', 'maxItems'].includes(key) &&
        (typeof value !== 'number' || !Number.isSafeInteger(value) || value < 0)
      )
        throw new TypeError(`${key} must be a nonnegative safe integer`);
    }
    for (const [min, max] of [
      ['minimum', 'maximum'],
      ['minLength', 'maxLength'],
      ['minItems', 'maxItems'],
    ]) {
      if (
        rules[min] !== undefined &&
        rules[max] !== undefined &&
        Number(rules[min]) > Number(rules[max])
      )
        throw new TypeError(`${min} must not exceed ${max}`);
    }
    const ancestors = [...current.ancestors, current.schema];
    if (Object.prototype.hasOwnProperty.call(rules, 'items'))
      pending.push({ schema: rules.items, ancestors });
    if (Object.prototype.hasOwnProperty.call(rules, 'properties')) {
      for (const [, child] of entries(rules.properties))
        pending.push({ schema: child, ancestors });
    }
  }
}

export function validateSchemaData(
  data: unknown,
  schema: DataSchema,
  limit: number
): { diagnostics: TemplateDiagnostic[]; coverage: ValidationCoverage } {
  const diagnostics: TemplateDiagnostic[] = [];
  const coverage = { checked: 0, skipped: 0 };
  const add = (path: string, message: string, skipped = false) => {
    if (skipped) coverage.skipped++;
    diagnostics.push({
      code: skipped ? 'UNCHECKED_SCHEMA' : 'SCHEMA_VIOLATION',
      severity: skipped ? 'warning' : 'error',
      source: 'schema',
      dataPath: path,
      command: '',
      location: { part: 'data' },
      message: `${path || '$'}: ${message}`,
    });
  };
  // Child iteration is lazy: a huge array/object cannot allocate an unbounded work queue.
  function visit(
    value: unknown,
    rule: DataSchema,
    path: string,
    accessor = false
  ): boolean {
    if (coverage.checked + coverage.skipped >= limit) {
      add(path, 'Schema check limit reached; remaining data not checked', true);
      return false;
    }
    if (accessor) {
      add(path, 'Accessor value not checked', true);
      return true;
    }
    coverage.checked++;
    if (value === undefined) {
      if (rule.required) add(path, 'Required value is missing');
      return true;
    }
    if (value === null) {
      if (!rule.nullable) add(path, 'Null is not allowed');
      return true;
    }
    const matches =
      rule.type === 'array'
        ? Array.isArray(value)
        : rule.type === 'object'
        ? typeof value === 'object' && !Array.isArray(value)
        : rule.type === 'integer'
        ? typeof value === 'number' && Number.isInteger(value)
        : rule.type === 'number'
        ? typeof value === 'number' && Number.isFinite(value)
        : typeof value === rule.type;
    if (!matches) {
      add(path, `Expected ${rule.type}`);
      return true;
    }
    if (rule.type === 'string') {
      const length = Array.from(value as string).length;
      if (rule.minLength !== undefined && length < rule.minLength)
        add(path, `String is shorter than ${rule.minLength} code points`);
      if (rule.maxLength !== undefined && length > rule.maxLength)
        add(path, `String is longer than ${rule.maxLength} code points`);
    } else if (rule.type === 'number' || rule.type === 'integer') {
      if (rule.minimum !== undefined && (value as number) < rule.minimum)
        add(path, `Number is below ${rule.minimum}`);
      if (rule.maximum !== undefined && (value as number) > rule.maximum)
        add(path, `Number exceeds ${rule.maximum}`);
    } else if (rule.type === 'object' && rule.properties) {
      for (const [key, child] of Object.entries(rule.properties)) {
        const nextPath = /^[A-Za-z_$][\w$]*$/.test(key)
          ? path
            ? `${path}.${key}`
            : key
          : `${path}[${JSON.stringify(key)}]`;
        const descriptor = Object.getOwnPropertyDescriptor(value, key);
        if (
          !visit(
            descriptor && 'value' in descriptor ? descriptor.value : undefined,
            child,
            nextPath,
            !!descriptor && !('value' in descriptor)
          )
        )
          return false;
      }
    } else if (rule.type === 'array') {
      const array = value as unknown[];
      if (rule.minItems !== undefined && array.length < rule.minItems)
        add(path, `Array has fewer than ${rule.minItems} items`);
      if (rule.maxItems !== undefined && array.length > rule.maxItems)
        add(path, `Array has more than ${rule.maxItems} items`);
      if (rule.items)
        for (let index = 0; index < array.length; index++) {
          const descriptor = Object.getOwnPropertyDescriptor(
            array,
            String(index)
          );
          if (
            !visit(
              descriptor && 'value' in descriptor
                ? descriptor.value
                : undefined,
              rule.items,
              `${path}[${index}]`,
              !!descriptor && !('value' in descriptor)
            )
          )
            return false;
        }
    }
    return true;
  }
  visit(data, schema, '');
  return { diagnostics, coverage };
}
