import createReport from './main';
export { listCommands, getMetadata, validateTemplate } from './main';
export type {
  ValidationResult,
  ValidationOptions,
  ValidationCoverage,
  ValidationChecks,
  TemplateDiagnostic,
  TemplateLocation,
} from './validation';
export * from './errors';
import type { QueryResolver } from './types';
export { createReport, QueryResolver };
export default createReport;

export { formatValidationReport } from './validation';

export type { DataSchema } from './schema';
