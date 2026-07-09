import type { PlannerLabelKey, ResolvedGraphUser } from "./planner";

export type ExcelTaskRow = {
  rowNumber: number;
  titulo: unknown;
  responsable: unknown;
  fechaInicio: unknown;
  fechaVencimiento: unknown;
  tareas: unknown;
  etiqueta: unknown;
};

export type NormalizedTaskRow = {
  rowNumber: number;
  titulo: string;
  responsables: string[];
  responsableRaw: string;
  fechaInicio: string | null;
  fechaVencimiento: string | null;
  fechaInicioInvalid?: boolean;
  fechaVencimientoInvalid?: boolean;
  tareas: string;
  etiqueta: string;
  etiquetaNormalizada: string;
};

export type ValidationSeverity = "error" | "warning";

export type ValidationIssue = {
  field:
    | "titulo"
    | "responsable"
    | "fechaInicio"
    | "fechaVencimiento"
    | "etiqueta"
    | "duplicado"
    | "archivo";
  severity: ValidationSeverity;
  message: string;
};

export type ValidatedImportRow = NormalizedTaskRow & {
  status: "valid" | "invalid" | "warning";
  issues: ValidationIssue[];
  resolvedUsers: ResolvedGraphUser[];
  labelKey?: PlannerLabelKey;
  duplicateTaskId?: string;
  shouldOmit?: boolean;
};

export type ImportOptions = {
  planId: string;
  bucketId: string;
  fileName: string;
  createMissingLabels: boolean;
  detectDuplicates: boolean;
  omitDuplicates: boolean;
};

export type ImportValidationRequest = {
  options: ImportOptions;
  rows: NormalizedTaskRow[];
};

export type ImportResultRow = {
  fila: number;
  titulo: string;
  responsable: string;
  fechaInicio: string | null;
  fechaVencimiento: string | null;
  etiqueta: string;
  estado: "Creada" | "Omitida" | "Error";
  mensaje: string;
  plannerTaskId?: string;
};

export type ImportProgressEvent =
  | {
      type: "started";
      importId: string;
      total: number;
    }
  | {
      type: "row-started";
      importId: string;
      rowNumber: number;
      title: string;
      processed: number;
      total: number;
    }
  | {
      type: "row-result";
      importId: string;
      row: ImportResultRow;
      processed: number;
      total: number;
    }
  | {
      type: "progress";
      importId: string;
      processed: number;
      total: number;
      successCount: number;
      omittedCount: number;
      errorCount: number;
    }
  | {
      type: "completed";
      importId: string;
      processed: number;
      total: number;
      successCount: number;
      omittedCount: number;
      errorCount: number;
      results: ImportResultRow[];
    }
  | {
      type: "error";
      importId: string;
      message: string;
    };

export type ParseExcelResponse = {
  fileName: string;
  totalRows: number;
  rows: NormalizedTaskRow[];
};

export type ValidationSummary = {
  totalRows: number;
  validRows: number;
  warningRows: number;
  errorRows: number;
  omittedRows: number;
};

export type ImportValidationResponse = {
  importId: string;
  options: ImportOptions;
  rows: ValidatedImportRow[];
  summary: ValidationSummary;
};
