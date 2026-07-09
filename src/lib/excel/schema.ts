import { z } from "zod";

export const EXPECTED_EXCEL_HEADERS = [
  "Titulo",
  "Responsable",
  "fecha de inicio",
  "fecha de vencimiento",
  "tareas",
  "etiqueta",
] as const;

export const NORMALIZED_EXCEL_HEADERS = EXPECTED_EXCEL_HEADERS.map((header) => normalizeTextKey(header));

export function normalizeTextKey(value: string) {
  return value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();
}

export function textFromCell(value: unknown): string {
  if (value == null) return "";
  if (value instanceof Date) return value.toISOString();
  if (typeof value === "string") return value.trim();
  if (typeof value === "number" || typeof value === "boolean") return String(value).trim();
  if (typeof value === "object") {
    const maybeCell = value as {
      text?: unknown;
      result?: unknown;
      richText?: { text?: string }[];
      hyperlink?: unknown;
    };
    if (typeof maybeCell.text === "string") return maybeCell.text.trim();
    if (maybeCell.result != null) return textFromCell(maybeCell.result);
    if (Array.isArray(maybeCell.richText)) {
      return maybeCell.richText.map((entry) => entry.text ?? "").join("").trim();
    }
  }
  return String(value).trim();
}

export const normalizedRowSchema = z.object({
  rowNumber: z.number().int().positive(),
  titulo: z.string(),
  responsables: z.array(z.string()),
  responsableRaw: z.string(),
  fechaInicio: z.string().nullable(),
  fechaVencimiento: z.string().nullable(),
  fechaInicioInvalid: z.boolean().optional(),
  fechaVencimientoInvalid: z.boolean().optional(),
  tareas: z.string(),
  etiqueta: z.string(),
  etiquetaNormalizada: z.string(),
});

export const importOptionsSchema = z.object({
  planId: z.string().min(1),
  bucketId: z.string().min(1),
  fileName: z.string().min(1),
  createMissingLabels: z.boolean().default(false),
  detectDuplicates: z.boolean().default(true),
  omitDuplicates: z.boolean().default(true),
});

export const importValidationRequestSchema = z.object({
  options: importOptionsSchema,
  rows: z.array(normalizedRowSchema),
});

export const reportRowsSchema = z.array(
  z.object({
    fila: z.number(),
    titulo: z.string(),
    responsable: z.string(),
    fechaInicio: z.string().nullable(),
    fechaVencimiento: z.string().nullable(),
    etiqueta: z.string(),
    estado: z.enum(["Creada", "Omitida", "Error"]),
    mensaje: z.string(),
    plannerTaskId: z.string().optional(),
  }),
);
