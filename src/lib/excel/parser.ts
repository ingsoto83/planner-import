import ExcelJS from "exceljs";

import { appConfig } from "@/lib/config";
import { parsePlannerDate } from "@/lib/dates/planner-date";
import type { NormalizedTaskRow, ParseExcelResponse } from "@/types/import";

import { EXPECTED_EXCEL_HEADERS, normalizeTextKey, textFromCell } from "./schema";

const ALLOWED_MIME_TYPES = new Set([
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "application/octet-stream",
  "",
]);

export class ExcelParseError extends Error {
  status = 400;

  constructor(message: string) {
    super(message);
    this.name = "ExcelParseError";
  }
}

function assertXlsxSignature(buffer: ArrayBuffer) {
  const bytes = new Uint8Array(buffer.slice(0, 4));
  const isZip = bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x03 && bytes[3] === 0x04;
  if (!isZip) throw new ExcelParseError("El archivo no parece ser un XLSX válido.");
}

export function validateExcelFile(file: File) {
  const maxBytes = appConfig.maxUploadSizeMb * 1024 * 1024;
  if (!file.name.toLowerCase().endsWith(".xlsx")) {
    throw new ExcelParseError("Solo se aceptan archivos .xlsx.");
  }
  if (!ALLOWED_MIME_TYPES.has(file.type)) {
    throw new ExcelParseError("El tipo de archivo no es compatible. Sube un XLSX válido.");
  }
  if (file.size > maxBytes) {
    throw new ExcelParseError(`El archivo excede el máximo de ${appConfig.maxUploadSizeMb} MB.`);
  }
}

function headerMapFromRow(row: ExcelJS.Row) {
  const map = new Map<string, number>();
  row.eachCell({ includeEmpty: false }, (cell, columnNumber) => {
    map.set(normalizeTextKey(textFromCell(cell.value)), columnNumber);
  });

  for (const expected of EXPECTED_EXCEL_HEADERS) {
    if (!map.has(normalizeTextKey(expected))) {
      throw new ExcelParseError(`Falta la columna obligatoria "${expected}".`);
    }
  }

  return map;
}

function splitResponsables(value: string) {
  return value
    .split(/[;,]/)
    .map((entry) => entry.trim().toLowerCase())
    .filter(Boolean);
}

function normalizeRow(row: ExcelJS.Row, headerMap: Map<string, number>): NormalizedTaskRow | null {
  const get = (header: string) => row.getCell(headerMap.get(normalizeTextKey(header)) ?? 0).value;
  const titulo = textFromCell(get("Titulo"));
  const responsableRaw = textFromCell(get("Responsable"));
  const fechaInicio = parsePlannerDate(get("fecha de inicio"));
  const fechaVencimiento = parsePlannerDate(get("fecha de vencimiento"));
  const fechaInicioRaw = textFromCell(get("fecha de inicio"));
  const fechaVencimientoRaw = textFromCell(get("fecha de vencimiento"));
  const tareas = textFromCell(get("tareas"));
  const etiqueta = textFromCell(get("etiqueta"));

  if (![titulo, responsableRaw, fechaInicio?.date ?? "", fechaVencimiento?.date ?? "", tareas, etiqueta].some(Boolean)) {
    return null;
  }

  return {
    rowNumber: row.number,
    titulo,
    responsables: splitResponsables(responsableRaw),
    responsableRaw,
    fechaInicio: fechaInicio?.date ?? null,
    fechaVencimiento: fechaVencimiento?.date ?? null,
    fechaInicioInvalid: Boolean(fechaInicioRaw && !fechaInicio),
    fechaVencimientoInvalid: Boolean(fechaVencimientoRaw && !fechaVencimiento),
    tareas,
    etiqueta,
    etiquetaNormalizada: normalizeTextKey(etiqueta),
  };
}

export async function parseExcelFile(file: File): Promise<ParseExcelResponse> {
  validateExcelFile(file);
  const buffer = await file.arrayBuffer();
  assertXlsxSignature(buffer);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new ExcelParseError("El archivo no contiene hojas.");

  const headerMap = headerMapFromRow(worksheet.getRow(1));
  const rows: NormalizedTaskRow[] = [];

  for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const normalized = normalizeRow(worksheet.getRow(rowNumber), headerMap);
    if (normalized) rows.push(normalized);
    if (rows.length > appConfig.maxImportRows) {
      throw new ExcelParseError(`El archivo excede el máximo de ${appConfig.maxImportRows} filas.`);
    }
  }

  return {
    fileName: file.name,
    totalRows: rows.length,
    rows,
  };
}
