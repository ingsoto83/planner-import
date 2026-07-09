import ExcelJS from "exceljs";
import { DateTime } from "luxon";

import { appConfig } from "@/lib/config";
import type { ImportResultRow } from "@/types/import";

export async function buildImportReport(rows: ImportResultRow[]) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = appConfig.appName;
  workbook.created = new Date();

  const worksheet = workbook.addWorksheet("Resultado");
  worksheet.columns = [
    { header: "fila", key: "fila", width: 10 },
    { header: "titulo", key: "titulo", width: 32 },
    { header: "responsable", key: "responsable", width: 36 },
    { header: "fechaInicio", key: "fechaInicio", width: 18 },
    { header: "fechaVencimiento", key: "fechaVencimiento", width: 20 },
    { header: "etiqueta", key: "etiqueta", width: 24 },
    { header: "estado", key: "estado", width: 14 },
    { header: "mensaje", key: "mensaje", width: 48 },
    { header: "plannerTaskId", key: "plannerTaskId", width: 32 },
  ];
  worksheet.getRow(1).font = { bold: true };

  rows.forEach((row) => worksheet.addRow(row));

  const buffer = await workbook.xlsx.writeBuffer();
  const stamp = DateTime.now().setZone(appConfig.timezone).toFormat("yyyyLLdd-HHmm");
  return {
    buffer,
    fileName: `planner-import-result-${stamp}.xlsx`,
  };
}
