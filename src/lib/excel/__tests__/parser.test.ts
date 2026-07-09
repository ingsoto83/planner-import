import { describe, expect, it } from "vitest";
import ExcelJS from "exceljs";

import { parseExcelFile } from "../parser";

async function workbookFile(rows: unknown[][]) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Tareas");
  rows.forEach((row) => worksheet.addRow(row));
  const buffer = await workbook.xlsx.writeBuffer();
  return new File([buffer], "planner.xlsx", {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

describe("parseExcelFile", () => {
  it("accepts headers with spacing and casing differences", async () => {
    const file = await workbookFile([
      ["  TITULO  ", "Responsable", "fecha  de inicio", "fecha de vencimiento", "tareas", "etiqueta"],
      ["Revisar ASN", "maria.trevino@empresa.com", "14/07/2026", 46218, "Detalle", " Desarrollo "],
    ]);

    const parsed = await parseExcelFile(file);

    expect(parsed.totalRows).toBe(1);
    expect(parsed.rows[0]).toMatchObject({
      titulo: "Revisar ASN",
      responsables: ["maria.trevino@empresa.com"],
      fechaInicio: "2026-07-14",
      fechaVencimiento: "2026-07-15",
      etiqueta: "Desarrollo",
      etiquetaNormalizada: "desarrollo",
    });
  });

  it("marks invalid date cells", async () => {
    const file = await workbookFile([
      ["Titulo", "Responsable", "fecha de inicio", "fecha de vencimiento", "tareas", "etiqueta"],
      ["Actualizar SSL", "", "99/99/2026", "", "", ""],
    ]);

    const parsed = await parseExcelFile(file);
    expect(parsed.rows[0]?.fechaInicioInvalid).toBe(true);
  });
});
