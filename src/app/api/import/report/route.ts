import { jsonError } from "@/lib/api/responses";
import { requireServerAccessToken } from "@/lib/auth/session";
import { buildImportReport } from "@/lib/excel/report";
import { reportRowsSchema } from "@/lib/excel/schema";

export const runtime = "nodejs";

export async function POST(req: Request) {
  try {
    await requireServerAccessToken(req);
    const rows = reportRowsSchema.parse(await req.json());
    const report = await buildImportReport(rows);

    return new Response(report.buffer as BodyInit, {
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="${report.fileName}"`,
        "Cache-Control": "no-store",
      },
    });
  } catch (error) {
    return jsonError(error, "No fue posible generar el reporte.");
  }
}
