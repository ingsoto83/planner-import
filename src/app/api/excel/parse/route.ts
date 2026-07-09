import { jsonError, jsonOk } from "@/lib/api/responses";
import { requireServerAccessToken } from "@/lib/auth/session";
import { parseExcelFile } from "@/lib/excel/parser";

export const runtime = "nodejs";

export async function POST(req: Request) {
  try {
    await requireServerAccessToken(req);
    const formData = await req.formData();
    const file = formData.get("file");
    if (!(file instanceof File)) {
      return Response.json({ message: "Carga un archivo XLSX válido." }, { status: 400 });
    }
    return jsonOk(await parseExcelFile(file));
  } catch (error) {
    return jsonError(error, "No fue posible leer el Excel.");
  }
}
