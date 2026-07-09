import { jsonError, jsonOk } from "@/lib/api/responses";
import { requireServerAccessToken } from "@/lib/auth/session";
import { validateImportRows } from "@/services/import-service";

export const runtime = "nodejs";

export async function POST(req: Request) {
  try {
    const token = await requireServerAccessToken(req);
    const body = await req.json();
    return jsonOk(await validateImportRows(token, body));
  } catch (error) {
    return jsonError(error, "No fue posible validar la importación.");
  }
}
