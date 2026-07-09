import { appConfig } from "@/lib/config";
import { jsonError, jsonOk } from "@/lib/api/responses";
import { requireServerAccessToken } from "@/lib/auth/session";
import { PlannerService } from "@/lib/graph/planner-service";

export const runtime = "nodejs";

async function photoDataUrl(token: string) {
  try {
    const response = await fetch(`${appConfig.graphBaseUrl}/me/photo/$value`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!response.ok) return null;
    const contentType = response.headers.get("content-type") ?? "image/jpeg";
    const buffer = Buffer.from(await response.arrayBuffer());
    return `data:${contentType};base64,${buffer.toString("base64")}`;
  } catch {
    return null;
  }
}

export async function GET(req: Request) {
  try {
    const token = await requireServerAccessToken(req);
    const service = new PlannerService(token);
    const me = await service.getMe();
    return jsonOk({
      id: me.id,
      name: me.displayName,
      email: me.mail ?? me.userPrincipalName ?? null,
      avatarUrl: await photoDataUrl(token),
    });
  } catch (error) {
    return jsonError(error, "No fue posible consultar el usuario autenticado.");
  }
}
