import { jsonError, jsonOk } from "@/lib/api/responses";
import { requireServerAccessToken } from "@/lib/auth/session";
import { PlannerService } from "@/lib/graph/planner-service";

export const runtime = "nodejs";

export async function GET(req: Request, { params }: { params: Promise<{ planId: string }> }) {
  try {
    const { planId } = await params;
    const token = await requireServerAccessToken(req);
    const service = new PlannerService(token);
    return jsonOk({ value: await service.getLabels(planId) });
  } catch (error) {
    return jsonError(error, "No fue posible consultar las etiquetas del Plan.");
  }
}
