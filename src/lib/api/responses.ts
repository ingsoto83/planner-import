import { ExcelParseError } from "@/lib/excel/parser";
import { GraphApiError, isPlannerPremiumOrUnsupported } from "@/lib/graph/graph-errors";

export function jsonError(error: unknown, fallback = "Ocurrió un error inesperado") {
  if (error instanceof Response) return error;

  if (error instanceof ExcelParseError) {
    return Response.json({ message: error.message }, { status: error.status });
  }

  if (isPlannerPremiumOrUnsupported(error)) {
    return Response.json(
      {
        message:
          "Este Plan no puede utilizarse mediante la API estándar de Microsoft Planner. Selecciona un Plan Basic compatible.",
      },
      { status: 422 },
    );
  }

  if (error instanceof GraphApiError) {
    return Response.json(
      {
        message: error.message,
        code: error.code,
        requestId: error.requestId,
      },
      { status: error.status || 500 },
    );
  }

  if (error instanceof Error) {
    return Response.json({ message: error.message }, { status: 500 });
  }

  return Response.json({ message: fallback }, { status: 500 });
}

export function jsonOk<T>(payload: T, init?: ResponseInit) {
  return Response.json(payload, init);
}
