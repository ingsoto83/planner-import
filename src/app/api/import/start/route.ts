import { jsonError } from "@/lib/api/responses";
import { requireServerAccessToken } from "@/lib/auth/session";
import { runImport } from "@/services/import-service";

export const runtime = "nodejs";

export async function POST(req: Request) {
  try {
    const token = await requireServerAccessToken(req);
    const body = await req.json();
    const encoder = new TextEncoder();

    return new Response(
      new ReadableStream({
        async start(controller) {
          const emit = (event: unknown) => {
            controller.enqueue(encoder.encode(`${JSON.stringify(event)}\n`));
          };

          try {
            await runImport({ token, request: body, emit });
          } catch (error) {
            emit({
              type: "error",
              importId: "unknown",
              message: error instanceof Error ? error.message : "No fue posible completar la importación.",
            });
          } finally {
            controller.close();
          }
        },
      }),
      {
        headers: {
          "Content-Type": "application/x-ndjson; charset=utf-8",
          "Cache-Control": "no-store",
        },
      },
    );
  } catch (error) {
    return jsonError(error, "No fue posible iniciar la importación.");
  }
}
