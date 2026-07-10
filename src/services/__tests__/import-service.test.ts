import { afterAll, afterEach, beforeAll, describe, expect, it } from "vitest";
import { http, HttpResponse } from "msw";
import { setupServer } from "msw/node";

import { runImport, validateImportRows } from "../import-service";

const server = setupServer();

beforeAll(() => server.listen());
afterEach(() => server.resetHandlers());
afterAll(() => server.close());

describe("validateImportRows", () => {
  it("validates users, labels, dates, title and duplicates", async () => {
    server.use(
      http.get("https://graph.microsoft.com/v1.0/planner/plans/plan-1/details", () =>
        HttpResponse.json({
          "@odata.etag": 'W/"plan"',
          categoryDescriptions: {
            category1: "Desarrollo",
            category2: "Infraestructura",
          },
        }),
      ),
      http.get("https://graph.microsoft.com/v1.0/planner/buckets/bucket-1/tasks", () =>
        HttpResponse.json({
          value: [
            {
              id: "task-1",
              planId: "plan-1",
              bucketId: "bucket-1",
              title: "Revisar ASN",
              dueDateTime: "2026-07-15T18:00:00Z",
            },
          ],
        }),
      ),
      http.get("https://graph.microsoft.com/v1.0/users/maria.trevino%40empresa.com", () =>
        HttpResponse.json({
          id: "user-1",
          displayName: "María Treviño",
          mail: "maria.trevino@empresa.com",
          userPrincipalName: "maria.trevino@empresa.com",
        }),
      ),
      http.get("https://graph.microsoft.com/v1.0/users/noexiste%40empresa.com", () =>
        HttpResponse.json({ error: { code: "Request_ResourceNotFound", message: "Not found" } }, { status: 404 }),
      ),
      http.get("https://graph.microsoft.com/v1.0/me", () =>
        HttpResponse.json({
          id: "current-user",
          displayName: "Current User",
          mail: "current.user@empresa.com",
          userPrincipalName: "current.user@empresa.com",
        }),
      ),
      http.get("https://graph.microsoft.com/v1.0/users", () => HttpResponse.json({ value: [] })),
    );

    const response = await validateImportRows("token", {
      options: {
        planId: "plan-1",
        bucketId: "bucket-1",
        fileName: "planner.xlsx",
        createMissingLabels: false,
        detectDuplicates: true,
        omitDuplicates: true,
      },
      rows: [
        {
          rowNumber: 2,
          titulo: "Revisar ASN",
          responsables: ["maria.trevino@empresa.com"],
          responsableRaw: "maria.trevino@empresa.com",
          fechaInicio: "2026-07-14",
          fechaVencimiento: "2026-07-15",
          tareas: "Detalle",
          etiqueta: "DESARROLLO",
          etiquetaNormalizada: "desarrollo",
        },
        {
          rowNumber: 3,
          titulo: "",
          responsables: ["noexiste@empresa.com"],
          responsableRaw: "noexiste@empresa.com",
          fechaInicio: "2026-07-16",
          fechaVencimiento: "2026-07-15",
          fechaInicioInvalid: false,
          fechaVencimientoInvalid: false,
          tareas: "",
          etiqueta: "No configurada",
          etiquetaNormalizada: "no configurada",
        },
      ],
    });

    expect(response.summary.totalRows).toBe(2);
    expect(response.rows[0]?.status).toBe("warning");
    expect(response.rows[0]?.shouldOmit).toBe(true);
    expect(response.rows[0]?.labelKey).toBe("category1");
    expect(response.rows[1]?.status).toBe("invalid");
    expect(response.rows[1]?.issues.map((entry) => entry.field)).toEqual(
      expect.arrayContaining(["titulo", "responsable", "fechaVencimiento", "etiqueta"]),
    );
  });

  it("reports a clear message when Planner rejects a resolved assignee", async () => {
    server.use(
      http.get("https://graph.microsoft.com/v1.0/planner/plans/plan-1/details", () =>
        HttpResponse.json({
          "@odata.etag": 'W/"plan"',
          categoryDescriptions: {
            category1: "Desarrollo",
          },
        }),
      ),
      http.get("https://graph.microsoft.com/v1.0/planner/buckets/bucket-1/tasks", () =>
        HttpResponse.json({ value: [] }),
      ),
      http.get("https://graph.microsoft.com/v1.0/users/externo%40empresa.com", () =>
        HttpResponse.json({
          id: "user-2",
          displayName: "Usuario Externo",
          mail: "externo@empresa.com",
          userPrincipalName: "externo@empresa.com",
        }),
      ),
      http.post("https://graph.microsoft.com/v1.0/planner/tasks", () =>
        HttpResponse.json(
          {
            error: {
              code: "Forbidden",
              message:
                "You do not have the required permissions to access this item, or the item may not exist.",
              innerError: { "request-id": "request-1" },
            },
          },
          { status: 403 },
        ),
      ),
    );

    const events: unknown[] = [];
    await runImport({
      token: "token",
      emit: (event) => events.push(event),
      request: {
        options: {
          planId: "plan-1",
          bucketId: "bucket-1",
          fileName: "planner.xlsx",
          createMissingLabels: false,
          detectDuplicates: true,
          omitDuplicates: true,
        },
        rows: [
          {
            rowNumber: 2,
            titulo: "Asignar responsable",
            responsables: ["externo@empresa.com"],
            responsableRaw: "externo@empresa.com",
            fechaInicio: "2026-07-14",
            fechaVencimiento: "2026-07-15",
            tareas: "",
            etiqueta: "Desarrollo",
            etiquetaNormalizada: "desarrollo",
          },
        ],
      },
    });

    const completed = events.find(
      (event): event is { type: "completed"; results: { estado: string; mensaje: string }[] } =>
        typeof event === "object" && event !== null && "type" in event && event.type === "completed",
    );

    expect(completed?.results[0]?.estado).toBe("Error");
    expect(completed?.results[0]?.mensaje).toContain("no permite asignarlo en este Plan/Bucket");
    expect(completed?.results[0]?.mensaje).toContain("request-1");
  });

  it("falls back to /me when a current external user cannot be resolved by /users/{mail}", async () => {
    server.use(
      http.get("https://graph.microsoft.com/v1.0/planner/plans/plan-1/details", () =>
        HttpResponse.json({
          "@odata.etag": 'W/"plan"',
          categoryDescriptions: {
            category1: "Desarrollo",
          },
        }),
      ),
      http.get("https://graph.microsoft.com/v1.0/planner/buckets/bucket-1/tasks", () =>
        HttpResponse.json({ value: [] }),
      ),
      http.get("https://graph.microsoft.com/v1.0/users/maria.trevino%40servicioexterno.com.mx", () =>
        HttpResponse.json({ error: { code: "BadRequest", message: "Bad request" } }, { status: 400 }),
      ),
      http.get("https://graph.microsoft.com/v1.0/me", () =>
        HttpResponse.json({
          id: "current-user",
          displayName: "Treviño Garcia Maria Fernanda",
          mail: "maria.trevino@servicioexterno.com.mx",
          userPrincipalName: "maria.trevino_servicioexterno.com.mx#EXT#@tenant.onmicrosoft.com",
        }),
      ),
    );

    const response = await validateImportRows("token", {
      options: {
        planId: "plan-1",
        bucketId: "bucket-1",
        fileName: "planner.xlsx",
        createMissingLabels: false,
        detectDuplicates: true,
        omitDuplicates: true,
      },
      rows: [
        {
          rowNumber: 2,
          titulo: "Tarea usuario actual",
          responsables: ["maria.trevino@servicioexterno.com.mx"],
          responsableRaw: "maria.trevino@servicioexterno.com.mx",
          fechaInicio: "2026-07-14",
          fechaVencimiento: "2026-07-15",
          tareas: "",
          etiqueta: "Desarrollo",
          etiquetaNormalizada: "desarrollo",
        },
      ],
    });

    expect(response.rows[0]?.status).toBe("valid");
    expect(response.rows[0]?.resolvedUsers[0]?.id).toBe("current-user");
  });

  it("keeps preview usable when duplicate detection returns Graph 400", async () => {
    server.use(
      http.get("https://graph.microsoft.com/v1.0/planner/plans/plan-1/details", () =>
        HttpResponse.json({
          "@odata.etag": 'W/"plan"',
          categoryDescriptions: {
            category1: "Desarrollo",
          },
        }),
      ),
      http.get("https://graph.microsoft.com/v1.0/planner/buckets/bucket-1/tasks", () =>
        HttpResponse.json(
          {
            error: {
              code: "BadRequest",
              message: "Invalid bucket query",
              innerError: { "request-id": "duplicate-request" },
            },
          },
          { status: 400 },
        ),
      ),
    );

    const response = await validateImportRows("token", {
      options: {
        planId: "plan-1",
        bucketId: "bucket-1",
        fileName: "planner.xlsx",
        createMissingLabels: false,
        detectDuplicates: true,
        omitDuplicates: true,
      },
      rows: [
        {
          rowNumber: 2,
          titulo: "Tarea sin responsable",
          responsables: [],
          responsableRaw: "",
          fechaInicio: "2026-07-14",
          fechaVencimiento: "2026-07-15",
          tareas: "",
          etiqueta: "Desarrollo",
          etiquetaNormalizada: "desarrollo",
        },
      ],
    });

    expect(response.rows[0]?.status).toBe("warning");
    expect(response.rows[0]?.issues[0]?.message).toContain("No fue posible detectar duplicados");
    expect(response.rows[0]?.issues[0]?.message).toContain("duplicate-request");
  });
});
