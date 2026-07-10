import { randomUUID } from "node:crypto";

import { DateTime } from "luxon";
import pLimit from "p-limit";

import { appConfig } from "@/lib/config";
import { compareDateOnly, graphDateTimeFromDateOnly } from "@/lib/dates/planner-date";
import { importValidationRequestSchema, normalizeTextKey } from "@/lib/excel/schema";
import { GraphApiError, isGraphApiError } from "@/lib/graph/graph-errors";
import { PlannerService, labelMapFromDescriptions } from "@/lib/graph/planner-service";
import { UsersService } from "@/lib/graph/users-service";
import type {
  ImportOptions,
  ImportProgressEvent,
  ImportResultRow,
  ImportValidationRequest,
  ImportValidationResponse,
  NormalizedTaskRow,
  ValidatedImportRow,
  ValidationIssue,
} from "@/types/import";
import type { GraphPlannerTask, ResolvedGraphUser } from "@/types/planner";

function localDateFromGraphDateTime(value?: string | null) {
  if (!value) return null;
  const parsed = DateTime.fromISO(value, { zone: "utc" }).setZone(appConfig.timezone);
  return parsed.isValid ? parsed.toISODate() : null;
}

function duplicateKey(title: string, dueDate: string | null) {
  return `${normalizeTextKey(title)}|${dueDate ?? ""}`;
}

function summarize(rows: ValidatedImportRow[]) {
  return {
    totalRows: rows.length,
    validRows: rows.filter((row) => row.status === "valid").length,
    warningRows: rows.filter((row) => row.status === "warning").length,
    errorRows: rows.filter((row) => row.status === "invalid").length,
    omittedRows: rows.filter((row) => row.shouldOmit).length,
  };
}

function issue(field: ValidationIssue["field"], severity: ValidationIssue["severity"], message: string): ValidationIssue {
  return { field, severity, message };
}

async function resolveUsers(row: NormalizedTaskRow, usersService: UsersService) {
  const users: ResolvedGraphUser[] = [];
  const issues: ValidationIssue[] = [];

  for (const responsable of row.responsables) {
    const user = await usersService.resolveUser(responsable);
    if (!user) {
      issues.push(issue("responsable", "error", "Responsable no encontrado en Microsoft Entra ID"));
    } else {
      users.push(user);
    }
  }

  return { users, issues };
}

function validationIssues(row: NormalizedTaskRow) {
  const issues: ValidationIssue[] = [];

  if (!row.titulo.trim()) issues.push(issue("titulo", "error", "El título está vacío."));
  if (row.fechaInicioInvalid) issues.push(issue("fechaInicio", "error", "La fecha de inicio no es válida."));
  if (row.fechaVencimientoInvalid) {
    issues.push(issue("fechaVencimiento", "error", "La fecha de vencimiento no es válida."));
  }
  if (compareDateOnly(row.fechaInicio, row.fechaVencimiento) > 0) {
    issues.push(issue("fechaVencimiento", "error", "La fecha de vencimiento es anterior a la fecha de inicio."));
  }

  return issues;
}

function statusForIssues(issues: ValidationIssue[]) {
  if (issues.some((entry) => entry.severity === "error")) return "invalid" as const;
  if (issues.some((entry) => entry.severity === "warning")) return "warning" as const;
  return "valid" as const;
}

async function resolveLabelMap(options: ImportOptions, rows: NormalizedTaskRow[], plannerService: PlannerService) {
  let details = await plannerService.getPlanDetails(options.planId);
  let labelMap = labelMapFromDescriptions(details.categoryDescriptions);
  const missingLabels = rows
    .filter((row) => row.etiqueta.trim())
    .filter((row) => !labelMap.has(row.etiquetaNormalizada))
    .map((row) => row.etiqueta);

  if (options.createMissingLabels && missingLabels.length > 0) {
    details = await plannerService.ensureLabels(options.planId, missingLabels);
    labelMap = labelMapFromDescriptions(details.categoryDescriptions);
  }

  return labelMap;
}

async function duplicateMap(options: ImportOptions, plannerService: PlannerService) {
  if (!options.detectDuplicates) return { duplicates: new Map<string, GraphPlannerTask>() };

  try {
    const existingTasks = await plannerService.listBucketTasks(options.bucketId);
    return {
      duplicates: new Map(
        existingTasks.map((task) => [
          duplicateKey(task.title, localDateFromGraphDateTime(task.dueDateTime)),
          task,
        ]),
      ),
    };
  } catch (error) {
    if (error instanceof GraphApiError && [400, 403, 404].includes(error.status)) {
      const requestId = error.requestId ? ` (requestId: ${error.requestId})` : "";
      return {
        duplicates: new Map<string, GraphPlannerTask>(),
        warning: issue(
          "duplicado",
          "warning",
          `No fue posible detectar duplicados en este Bucket. La validación continúa sin bloquear la importación.${requestId}`,
        ),
      };
    }
    throw error;
  }
}

export async function validateImportRows(token: string, request: ImportValidationRequest): Promise<ImportValidationResponse> {
  const parsed = importValidationRequestSchema.parse(request);
  const importId = randomUUID();
  const plannerService = new PlannerService(token);
  const usersService = new UsersService(token);
  const labelMap = await resolveLabelMap(parsed.options, parsed.rows, plannerService);
  const duplicateResult = await duplicateMap(parsed.options, plannerService);

  const rows: ValidatedImportRow[] = [];

  for (const row of parsed.rows) {
    const issues = validationIssues(row);
    if (duplicateResult.warning) issues.push(duplicateResult.warning);
    const { users, issues: userIssues } = await resolveUsers(row, usersService);
    issues.push(...userIssues);

    const labelKey = row.etiquetaNormalizada ? labelMap.get(row.etiquetaNormalizada) : undefined;
    if (row.etiqueta.trim() && !labelKey) {
      issues.push(issue("etiqueta", "error", `Etiqueta "${row.etiqueta}" no existe en el Plan seleccionado.`));
    }

    const duplicate = duplicateResult.duplicates.get(duplicateKey(row.titulo, row.fechaVencimiento));
    const shouldOmit = Boolean(duplicate && parsed.options.omitDuplicates);
    if (duplicate) {
      issues.push(
        issue(
          "duplicado",
          "warning",
          "Posible duplicado: ya existe una tarea con el mismo título y fecha de vencimiento.",
        ),
      );
    }

    rows.push({
      ...row,
      status: statusForIssues(issues),
      issues,
      resolvedUsers: users,
      labelKey,
      duplicateTaskId: duplicate?.id,
      shouldOmit,
    });
  }

  return {
    importId,
    options: parsed.options,
    rows,
    summary: summarize(rows),
  };
}

function rowToOmittedResult(row: ValidatedImportRow, message?: string): ImportResultRow {
  return {
    fila: row.rowNumber,
    titulo: row.titulo,
    responsable: row.responsableRaw,
    fechaInicio: row.fechaInicio,
    fechaVencimiento: row.fechaVencimiento,
    etiqueta: row.etiqueta,
    estado: "Omitida",
    mensaje: message ?? (row.issues.map((entry) => entry.message).join(" | ") || "Fila omitida."),
  };
}

function graphMessage(error: unknown) {
  if (isGraphApiError(error)) {
    return error.requestId ? `${error.message} (requestId: ${error.requestId})` : error.message;
  }
  return error instanceof Error ? error.message : "Error desconocido durante la creación.";
}

function assignmentGraphMessage(error: unknown, row: ValidatedImportRow) {
  if (!row.resolvedUsers.length || !isGraphApiError(error)) return graphMessage(error);

  const graphText = `${error.code ?? ""} ${error.message}`.toLowerCase();
  const looksLikeAssignmentAccessError =
    (error.status === 403 || error.status === 404) &&
    (graphText.includes("required permissions") ||
      graphText.includes("may not exist") ||
      graphText.includes("access this item") ||
      graphText.includes("forbidden"));

  if (!looksLikeAssignmentAccessError) return graphMessage(error);

  const requestId = error.requestId ? ` (requestId: ${error.requestId})` : "";
  return `El responsable existe en Microsoft Entra ID, pero Planner no permite asignarlo en este Plan/Bucket. Agrega al usuario como miembro del Plan o del grupo de Microsoft 365 asociado, o deja Responsable vacío.${requestId}`;
}

function taskBody(row: ValidatedImportRow, options: ImportOptions) {
  const assignments = Object.fromEntries(
    row.resolvedUsers.map((user) => [
      user.id,
      {
        "@odata.type": "#microsoft.graph.plannerAssignment",
        orderHint: " !",
      },
    ]),
  );
  const appliedCategories = row.labelKey ? { [row.labelKey]: true } : undefined;

  return {
    planId: options.planId,
    bucketId: options.bucketId,
    title: row.titulo,
    startDateTime: graphDateTimeFromDateOnly(row.fechaInicio),
    dueDateTime: graphDateTimeFromDateOnly(row.fechaVencimiento),
    ...(Object.keys(assignments).length ? { assignments } : {}),
    ...(appliedCategories ? { appliedCategories } : {}),
  };
}

async function importRow(row: ValidatedImportRow, options: ImportOptions, plannerService: PlannerService): Promise<ImportResultRow> {
  try {
    const task = await plannerService.createTask(taskBody(row, options));
    if (row.tareas.trim()) {
      await plannerService.updateTaskDetailsDescription(task.id, row.tareas.trim());
    }
    return {
      fila: row.rowNumber,
      titulo: row.titulo,
      responsable: row.responsableRaw,
      fechaInicio: row.fechaInicio,
      fechaVencimiento: row.fechaVencimiento,
      etiqueta: row.etiqueta,
      estado: "Creada",
      mensaje: "Task ID creado correctamente",
      plannerTaskId: task.id,
    };
  } catch (error) {
    return {
      fila: row.rowNumber,
      titulo: row.titulo,
      responsable: row.responsableRaw,
      fechaInicio: row.fechaInicio,
      fechaVencimiento: row.fechaVencimiento,
      etiqueta: row.etiqueta,
      estado: "Error",
      mensaje: assignmentGraphMessage(error, row),
    };
  }
}

export async function runImport({
  token,
  request,
  emit,
}: {
  token: string;
  request: ImportValidationRequest;
  emit: (event: ImportProgressEvent) => void;
}) {
  const validation = await validateImportRows(token, request);
  const importId = validation.importId;
  const plannerService = new PlannerService(token);
  const importableRows = validation.rows.filter((row) => row.status !== "invalid" && !row.shouldOmit);
  const results: ImportResultRow[] = validation.rows
    .filter((row) => row.status === "invalid" || row.shouldOmit)
    .map((row) =>
      row.shouldOmit
        ? rowToOmittedResult(row, "Omitida por posible duplicado.")
        : rowToOmittedResult(row),
    );

  let processed = 0;
  let successCount = 0;
  let omittedCount = results.length;
  let errorCount = 0;
  const startedAt = Date.now();

  emit({ type: "started", importId, total: importableRows.length });

  const limit = pLimit(appConfig.importConcurrency);
  await Promise.all(
    importableRows.map((row) =>
      limit(async () => {
        emit({
          type: "row-started",
          importId,
          rowNumber: row.rowNumber,
          title: row.titulo,
          processed,
          total: importableRows.length,
        });
        const result = await importRow(row, validation.options, plannerService);
        results.push(result);
        processed += 1;
        if (result.estado === "Creada") successCount += 1;
        if (result.estado === "Omitida") omittedCount += 1;
        if (result.estado === "Error") errorCount += 1;

        emit({ type: "row-result", importId, row: result, processed, total: importableRows.length });
        emit({
          type: "progress",
          importId,
          processed,
          total: importableRows.length,
          successCount,
          omittedCount,
          errorCount,
        });
      }),
    ),
  );

  const sortedResults = results.sort((a, b) => a.fila - b.fila);

  console.info(
    JSON.stringify({
      event: "planner_import_completed",
      importId,
      planId: validation.options.planId,
      bucketId: validation.options.bucketId,
      totalRows: validation.rows.length,
      validRows: importableRows.length,
      successCount,
      errorCount,
      omittedCount,
      durationMs: Date.now() - startedAt,
    }),
  );

  emit({
    type: "completed",
    importId,
    processed,
    total: importableRows.length,
    successCount,
    omittedCount,
    errorCount,
    results: sortedResults,
  });
}
