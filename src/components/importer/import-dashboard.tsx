"use client";

import * as React from "react";
import {
  AlertCircle,
  ArrowDownToLine,
  CheckCircle2,
  FileSpreadsheet,
  Loader2,
  LogOut,
  Play,
  RefreshCcw,
  Search,
  UploadCloud,
  XCircle,
} from "lucide-react";
import {
  ColumnDef,
  flexRender,
  getCoreRowModel,
  getFilteredRowModel,
  useReactTable,
} from "@tanstack/react-table";
import Image from "next/image";
import { signOut } from "next-auth/react";

import { Avatar } from "@/components/ui/avatar";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Checkbox } from "@/components/ui/checkbox";
import { Progress } from "@/components/ui/progress";
import { Select } from "@/components/ui/select";
import { Skeleton } from "@/components/ui/skeleton";
import type {
  ImportOptions,
  ImportProgressEvent,
  ImportResultRow,
  ImportValidationResponse,
  NormalizedTaskRow,
  ParseExcelResponse,
  ValidatedImportRow,
} from "@/types/import";
import type { PlannerBucket, PlannerLabel, PlannerPlan } from "@/types/planner";

type ApiList<T> = { value: T[] };
type ApiUser = { id: string; name: string; email?: string | null; avatarUrl?: string | null };
type FilterMode = "all" | "valid" | "errors" | "warnings";

const filterLabels: Record<FilterMode, string> = {
  all: "Todas",
  valid: "Válidas",
  errors: "Errores",
  warnings: "Advertencias",
};

async function fetchJson<T>(input: RequestInfo | URL, init?: RequestInit): Promise<T> {
  const response = await fetch(input, init);
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(payload.message ?? "La solicitud no pudo completarse.");
  }
  return payload as T;
}

function statusBadge(row: ValidatedImportRow) {
  if (row.status === "invalid") return <Badge tone="error"><XCircle className="h-3 w-3" />Error</Badge>;
  if (row.status === "warning") return <Badge tone="warning"><AlertCircle className="h-3 w-3" />Advertencia</Badge>;
  return <Badge tone="success"><CheckCircle2 className="h-3 w-3" />Válida</Badge>;
}

function displayDate(value: string | null) {
  if (!value) return "";
  const [year, month, day] = value.split("-");
  return year && month && day ? `${day}/${month}/${year}` : value;
}

function resultTone(row: ImportResultRow) {
  if (row.estado === "Creada") return "success" as const;
  if (row.estado === "Omitida") return "warning" as const;
  return "error" as const;
}

export function ImportDashboard({ initialUser }: { initialUser: { name?: string | null; email?: string | null } }) {
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const [user, setUser] = React.useState<ApiUser | null>(null);
  const [plans, setPlans] = React.useState<PlannerPlan[]>([]);
  const [buckets, setBuckets] = React.useState<PlannerBucket[]>([]);
  const [labels, setLabels] = React.useState<PlannerLabel[]>([]);
  const [selectedPlanId, setSelectedPlanId] = React.useState("");
  const [selectedBucketId, setSelectedBucketId] = React.useState("");
  const [loadingPlans, setLoadingPlans] = React.useState(true);
  const [loadingBuckets, setLoadingBuckets] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);
  const [fileData, setFileData] = React.useState<ParseExcelResponse | null>(null);
  const [validation, setValidation] = React.useState<ImportValidationResponse | null>(null);
  const [validating, setValidating] = React.useState(false);
  const [uploading, setUploading] = React.useState(false);
  const [filter, setFilter] = React.useState<FilterMode>("all");
  const [createMissingLabels, setCreateMissingLabels] = React.useState(false);
  const [detectDuplicates, setDetectDuplicates] = React.useState(true);
  const [omitDuplicates, setOmitDuplicates] = React.useState(true);
  const [confirming, setConfirming] = React.useState(false);
  const [importing, setImporting] = React.useState(false);
  const [progress, setProgress] = React.useState({ processed: 0, total: 0, success: 0, omitted: 0, errors: 0 });
  const [activity, setActivity] = React.useState<string[]>([]);
  const [results, setResults] = React.useState<ImportResultRow[]>([]);

  const selectedPlan = plans.find((plan) => plan.id === selectedPlanId);
  const selectedBucket = buckets.find((bucket) => bucket.id === selectedBucketId);

  const options = React.useMemo<ImportOptions | null>(() => {
    if (!selectedPlanId || !selectedBucketId || !fileData) return null;
    return {
      planId: selectedPlanId,
      bucketId: selectedBucketId,
      fileName: fileData.fileName,
      createMissingLabels,
      detectDuplicates,
      omitDuplicates,
    };
  }, [createMissingLabels, detectDuplicates, fileData, omitDuplicates, selectedBucketId, selectedPlanId]);

  const loadPlans = React.useCallback(async () => {
    setLoadingPlans(true);
    setError(null);
    try {
      const [me, response] = await Promise.all([
        fetchJson<ApiUser>("/api/me"),
        fetchJson<ApiList<PlannerPlan>>("/api/planner/plans"),
      ]);
      setUser(me);
      setPlans(response.value);
    } catch (loadError) {
      setError(loadError instanceof Error ? loadError.message : "No fue posible consultar Microsoft Planner.");
    } finally {
      setLoadingPlans(false);
    }
  }, []);

  const loadBuckets = React.useCallback(async (planId: string) => {
    if (!planId) return;
    setLoadingBuckets(true);
    setError(null);
    setSelectedBucketId("");
    setBuckets([]);
    setLabels([]);
    setValidation(null);
    try {
      const [bucketResponse, labelResponse] = await Promise.all([
        fetchJson<ApiList<PlannerBucket>>(`/api/planner/plans/${planId}/buckets`),
        fetchJson<ApiList<PlannerLabel>>(`/api/planner/plans/${planId}/labels`),
      ]);
      setBuckets(bucketResponse.value);
      setLabels(labelResponse.value);
    } catch (loadError) {
      setError(loadError instanceof Error ? loadError.message : "No fue posible consultar Buckets.");
    } finally {
      setLoadingBuckets(false);
    }
  }, []);

  const validateRows = React.useCallback(
    async (rows: NormalizedTaskRow[], currentOptions: ImportOptions) => {
      setValidating(true);
      setError(null);
      try {
        const response = await fetchJson<ImportValidationResponse>("/api/import/validate", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ options: currentOptions, rows }),
        });
        setValidation(response);
      } catch (validationError) {
        setError(validationError instanceof Error ? validationError.message : "No fue posible validar el Excel.");
      } finally {
        setValidating(false);
      }
    },
    [],
  );

  React.useEffect(() => {
    void loadPlans();
  }, [loadPlans]);

  React.useEffect(() => {
    if (selectedPlanId) void loadBuckets(selectedPlanId);
  }, [loadBuckets, selectedPlanId]);

  React.useEffect(() => {
    if (fileData && options) void validateRows(fileData.rows, options);
  }, [fileData, options, validateRows]);

  const handleFile = async (file: File) => {
    setUploading(true);
    setError(null);
    setValidation(null);
    setResults([]);
    setActivity([]);
    try {
      const formData = new FormData();
      formData.append("file", file);
      const parsed = await fetchJson<ParseExcelResponse>("/api/excel/parse", {
        method: "POST",
        body: formData,
      });
      setFileData(parsed);
    } catch (uploadError) {
      setError(uploadError instanceof Error ? uploadError.message : "No fue posible leer el Excel.");
    } finally {
      setUploading(false);
    }
  };

  const importReadyCount =
    validation?.rows.filter((row) => row.status !== "invalid" && !row.shouldOmit).length ?? 0;

  const filteredRows = React.useMemo(() => {
    const rows = validation?.rows ?? [];
    if (filter === "valid") return rows.filter((row) => row.status === "valid");
    if (filter === "errors") return rows.filter((row) => row.status === "invalid");
    if (filter === "warnings") return rows.filter((row) => row.status === "warning");
    return rows;
  }, [filter, validation?.rows]);

  const columns = React.useMemo<ColumnDef<ValidatedImportRow>[]>(
    () => [
      { header: "Estado", cell: ({ row }) => statusBadge(row.original) },
      { accessorKey: "titulo", header: "Título" },
      { accessorKey: "responsableRaw", header: "Responsable" },
      { header: "Inicio", cell: ({ row }) => displayDate(row.original.fechaInicio) },
      { header: "Vencimiento", cell: ({ row }) => displayDate(row.original.fechaVencimiento) },
      { accessorKey: "etiqueta", header: "Etiqueta" },
      {
        header: "Detalle",
        cell: ({ row }) =>
          row.original.issues.length ? (
            <span className="text-xs text-slate-600" title={row.original.issues.map((entry) => entry.message).join("\n")}>
              {row.original.issues[0]?.message}
            </span>
          ) : (
            <span className="text-xs text-slate-400">Sin observaciones</span>
          ),
      },
    ],
    [],
  );

  // eslint-disable-next-line react-hooks/incompatible-library
  const table = useReactTable({
    data: filteredRows,
    columns,
    getCoreRowModel: getCoreRowModel(),
    getFilteredRowModel: getFilteredRowModel(),
  });

  const startImport = async () => {
    if (!fileData || !options) return;
    setConfirming(false);
    setImporting(true);
    setResults([]);
    setActivity([]);
    setProgress({ processed: 0, total: importReadyCount, success: 0, omitted: 0, errors: 0 });
    setError(null);

    try {
      const response = await fetch("/api/import/start", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ options, rows: fileData.rows }),
      });
      if (!response.ok || !response.body) {
        const payload = await response.json().catch(() => ({}));
        throw new Error(payload.message ?? "No fue posible iniciar la importación.");
      }

      const reader = response.body.getReader();
      const decoder = new TextDecoder();
      let pending = "";

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        pending += decoder.decode(value, { stream: true });
        const lines = pending.split("\n");
        pending = lines.pop() ?? "";

        for (const line of lines) {
          if (!line.trim()) continue;
          const event = JSON.parse(line) as ImportProgressEvent;
          if (event.type === "started") {
            setProgress((current) => ({ ...current, total: event.total }));
          }
          if (event.type === "row-started") {
            setActivity((current) => [`Creando: ${event.title}`, ...current].slice(0, 8));
          }
          if (event.type === "row-result") {
            setActivity((current) => [`${event.row.estado}: ${event.row.titulo}`, ...current].slice(0, 8));
          }
          if (event.type === "progress") {
            setProgress({
              processed: event.processed,
              total: event.total,
              success: event.successCount,
              omitted: event.omittedCount,
              errors: event.errorCount,
            });
          }
          if (event.type === "completed") {
            setResults(event.results);
            setProgress({
              processed: event.processed,
              total: event.total,
              success: event.successCount,
              omitted: event.omittedCount,
              errors: event.errorCount,
            });
          }
          if (event.type === "error") {
            throw new Error(event.message);
          }
        }
      }
    } catch (importError) {
      setError(importError instanceof Error ? importError.message : "No fue posible completar la importación.");
    } finally {
      setImporting(false);
    }
  };

  const downloadReport = async () => {
    if (!results.length) return;
    const response = await fetch("/api/import/report", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(results),
    });
    if (!response.ok) {
      const payload = await response.json().catch(() => ({}));
      setError(payload.message ?? "No fue posible descargar el reporte.");
      return;
    }
    const blob = await response.blob();
    const disposition = response.headers.get("content-disposition");
    const fileName = disposition?.match(/filename="([^"]+)"/)?.[1] ?? "planner-import-result.xlsx";
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = fileName;
    anchor.click();
    URL.revokeObjectURL(url);
  };

  const resetImport = () => {
    setFileData(null);
    setValidation(null);
    setResults([]);
    setActivity([]);
    setFilter("all");
    setError(null);
    setProgress({ processed: 0, total: 0, success: 0, omitted: 0, errors: 0 });
    setConfirming(false);
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  };

  const progressPct = progress.total ? Math.round((progress.processed / progress.total) * 100) : 0;

  return (
    <div className="min-h-screen bg-[#F7FAFC] text-[#172033]">
      <header className="sticky top-0 z-20 border-b border-slate-200 bg-white/95 backdrop-blur">
        <div className="mx-auto flex max-w-7xl items-center justify-between px-5 py-3">
          <div className="flex min-w-0 items-center gap-3">
            <div className="flex h-11 shrink-0 items-center rounded-md bg-[#0D47A1] px-3">
              <Image
                src="/gis.png"
                alt="Grupo Industrial Saltillo"
                width={121}
                height={40}
                className="h-8 w-auto object-contain"
                priority
              />
            </div>
            <div className="min-w-0">
              <p className="truncate text-sm font-semibold text-[#0D47A1]">Planner Importer</p>
              <p className="hidden text-xs text-slate-500 sm:block">Importación masiva hacia Microsoft Planner</p>
            </div>
          </div>
          <div className="flex items-center gap-3">
            <Avatar src={user?.avatarUrl} name={user?.name ?? initialUser.name} />
            <div className="hidden text-right sm:block">
              <p className="text-sm font-medium">{user?.name ?? initialUser.name}</p>
              <p className="text-xs text-slate-500">{user?.email ?? initialUser.email}</p>
            </div>
            <Button variant="ghost" size="icon" title="Cerrar sesión" onClick={() => signOut({ callbackUrl: "/login" })}>
              <LogOut className="h-4 w-4" />
            </Button>
          </div>
        </div>
      </header>

      <main className="mx-auto grid max-w-7xl gap-5 px-5 py-6 lg:grid-cols-[360px_minmax(0,1fr)]">
        <aside className="space-y-5">
          <Card>
            <CardHeader>
              <CardTitle>Paso 1 · Destino</CardTitle>
            </CardHeader>
            <CardContent className="space-y-4">
              {loadingPlans ? (
                <div className="space-y-3">
                  <Skeleton className="h-10" />
                  <Skeleton className="h-10" />
                </div>
              ) : plans.length === 0 ? (
                <div className="rounded-md border border-dashed border-slate-200 bg-slate-50 p-4 text-sm text-slate-600">
                  No se encontraron Planes de Planner para tu cuenta.
                </div>
              ) : (
                <>
                  <label className="space-y-1 text-sm font-medium">
                    <span>Plan</span>
                    <Select value={selectedPlanId} onChange={(event) => setSelectedPlanId(event.target.value)}>
                      <option value="">Selecciona un Plan</option>
                      {plans.map((plan) => (
                        <option key={plan.id} value={plan.id}>
                          {plan.title}
                        </option>
                      ))}
                    </Select>
                  </label>
                  <label className="space-y-1 text-sm font-medium">
                    <span>Bucket</span>
                    <Select
                      value={selectedBucketId}
                      disabled={!selectedPlanId || loadingBuckets}
                      onChange={(event) => setSelectedBucketId(event.target.value)}
                    >
                      <option value="">{loadingBuckets ? "Cargando Buckets..." : "Selecciona un Bucket"}</option>
                      {buckets.map((bucket) => (
                        <option key={bucket.id} value={bucket.id}>
                          {bucket.name}
                        </option>
                      ))}
                    </Select>
                  </label>
                </>
              )}
              {labels.length > 0 ? (
                <div>
                  <p className="mb-2 text-xs font-medium uppercase tracking-wide text-slate-500">Etiquetas del Plan</p>
                  <div className="flex flex-wrap gap-2">
                    {labels.map((label) => (
                      <Badge key={label.key} tone="info">{label.name}</Badge>
                    ))}
                  </div>
                </div>
              ) : null}
            </CardContent>
          </Card>

          <Card>
            <CardHeader>
              <CardTitle>Paso 2 · Archivo</CardTitle>
            </CardHeader>
            <CardContent className="space-y-4">
              <label
                className="flex cursor-pointer flex-col items-center justify-center rounded-lg border border-dashed border-slate-300 bg-white px-4 py-8 text-center transition hover:border-[#1565C0] hover:bg-[#EAF3FF]"
                onDragOver={(event) => event.preventDefault()}
                onDrop={(event) => {
                  event.preventDefault();
                  const file = event.dataTransfer.files.item(0);
                  if (file) void handleFile(file);
                }}
              >
                {uploading ? <Loader2 className="mb-3 h-9 w-9 animate-spin text-[#1565C0]" /> : <UploadCloud className="mb-3 h-9 w-9 text-[#1565C0]" />}
                <span className="text-sm font-semibold">Arrastra tu Excel aquí</span>
                <span className="mt-1 text-xs text-slate-500">XLSX · Máximo 10 MB</span>
                <span className="mt-4 inline-flex h-9 items-center rounded-md bg-[#1565C0] px-3 text-sm font-medium text-white">
                  Seleccionar archivo
                </span>
                <input
                  ref={fileInputRef}
                  type="file"
                  accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                  className="sr-only"
                  onClick={(event) => {
                    event.currentTarget.value = "";
                  }}
                  onChange={(event) => {
                    const file = event.currentTarget.files?.item(0);
                    if (file) void handleFile(file);
                    event.currentTarget.value = "";
                  }}
                />
              </label>
              {fileData ? (
                <div className="flex items-center gap-2 rounded-md border border-slate-200 bg-slate-50 px-3 py-2 text-sm">
                  <FileSpreadsheet className="h-4 w-4 text-[#1565C0]" />
                  <span className="truncate">{fileData.fileName}</span>
                  <Badge>{fileData.totalRows} filas</Badge>
                </div>
              ) : null}
              <div className="space-y-3 border-t border-slate-100 pt-4">
                <label className="flex items-start gap-2 text-sm">
                  <Checkbox checked={createMissingLabels} onChange={(event) => setCreateMissingLabels(event.currentTarget.checked)} />
                  <span>Crear automáticamente etiquetas faltantes</span>
                </label>
                <label className="flex items-start gap-2 text-sm">
                  <Checkbox checked={detectDuplicates} onChange={(event) => setDetectDuplicates(event.currentTarget.checked)} />
                  <span>Detectar posibles duplicados</span>
                </label>
                <label className="flex items-start gap-2 text-sm">
                  <Checkbox checked={omitDuplicates} onChange={(event) => setOmitDuplicates(event.currentTarget.checked)} disabled={!detectDuplicates} />
                  <span>Omitir duplicados detectados</span>
                </label>
              </div>
            </CardContent>
          </Card>
        </aside>

        <section className="space-y-5">
          {error ? (
            <div className="flex items-start gap-3 rounded-lg border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-800">
              <AlertCircle className="mt-0.5 h-4 w-4" />
              <div className="flex-1">{error}</div>
              <Button variant="ghost" size="sm" onClick={() => setError(null)}>Cerrar</Button>
            </div>
          ) : null}

          <Card>
            <CardHeader className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
              <div>
                <CardTitle>Importar tareas</CardTitle>
                <p className="mt-1 text-sm text-slate-500">
                  Todas las filas se cargarán al Plan y Bucket seleccionados en la interfaz.
                </p>
              </div>
              <Button variant="secondary" size="sm" onClick={() => void loadPlans()}>
                <RefreshCcw className="h-4 w-4" />
                Actualizar
              </Button>
            </CardHeader>
            <CardContent>
              <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
                <div className="rounded-lg border border-slate-200 bg-slate-50 p-4">
                  <p className="text-xs uppercase tracking-wide text-slate-500">Plan</p>
                  <p className="mt-1 truncate text-sm font-semibold">{selectedPlan?.title ?? "Sin seleccionar"}</p>
                </div>
                <div className="rounded-lg border border-slate-200 bg-slate-50 p-4">
                  <p className="text-xs uppercase tracking-wide text-slate-500">Bucket</p>
                  <p className="mt-1 truncate text-sm font-semibold">{selectedBucket?.name ?? "Sin seleccionar"}</p>
                </div>
                <div className="rounded-lg border border-slate-200 bg-slate-50 p-4">
                  <p className="text-xs uppercase tracking-wide text-slate-500">Archivo</p>
                  <p className="mt-1 truncate text-sm font-semibold">{fileData?.fileName ?? "Sin archivo"}</p>
                </div>
                <div className="rounded-lg border border-slate-200 bg-slate-50 p-4">
                  <p className="text-xs uppercase tracking-wide text-slate-500">Listas</p>
                  <p className="mt-1 text-sm font-semibold">{importReadyCount} tareas válidas</p>
                </div>
              </div>
            </CardContent>
          </Card>

          {validation ? (
            <Card>
              <CardHeader className="space-y-4">
                <div className="grid gap-3 sm:grid-cols-4">
                  <div className="rounded-lg border border-slate-200 p-4 text-center">
                    <p className="text-2xl font-semibold">{validation.summary.totalRows}</p>
                    <p className="text-xs text-slate-500">Filas</p>
                  </div>
                  <div className="rounded-lg border border-emerald-200 bg-emerald-50 p-4 text-center">
                    <p className="text-2xl font-semibold text-emerald-700">{validation.summary.validRows}</p>
                    <p className="text-xs text-emerald-700">Válidas</p>
                  </div>
                  <div className="rounded-lg border border-amber-200 bg-amber-50 p-4 text-center">
                    <p className="text-2xl font-semibold text-amber-700">{validation.summary.warningRows}</p>
                    <p className="text-xs text-amber-700">Advertencias</p>
                  </div>
                  <div className="rounded-lg border border-red-200 bg-red-50 p-4 text-center">
                    <p className="text-2xl font-semibold text-red-700">{validation.summary.errorRows}</p>
                    <p className="text-xs text-red-700">Con errores</p>
                  </div>
                </div>
                <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
                  <div className="flex flex-wrap gap-2">
                    {(Object.keys(filterLabels) as FilterMode[]).map((mode) => (
                      <Button
                        key={mode}
                        variant={filter === mode ? "default" : "secondary"}
                        size="sm"
                        onClick={() => setFilter(mode)}
                      >
                        {filterLabels[mode]}
                      </Button>
                    ))}
                  </div>
                  <Button disabled={!importReadyCount || validating || importing} onClick={() => setConfirming(true)}>
                    <Play className="h-4 w-4" />
                    Importar {importReadyCount} tareas
                  </Button>
                </div>
              </CardHeader>
              <CardContent>
                {validating ? (
                  <div className="space-y-3">
                    <Skeleton className="h-10" />
                    <Skeleton className="h-10" />
                    <Skeleton className="h-10" />
                  </div>
                ) : (
                  <div className="overflow-x-auto rounded-lg border border-slate-200">
                    <table className="min-w-[900px] w-full text-left text-sm">
                      <thead className="bg-slate-50 text-xs uppercase tracking-wide text-slate-500">
                        {table.getHeaderGroups().map((headerGroup) => (
                          <tr key={headerGroup.id}>
                            {headerGroup.headers.map((header) => (
                              <th key={header.id} className="px-3 py-3 font-semibold">
                                {flexRender(header.column.columnDef.header, header.getContext())}
                              </th>
                            ))}
                          </tr>
                        ))}
                      </thead>
                      <tbody className="divide-y divide-slate-100 bg-white">
                        {table.getRowModel().rows.map((row) => (
                          <tr key={row.id}>
                            {row.getVisibleCells().map((cell) => (
                              <td key={cell.id} className="max-w-[260px] px-3 py-3 align-top">
                                <div className="truncate">{flexRender(cell.column.columnDef.cell, cell.getContext())}</div>
                              </td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </CardContent>
            </Card>
          ) : (
            <Card>
              <CardContent className="flex min-h-72 flex-col items-center justify-center text-center">
                <Search className="mb-3 h-10 w-10 text-slate-300" />
                <h2 className="text-base font-semibold">Carga un Excel para previsualizar y validar</h2>
                <p className="mt-2 max-w-md text-sm text-slate-500">
                  Selecciona Plan y Bucket, luego sube un XLSX con las columnas Titulo, Responsable,
                  fecha de inicio, fecha de vencimiento, tareas y etiqueta.
                </p>
              </CardContent>
            </Card>
          )}

          {confirming ? (
            <Card className="border-[#1565C0]">
              <CardContent className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
                <div>
                  <p className="text-lg font-semibold">{importReadyCount} tareas listas para importar</p>
                  <p className="mt-1 text-sm text-slate-500">
                    Plan: {selectedPlan?.title} · Bucket: {selectedBucket?.name}
                  </p>
                </div>
                <div className="flex gap-2">
                  <Button variant="secondary" onClick={() => setConfirming(false)}>Cancelar</Button>
                  <Button onClick={() => void startImport()}>Importar {importReadyCount} tareas</Button>
                </div>
              </CardContent>
            </Card>
          ) : null}

          {importing || results.length > 0 ? (
            <Card>
              <CardHeader>
                <CardTitle>{importing ? "Importando tareas..." : "Importación completada"}</CardTitle>
              </CardHeader>
              <CardContent className="space-y-5">
                <div className="space-y-2">
                  <div className="flex items-center justify-between text-sm">
                    <span>{progress.processed} de {progress.total} tareas procesadas</span>
                    <span className="font-semibold">{progressPct}%</span>
                  </div>
                  <Progress value={progressPct} />
                </div>
                <div className="grid gap-3 sm:grid-cols-3">
                  <Badge tone="success">{progress.success} creadas correctamente</Badge>
                  <Badge tone="warning">{progress.omitted} omitidas</Badge>
                  <Badge tone="error">{progress.errors} errores durante la creación</Badge>
                </div>
                {activity.length ? (
                  <div className="rounded-lg border border-slate-200 bg-slate-50 p-3 text-sm text-slate-700">
                    {activity.map((entry, index) => (
                      <p key={`${entry}-${index}`} className="py-1">{entry}</p>
                    ))}
                  </div>
                ) : null}
                {results.length ? (
                  <>
                    <div className="overflow-x-auto rounded-lg border border-slate-200">
                      <table className="min-w-[760px] w-full text-left text-sm">
                        <thead className="bg-slate-50 text-xs uppercase tracking-wide text-slate-500">
                          <tr>
                            <th className="px-3 py-3">Resultado</th>
                            <th className="px-3 py-3">Tarea</th>
                            <th className="px-3 py-3">Detalle</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {results.map((row) => (
                            <tr key={`${row.fila}-${row.titulo}`}>
                              <td className="px-3 py-3"><Badge tone={resultTone(row)}>{row.estado}</Badge></td>
                              <td className="px-3 py-3">{row.titulo}</td>
                              <td className="px-3 py-3 text-slate-600">{row.mensaje}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                    <div className="flex flex-wrap gap-2">
                      <Button variant="secondary" onClick={resetImport}>Nueva importación</Button>
                      <Button onClick={() => void downloadReport()}>
                        <ArrowDownToLine className="h-4 w-4" />
                        Descargar reporte
                      </Button>
                    </div>
                  </>
                ) : null}
              </CardContent>
            </Card>
          ) : null}
        </section>
      </main>
    </div>
  );
}
