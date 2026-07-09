import { randomUUID } from "node:crypto";

import { appConfig } from "@/lib/config";
import { GraphApiError } from "@/lib/graph/graph-errors";

type GraphFetchOptions = {
  path: string;
  token: string;
  method?: "GET" | "POST" | "PATCH" | "DELETE";
  body?: unknown;
  headers?: HeadersInit;
  maxRetries?: number;
  timeoutMs?: number;
  retryConflict?: boolean;
};

type GraphErrorPayload = {
  error?: {
    code?: string;
    message?: string;
    innerError?: {
      "request-id"?: string;
      requestId?: string;
      date?: string;
    };
  };
};

function sleep(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function retryAfterMs(headers: Headers) {
  const retryAfter = headers.get("retry-after");
  if (!retryAfter) return undefined;
  const numeric = Number(retryAfter);
  if (Number.isFinite(numeric)) return numeric * 1000;
  const date = Date.parse(retryAfter);
  return Number.isFinite(date) ? Math.max(0, date - Date.now()) : undefined;
}

function shouldRetry(status: number, retryConflict?: boolean) {
  return status === 429 || status === 500 || status === 502 || status === 503 || (retryConflict && (status === 409 || status === 412));
}

function graphUrl(path: string) {
  if (path.startsWith("https://")) return path;
  return `${appConfig.graphBaseUrl}${path.startsWith("/") ? path : `/${path}`}`;
}

async function parseGraphError(response: Response) {
  let payload: GraphErrorPayload | undefined;
  try {
    payload = (await response.json()) as GraphErrorPayload;
  } catch {
    payload = undefined;
  }

  const requestId =
    response.headers.get("request-id") ??
    response.headers.get("client-request-id") ??
    payload?.error?.innerError?.["request-id"] ??
    payload?.error?.innerError?.requestId ??
    undefined;

  return new GraphApiError({
    status: response.status,
    code: payload?.error?.code,
    message: payload?.error?.message ?? `Microsoft Graph returned HTTP ${response.status}`,
    requestId,
    retryAfter: retryAfterMs(response.headers),
    details: payload,
  });
}

export async function graphFetch<T>({
  path,
  token,
  method = "GET",
  body,
  headers,
  maxRetries = appConfig.importMaxRetries,
  timeoutMs = 30_000,
  retryConflict = false,
}: GraphFetchOptions): Promise<T> {
  let attempt = 0;
  let lastError: unknown;

  while (attempt <= maxRetries) {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), timeoutMs);
    const clientRequestId = randomUUID();

    try {
      const response = await fetch(graphUrl(path), {
        method,
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: "application/json",
          "client-request-id": clientRequestId,
          ...(body ? { "Content-Type": "application/json" } : {}),
          ...headers,
        },
        body: body ? JSON.stringify(body) : undefined,
        signal: controller.signal,
      });

      clearTimeout(timeout);

      if (response.ok) {
        if (response.status === 204) return undefined as T;
        return (await response.json()) as T;
      }

      const graphError = await parseGraphError(response);
      lastError = graphError;

      if (!shouldRetry(response.status, retryConflict) || attempt === maxRetries) {
        throw graphError;
      }

      const retryMs =
        graphError.retryAfter ??
        Math.min(20_000, 500 * 2 ** attempt + Math.floor(Math.random() * 300));
      await sleep(retryMs);
    } catch (error) {
      clearTimeout(timeout);
      lastError = error;

      if (error instanceof GraphApiError) {
        if (!shouldRetry(error.status, retryConflict) || attempt === maxRetries) throw error;
      } else if (attempt === maxRetries) {
        throw error;
      }

      await sleep(Math.min(20_000, 500 * 2 ** attempt + Math.floor(Math.random() * 300)));
    }

    attempt += 1;
  }

  throw lastError instanceof Error ? lastError : new Error("Microsoft Graph request failed");
}

export async function graphFetchCollection<T>(path: string, token: string) {
  const items: T[] = [];
  let nextPath: string | undefined = path;

  while (nextPath) {
    const page: { value: T[]; "@odata.nextLink"?: string } = await graphFetch({
      path: nextPath,
      token,
    });
    items.push(...page.value);
    nextPath = page["@odata.nextLink"];
  }

  return items;
}
