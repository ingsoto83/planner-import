export class GraphApiError extends Error {
  status: number;
  code?: string;
  requestId?: string;
  retryAfter?: number;
  details?: unknown;

  constructor({
    message,
    status,
    code,
    requestId,
    retryAfter,
    details,
  }: {
    message: string;
    status: number;
    code?: string;
    requestId?: string;
    retryAfter?: number;
    details?: unknown;
  }) {
    super(message);
    this.name = "GraphApiError";
    this.status = status;
    this.code = code;
    this.requestId = requestId;
    this.retryAfter = retryAfter;
    this.details = details;
  }
}

export function isGraphApiError(error: unknown): error is GraphApiError {
  return error instanceof GraphApiError;
}

export function isPlannerPremiumOrUnsupported(error: unknown) {
  if (!isGraphApiError(error)) return false;
  const text = `${error.code ?? ""} ${error.message}`.toLowerCase();
  return (
    error.status === 403 &&
    (text.includes("premium") ||
      text.includes("project") ||
      text.includes("not supported") ||
      text.includes("unsupported") ||
      text.includes("notallowed"))
  );
}
