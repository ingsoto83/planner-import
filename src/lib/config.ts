const DEFAULT_TIMEZONE = "America/Monterrey";

function intEnv(name: string, fallback: number) {
  const raw = process.env[name];
  if (!raw) return fallback;
  const parsed = Number.parseInt(raw, 10);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
}

export const appConfig = {
  appName: process.env.NEXT_PUBLIC_APP_NAME ?? "Planner Importer",
  timezone: process.env.APP_TIMEZONE ?? DEFAULT_TIMEZONE,
  maxUploadSizeMb: intEnv("MAX_UPLOAD_SIZE_MB", 10),
  maxImportRows: intEnv("MAX_IMPORT_ROWS", 2000),
  importConcurrency: intEnv("IMPORT_CONCURRENCY", 5),
  importMaxRetries: intEnv("IMPORT_MAX_RETRIES", 4),
  graphBaseUrl: "https://graph.microsoft.com/v1.0",
};

export const GRAPH_SCOPES = [
  "openid",
  "profile",
  "offline_access",
  "User.Read",
  "User.ReadBasic.All",
  "Tasks.ReadWrite",
].join(" ");
