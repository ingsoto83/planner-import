import { DateTime } from "luxon";

import { appConfig } from "@/lib/config";

export type ParsedPlannerDate = {
  date: string;
  graphDateTime: string;
  display: string;
};

const DATE_FORMATS = ["d/M/yyyy", "dd/MM/yyyy", "d-M-yyyy", "dd-MM-yyyy", "yyyy-MM-dd"];

function fromDateParts(year: number, month: number, day: number) {
  return DateTime.fromObject({ year, month, day, hour: 12, minute: 0, second: 0 }, { zone: appConfig.timezone });
}

function fromExcelSerial(value: number) {
  if (!Number.isFinite(value) || value <= 0) return null;
  const wholeDays = Math.floor(value);
  const fraction = value - wholeDays;
  return DateTime.fromObject({ year: 1899, month: 12, day: 30 }, { zone: appConfig.timezone })
    .plus({ days: wholeDays, milliseconds: Math.round(fraction * 86_400_000) })
    .set({ hour: 12, minute: 0, second: 0, millisecond: 0 });
}

function normalizeDateTime(dateTime: DateTime) {
  if (!dateTime.isValid) return null;
  const local = fromDateParts(dateTime.year, dateTime.month, dateTime.day);
  return {
    date: local.toISODate() ?? "",
    graphDateTime: local.toUTC().toISO({ suppressMilliseconds: true }) ?? "",
    display: local.toFormat("dd/MM/yyyy"),
  };
}

export function parsePlannerDate(value: unknown): ParsedPlannerDate | null {
  if (value == null || value === "") return null;

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return normalizeDateTime(fromDateParts(value.getUTCFullYear(), value.getUTCMonth() + 1, value.getUTCDate()));
  }

  if (typeof value === "number") {
    return normalizeDateTime(fromExcelSerial(value) ?? DateTime.invalid("invalid serial"));
  }

  const raw = String(value).trim();
  if (!raw) return null;

  const numeric = Number(raw);
  if (Number.isFinite(numeric) && raw.length <= 8) {
    return normalizeDateTime(fromExcelSerial(numeric) ?? DateTime.invalid("invalid serial"));
  }

  const iso = DateTime.fromISO(raw, { zone: appConfig.timezone });
  if (iso.isValid) return normalizeDateTime(iso);

  for (const format of DATE_FORMATS) {
    const parsed = DateTime.fromFormat(raw, format, { zone: appConfig.timezone, locale: "es-MX" });
    if (parsed.isValid) return normalizeDateTime(parsed);
  }

  return null;
}

export function compareDateOnly(a: string | null, b: string | null) {
  if (!a || !b) return 0;
  return a.localeCompare(b);
}

export function graphDateTimeFromDateOnly(date: string | null) {
  if (!date) return undefined;
  const parsed = DateTime.fromISO(date, { zone: appConfig.timezone });
  if (!parsed.isValid) return undefined;
  return normalizeDateTime(parsed)?.graphDateTime;
}
