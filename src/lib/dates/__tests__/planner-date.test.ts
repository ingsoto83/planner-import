import { describe, expect, it } from "vitest";

import { graphDateTimeFromDateOnly, parsePlannerDate } from "../planner-date";

describe("planner date handling", () => {
  it("keeps DD/MM/YYYY dates in America/Monterrey without shifting day", () => {
    const parsed = parsePlannerDate("15/07/2026");

    expect(parsed?.date).toBe("2026-07-15");
    expect(parsed?.display).toBe("15/07/2026");
    expect(graphDateTimeFromDateOnly(parsed?.date ?? null)).toContain("T");
  });

  it("parses Excel serial dates", () => {
    expect(parsePlannerDate(46218)?.date).toBe("2026-07-15");
  });

  it("returns null for invalid dates", () => {
    expect(parsePlannerDate("99/99/2026")).toBeNull();
  });
});
