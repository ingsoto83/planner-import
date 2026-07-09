import { describe, expect, it } from "vitest";

import { labelMapFromDescriptions, labelsFromDescriptions } from "../planner-service";

describe("Planner label mapping", () => {
  it("normalizes category descriptions into a lookup map", () => {
    const descriptions = {
      category1: "Desarrollo",
      category2: "Infraestructura",
      category3: "",
    };

    expect(labelsFromDescriptions(descriptions)).toEqual([
      { key: "category1", name: "Desarrollo", normalizedName: "desarrollo" },
      { key: "category2", name: "Infraestructura", normalizedName: "infraestructura" },
    ]);
    expect(Object.fromEntries(labelMapFromDescriptions(descriptions))).toEqual({
      desarrollo: "category1",
      infraestructura: "category2",
    });
  });
});
