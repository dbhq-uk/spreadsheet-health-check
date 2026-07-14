import { describe, it, expect } from "vitest";
import type { Finding, HealthReport } from "../src/types";

describe("types", () => {
  it("Finding and HealthReport shapes are usable", () => {
    const f: Finding = {
      id: "x.y", category: "key-person", severity: "medium",
      title: "t", soWhat: "s", locations: [], count: 1,
    };
    const r: HealthReport = {
      riskBand: "Low", score: 0, verdict: "v",
      categories: [], findings: [f],
      meta: { sheetCount: 1, fileFormat: "xlsx", author: null, lastModifiedBy: null, fileSizeBytes: 10 },
    };
    expect(r.findings[0].id).toBe("x.y");
  });
});
