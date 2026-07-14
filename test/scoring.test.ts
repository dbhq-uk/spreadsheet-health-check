import { describe, it, expect } from "vitest";
import { computeScore } from "../src/scoring";
import type { Finding } from "../src/types";
const f = (o: Partial<Finding>): Finding => ({ id: "x", category: "broken-today", severity: "high", title: "", soWhat: "", locations: [], count: 1, ...o });

describe("computeScore", () => {
  it("no findings => Low band, score 0", () => {
    const r = computeScore([]);
    expect(r.band).toBe("Low");
    expect(r.score).toBe(0);
  });
  it("a high-severity finding raises the band above Low", () => {
    const r = computeScore([f({ severity: "high" })]);
    expect(r.score).toBeGreaterThan(0);
    expect(["Moderate", "High", "Critical"]).toContain(r.band);
  });
  it("groups findings into their categories", () => {
    const r = computeScore([f({ category: "key-person" }), f({ category: "broken-today" })]);
    const cats = Object.fromEntries(r.categories.map(c => [c.category, c.findings.length]));
    expect(cats["key-person"]).toBe(1);
    expect(cats["broken-today"]).toBe(1);
  });
});
