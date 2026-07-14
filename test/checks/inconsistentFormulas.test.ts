import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { parseWorkbook } from "../../src/parse";
import { inconsistentFormulas } from "../../src/checks/inconsistentFormulas";
const ctx = (n: string) => parseWorkbook(new Uint8Array(readFileSync(join(__dirname, "..", "fixtures", n))));

describe("inconsistentFormulas", () => {
  it("flags the outlier formula in a consistent column", () => {
    const f = inconsistentFormulas(ctx("inconsistent.xlsx"));
    expect(f?.id).toBe("hidden-fragility.inconsistent-formulas");
    expect(f?.count).toBe(1);
    expect(f?.locations[0]).toContain("Calc!B4"); // the *3 outlier
  });
  it("does not fire when all formulas are consistent", () => {
    expect(inconsistentFormulas(ctx("clean.xlsx"))).toBeNull();
  });
  it("does not flag an aggregate totals row at the foot of a column, but still flags a real trailing outlier", () => {
    const f = inconsistentFormulas(ctx("totals-row.xlsx"));
    // B6 is =SUM(B2:B5) under a column of =A*2 - a deliberate total, not a mistake.
    expect(f?.locations.some(l => l.includes("Calc!B6"))).toBe(false);
    // C6 is =A6*3 under a column of =A+1 - the classic wrong-cell copy, still flagged.
    expect(f?.locations).toContain("Calc!C6");
  });
  it("does not flag a consistent column that uses a function name (LOG10)", () => {
    expect(inconsistentFormulas(ctx("functions.xlsx"))).toBeNull();
  });
});
