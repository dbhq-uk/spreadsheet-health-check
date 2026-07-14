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
  it("does not flag a consistent column that uses a function name (LOG10)", () => {
    expect(inconsistentFormulas(ctx("functions.xlsx"))).toBeNull();
  });
});
