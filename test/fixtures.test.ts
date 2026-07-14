// test/fixtures.test.ts
import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join } from "node:path";
const f = (n: string) => readFileSync(join(__dirname, "fixtures", n));

describe("fixtures", () => {
  it("all fixtures exist and are non-empty", () => {
    for (const n of ["clean.xlsx","errors.xlsx","hidden.xlsx","inconsistent.xlsx","circular.xlsx","external-links.xlsx","macros.xlsm","legacy.xls","key-person.xlsx","manual-calc.xlsx","hardcoded.xlsx","duplicate-sheets.xlsx","merged.xlsx","documented.xlsx"]) {
      expect(f(n).byteLength).toBeGreaterThan(0);
    }
  });
});
