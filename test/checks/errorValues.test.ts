import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { parseWorkbook } from "../../src/parse";
import { errorValues } from "../../src/checks/errorValues";
const ctx = (n: string) => parseWorkbook(new Uint8Array(readFileSync(join(__dirname, "..", "fixtures", n))));

describe("errorValues", () => {
  it("counts error cells and lists addresses", () => {
    const f = errorValues(ctx("errors.xlsx"));
    expect(f?.id).toBe("broken-today.error-values");
    expect(f?.count).toBe(3);
    expect(f?.locations.some(l => l.includes("Calc!"))).toBe(true);
    expect(f?.severity).toBe("high");
  });
  it("does not fire on a clean file", () => {
    expect(errorValues(ctx("clean.xlsx"))).toBeNull();
  });
});
