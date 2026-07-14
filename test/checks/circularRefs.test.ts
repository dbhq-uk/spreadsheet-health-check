import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { parseWorkbook } from "../../src/parse";
import { circularRefs } from "../../src/checks/circularRefs";
const ctx = (n: string) => parseWorkbook(new Uint8Array(readFileSync(join(__dirname, "..", "fixtures", n))));

describe("circularRefs", () => {
  it("fires when iterative calc is enabled", () => {
    expect(circularRefs(ctx("circular.xlsx"))?.id).toBe("broken-today.circular-refs");
  });
  it("does not fire on a clean file", () => {
    expect(circularRefs(ctx("clean.xlsx"))).toBeNull();
  });
});
