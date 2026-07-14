import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { parseWorkbook } from "../../src/parse";
import { macros } from "../../src/checks/macros";
const ctx = (n: string) => parseWorkbook(new Uint8Array(readFileSync(join(__dirname, "..", "fixtures", n))));

describe("macros", () => {
  it("fires on an xlsm with vbaProject.bin", () => {
    expect(macros(ctx("macros.xlsm"))?.id).toBe("key-person.macros");
  });
  it("does not fire on a plain xlsx", () => {
    expect(macros(ctx("clean.xlsx"))).toBeNull();
  });
});
