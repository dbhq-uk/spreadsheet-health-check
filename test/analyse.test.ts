import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { analyse } from "../src/index";
const bytes = (n: string) => new Uint8Array(readFileSync(join(__dirname, "fixtures", n)));

describe("analyse", () => {
  it("a clean workbook scores Low with no findings", () => {
    const r = analyse(bytes("clean.xlsx"));
    expect(r.riskBand).toBe("Low");
    expect(r.findings.length).toBe(0);
    expect(r.meta.fileFormat).toBe("xlsx");
  });
  it("a macro-enabled, single-author file raises key-person findings", () => {
    const r = analyse(bytes("macros.xlsm"));
    const ids = r.findings.map(f => f.id);
    expect(ids).toContain("key-person.macros");
    expect(r.verdict).toMatch(/risk/i);
  });
  it("the errors file reports broken-today error values", () => {
    const r = analyse(bytes("errors.xlsx"));
    expect(r.findings.map(f => f.id)).toContain("broken-today.error-values");
    expect(["Moderate","High","Critical"]).toContain(r.riskBand);
  });
});
