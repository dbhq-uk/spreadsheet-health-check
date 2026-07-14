import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { parseWorkbook } from "../src/parse";
const bytes = (n: string) => new Uint8Array(readFileSync(join(__dirname, "fixtures", n)));

describe("parseWorkbook", () => {
  it("reads sheets, format and props from a clean xlsx", () => {
    const ctx = parseWorkbook(bytes("clean.xlsx"));
    expect(ctx.fileFormat).toBe("xlsx");
    expect(ctx.sheets.map(s => s.name)).toContain("Sales");
    expect(ctx.sheets[0].visibility).toBe("visible");
    expect(ctx.props.author).toBe("Alice");
    expect(ctx.workbookXml).toContain("<workbook");
  });
  it("detects hidden and veryHidden visibility", () => {
    const ctx = parseWorkbook(bytes("hidden.xlsx"));
    const vis = Object.fromEntries(ctx.sheets.map(s => [s.name, s.visibility]));
    expect(vis["Old"]).toBe("hidden");
    expect(vis["Working"]).toBe("veryHidden");
  });
  it("classifies legacy xls and xlsm", () => {
    expect(parseWorkbook(bytes("legacy.xls")).fileFormat).toBe("xls");
    expect(parseWorkbook(bytes("macros.xlsm")).fileFormat).toBe("xlsm");
  });
});
