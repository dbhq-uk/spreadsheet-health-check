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

  // A file that is not a spreadsheet must be refused, not described. SheetJS is permissive
  // enough to hand back an empty workbook for arbitrary bytes, and the old format sniff then
  // read "no PK header" as "legacy .xls" - so a renamed JPEG scored a confident Moderate,
  // "near the limits of what a spreadsheet can hold". A report about a file we never parsed is
  // worse than no report: it is the finding a sceptical owner uses to dismiss the whole tool.
  it("refuses bytes that are not a spreadsheet at all", () => {
    expect(() => parseWorkbook(new Uint8Array([1, 2, 3, 4]))).toThrow(/not a spreadsheet/i);
    expect(() => parseWorkbook(new Uint8Array())).toThrow(/not a spreadsheet/i);
    // a JPEG, renamed to .xls by a hopeful user
    expect(() => parseWorkbook(new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10]))).toThrow(/not a spreadsheet/i);
  });

  it("refuses a zip that is not a workbook", () => {
    // a .zip renamed .xlsx: unzips cleanly, but there is no workbook inside it
    const notAWorkbook = new Uint8Array([0x50, 0x4b, 0x03, 0x04, ...new Array(26).fill(0)]);
    expect(() => parseWorkbook(notAWorkbook)).toThrow(/not a spreadsheet/i);
  });
});
