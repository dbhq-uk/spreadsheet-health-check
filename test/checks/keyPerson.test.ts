// The rebuilt key-person signal: evidence that only the author can safely change the
// workbook. Never metadata - see src/checks/index.ts for why single-author was deleted.
import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { parseWorkbook } from "../../src/parse";
import { opaqueReferences } from "../../src/checks/opaqueReferences";
import { formulaComplexity } from "../../src/checks/formulaComplexity";
import { formulaSprawl } from "../../src/checks/formulaSprawl";
import { protectedSheets } from "../../src/checks/protectedSheets";
import { undocumented } from "../../src/checks/undocumented";

const ctx = (n: string) => parseWorkbook(new Uint8Array(readFileSync(join(__dirname, "..", "fixtures", n))));

describe("opaqueReferences", () => {
  it("fires on INDIRECT/OFFSET - references that cannot be traced by reading them", () => {
    const f = opaqueReferences(ctx("key-person.xlsx"));
    expect(f?.id).toBe("key-person.opaque-references");
    expect(f?.count).toBe(6);
  });
  it("does not fire on a workbook of plain references", () => {
    expect(opaqueReferences(ctx("clean.xlsx"))).toBeNull();
    expect(opaqueReferences(ctx("documented.xlsx"))).toBeNull();
  });
});

describe("formulaComplexity", () => {
  it("fires on formulas too deep or too long to check by reading", () => {
    const f = formulaComplexity(ctx("key-person.xlsx"));
    expect(f?.id).toBe("key-person.formula-complexity");
    expect(f?.locations[0]).toMatch(/deep|characters|lookups/);
  });
  it("does not fire on ordinary formulas", () => {
    expect(formulaComplexity(ctx("clean.xlsx"))).toBeNull();
    expect(formulaComplexity(ctx("documented.xlsx"))).toBeNull();
  });
});

describe("formulaSprawl", () => {
  it("fires when the formulas are one-offs rather than a repeated pattern", () => {
    const f = formulaSprawl(ctx("key-person.xlsx"));
    expect(f?.id).toBe("key-person.formula-sprawl");
  });
  // documented.xlsx is 40 formulas of a single shape - the pattern a good workbook has.
  it("does not fire on a large workbook built on one consistent pattern", () => {
    expect(formulaSprawl(ctx("documented.xlsx"))).toBeNull();
  });
  it("does not fire on a workbook too small to have a pattern at all", () => {
    expect(formulaSprawl(ctx("clean.xlsx"))).toBeNull();
  });
});

describe("protectedSheets", () => {
  it("fires on a sheet locked against editing", () => {
    expect(protectedSheets(ctx("key-person.xlsx"))?.id).toBe("key-person.protected-sheets");
  });
  it("is scored low, because locking is often just good practice", () => {
    expect(protectedSheets(ctx("key-person.xlsx"))?.severity).toBe("low");
  });
  it("does not fire on an unlocked workbook", () => {
    expect(protectedSheets(ctx("clean.xlsx"))).toBeNull();
  });
});

describe("undocumented", () => {
  it("fires when a workbook with real logic explains none of it", () => {
    expect(undocumented(ctx("hardcoded.xlsx"))?.id).toBe("key-person.undocumented");
  });
  it("does not fire when there is a notes sheet and a cell comment", () => {
    expect(undocumented(ctx("documented.xlsx"))).toBeNull();
  });
  // The single-author lesson: an absence-of-X check must not fire on a workbook that never
  // had anything to document. A four-formula sheet needs no notes sheet.
  it("does not fire on a workbook too small to be worth documenting", () => {
    expect(undocumented(ctx("clean.xlsx"))).toBeNull();
  });
});
