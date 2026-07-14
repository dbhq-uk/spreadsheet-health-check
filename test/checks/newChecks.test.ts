import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { parseWorkbook } from "../../src/parse";
import { manualCalc } from "../../src/checks/manualCalc";
import { hardcodedConstants } from "../../src/checks/hardcodedConstants";
import { duplicateSheets } from "../../src/checks/duplicateSheets";
import { mergedCells } from "../../src/checks/mergedCells";
import { missingValidation } from "../../src/checks/missingValidation";
import type { ParsedContext } from "../../src/types";

const ctx = (n: string) => parseWorkbook(new Uint8Array(readFileSync(join(__dirname, "..", "fixtures", n))));

describe("manualCalc", () => {
  it("fires when automatic calculation is switched off", () => {
    const f = manualCalc(ctx("manual-calc.xlsx"));
    expect(f?.id).toBe("broken-today.manual-calc");
    expect(f?.severity).toBe("high");
  });
  it("does not fire on a normal workbook", () => {
    expect(manualCalc(ctx("clean.xlsx"))).toBeNull();
  });
  // This check is only worth its alarming wording if it has no false positives: fullCalcOnLoad
  // forces a recalculation when the file opens, so the numbers on screen are not stale.
  it("does not fire when the workbook fully recalculates on load", () => {
    const c = ctx("manual-calc.xlsx");
    const patched: ParsedContext = {
      ...c,
      workbookXml: c.workbookXml.replace("<calcPr", '<calcPr fullCalcOnLoad="1" '),
    };
    expect(manualCalc(patched)).toBeNull();
  });
});

describe("hardcodedConstants", () => {
  it("fires on a rate typed into formula after formula", () => {
    const f = hardcodedConstants(ctx("hardcoded.xlsx"));
    expect(f?.id).toBe("hidden-fragility.hardcoded-constants");
    expect(f?.locations[0]).toMatch(/^1\.2 appears in 25 formulas/);
  });
  // The VLOOKUP column index in every one of those same formulas is a 2. Reporting it would be
  // a false positive on a perfectly normal lookup - exactly the kind that costs you the report.
  it("does not report the lookup column index as a buried business rule", () => {
    const f = hardcodedConstants(ctx("hardcoded.xlsx"));
    const reported = f!.locations.map(l => l.split(" ")[0]);
    expect(reported).toEqual(["1.2"]);
  });
  it("does not fire on formulas that reference cells instead of burying numbers", () => {
    expect(hardcodedConstants(ctx("clean.xlsx"))).toBeNull();
    expect(hardcodedConstants(ctx("documented.xlsx"))).toBeNull();
  });
});

describe("duplicateSheets", () => {
  it("fires on the same sheet copied per month", () => {
    const f = duplicateSheets(ctx("duplicate-sheets.xlsx"));
    expect(f?.id).toBe("wont-scale.duplicate-sheets");
    expect(f?.count).toBe(4);
    expect(f?.locations[0]).toBe("Jan, Feb, Mar, Apr");
  });
  it("does not fire on sheets that genuinely differ", () => {
    expect(duplicateSheets(ctx("hidden.xlsx"))).toBeNull();
    expect(duplicateSheets(ctx("clean.xlsx"))).toBeNull();
  });
});

describe("mergedCells", () => {
  it("fires on merges inside the data region", () => {
    const f = mergedCells(ctx("merged.xlsx"));
    expect(f?.id).toBe("hidden-fragility.merged-cells");
    expect(f?.count).toBe(2);
  });
  // A merged title across the top of a sheet is formatting, not a fault. Counting it would fire
  // on a large share of ordinary workbooks.
  it("ignores a merged title banner across the top", () => {
    expect(mergedCells(ctx("merged.xlsx"))?.locations.join(" ")).not.toMatch(/A1:C1/);
  });
  it("does not fire on a workbook with no merges", () => {
    expect(mergedCells(ctx("clean.xlsx"))).toBeNull();
  });
});

describe("missingValidation", () => {
  it("fires when a workbook full of typed-in numbers validates none of them", () => {
    expect(missingValidation(ctx("hardcoded.xlsx"))?.id).toBe("hidden-fragility.missing-validation");
  });
  it("does not fire when the inputs are validated", () => {
    expect(missingValidation(ctx("documented.xlsx"))).toBeNull();
  });
  it("does not fire on a workbook too small to have real inputs", () => {
    expect(missingValidation(ctx("clean.xlsx"))).toBeNull();
  });
  // Legacy .xls is not unzipped, so its validation rules are invisible to us. Claiming they are
  // absent would be a guess dressed up as a finding.
  it("stays silent on legacy .xls rather than guessing", () => {
    expect(missingValidation(ctx("legacy.xls"))).toBeNull();
  });
});
