import type { Check } from "../types";
import { SUBSTANTIAL_FORMULA_COUNT } from "../types";
import { forEachCell } from "../cells";

const ENOUGH_INPUTS = 10;

// Only meaningful for OOXML: legacy .xls is not unzipped, so we cannot see its validation
// rules and must not claim they are absent. Gated, too, on the workbook actually having
// logic and typed-in inputs - "no validation" on a four-cell sheet is noise, not a finding.
export const missingValidation: Check = (ctx) => {
  if (ctx.fileFormat !== "xlsx" && ctx.fileFormat !== "xlsm") return null;
  if (ctx.formulaCells.length < SUBSTANTIAL_FORMULA_COUNT) return null;

  const hasValidation = Object.values(ctx.sheetXml).some(xml => /<dataValidation\b/.test(xml));
  if (hasValidation) return null;

  let inputs = 0;
  forEachCell(ctx, (_sheet, _addr, cell) => {
    if (!cell.f && cell.t === "n") inputs++;
  });
  if (inputs < ENOUGH_INPUTS) return null;

  return {
    id: "hidden-fragility.missing-validation",
    category: "hidden-fragility",
    severity: "low",
    title: "Nothing stops a bad value being typed in",
    soWhat: `Not one of the ${inputs} numbers typed into this workbook is protected by a validation rule, so a date in a price cell, a text entry in a quantity, or a misplaced decimal point is accepted silently and flows straight through the calculations.`,
    action: "Put data validation on the input cells - a number range, a date range, a dropdown list - so the workbook rejects the mistake at the point it is made.",
    locations: [],
    count: 1,
  };
};
