import type { Check } from "../types";

// Excel recalculates automatically unless someone turns it off. With calculation set to
// manual, every formula on screen shows the result of the last time somebody pressed F9 -
// the inputs may have changed a dozen times since. The numbers can be confidently, silently,
// completely wrong, and nothing on the screen says so.
//
// fullCalcOnLoad forces a recalculation when the file opens, which removes the staleness, so
// a workbook carrying it is not flagged. That precision is the point: this check is only
// worth its alarming wording if it has no false positives.
export const manualCalc: Check = (ctx) => {
  const calcPr = /<calcPr\b[^>]*>/.exec(ctx.workbookXml)?.[0];
  if (!calcPr) return null;
  if (!/\bcalcMode="manual"/i.test(calcPr)) return null;
  if (/\bfullCalcOnLoad="(1|true)"/i.test(calcPr)) return null;

  return {
    id: "broken-today.manual-calc",
    category: "broken-today",
    severity: "high",
    title: "Automatic calculation is switched off",
    soWhat: "Formulas only recalculate when someone presses F9, so the numbers on screen can be out of date and still look completely normal - a figure could go into a board pack, an invoice or a return long after the data behind it changed.",
    action: "Turn calculation back to automatic (Formulas > Calculation Options > Automatic) and check every number that has been reported out of this workbook recently.",
    locations: [],
    count: 1,
  };
};
