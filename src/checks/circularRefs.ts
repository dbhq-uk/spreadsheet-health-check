import type { Check } from "../types";

export const circularRefs: Check = (ctx) => {
  const m = ctx.workbookXml.match(/<calcPr[^>]*\biterate="(1|true)"/i);
  if (!m) return null;
  return {
    id: "broken-today.circular-refs",
    category: "broken-today",
    severity: "medium",
    title: "The workbook relies on circular references",
    soWhat: "Iterative calculation has been switched on to stop a formula loop from erroring - a fragile setup where a small change can silently shift every result.",
    locations: [],
    count: 1,
  };
};
