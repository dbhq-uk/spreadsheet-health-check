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
    action: "Find the loop and break it. Iterative calculation hides the problem rather than solving it, and the answers it settles on depend on where it started.",
    locations: [],
    count: 1,
  };
};
