import type { Check } from "../types";
import { forEachCell } from "../cells";

export const errorValues: Check = (ctx) => {
  const locations: string[] = [];
  forEachCell(ctx, (sheet, addr, cell) => {
    if (cell.t === "e") locations.push(`${sheet}!${addr}`);
  });
  if (locations.length === 0) return null;
  return {
    id: "broken-today.error-values",
    category: "broken-today",
    severity: "high",
    title: locations.length === 1 ? "A cell is showing an error" : `${locations.length} cells are showing errors`,
    soWhat: "These calculations are broken right now - the wrong number can reach a customer, a board pack or a return before anyone notices.",
    action: "Fix these cells before the next number goes out of this workbook, and check anything already reported from it - a total that includes an error cell is wrong, not just untidy.",
    locations: locations.slice(0, 20),
    count: locations.length,
  };
};
