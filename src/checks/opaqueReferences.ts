import type { Check } from "../types";
import { functionsUsed } from "../formula";

// INDIRECT and OFFSET build a reference at calculation time from text or an offset, so the
// cell a formula actually reads cannot be established by reading the formula. Excel's own
// trace-precedents cannot follow them either. Whoever wrote them is the only person who
// knows where the numbers come from.
const OPAQUE = new Set(["INDIRECT", "OFFSET"]);

export const opaqueReferences: Check = (ctx) => {
  const locations: string[] = [];
  for (const { sheet, addr, f } of ctx.formulaCells) {
    if (functionsUsed(f).some(fn => OPAQUE.has(fn))) locations.push(`${sheet}!${addr}`);
  }
  if (locations.length === 0) return null;
  return {
    id: "key-person.opaque-references",
    category: "key-person",
    severity: locations.length >= 10 ? "high" : "medium",
    title: locations.length === 1
      ? "A formula builds its own reference (INDIRECT/OFFSET)"
      : `${locations.length} formulas build their own references (INDIRECT/OFFSET)`,
    soWhat: "These work out which cell to read while they calculate, so you cannot tell what feeds them by reading them - not even Excel's own formula tracing can follow them.",
    action: "Ask whoever wrote them to replace each one with a direct reference or a named range, so the workbook can be read and checked by someone else.",
    locations: locations.slice(0, 20),
    count: locations.length,
  };
};
