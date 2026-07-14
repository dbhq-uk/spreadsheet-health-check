import type { Check } from "../types";
import { numericLiterals, isBusinessConstant } from "../formula";

const REPEATED = 3;   // the same number written into this many formulas is a rule, not a one-off
const PERVASIVE = 20; // ...and at this many it is a rule the business cannot change safely

export const hardcodedConstants: Check = (ctx) => {
  // value -> the cells that bury it
  const byValue = new Map<number, string[]>();
  for (const { sheet, addr, f } of ctx.formulaCells) {
    for (const n of new Set(numericLiterals(f))) {
      if (!isBusinessConstant(n)) continue;
      const cells = byValue.get(n) ?? [];
      cells.push(`${sheet}!${addr}`);
      byValue.set(n, cells);
    }
  }

  const repeated = [...byValue.entries()]
    .filter(([, cells]) => cells.length >= REPEATED)
    .sort((a, b) => b[1].length - a[1].length);
  if (repeated.length === 0) return null;

  const worst = repeated[0][1].length;
  const locations = repeated.slice(0, 8).map(([value, cells]) => `${value} appears in ${cells.length} formulas (${cells.slice(0, 3).join(", ")}${cells.length > 3 ? ", ..." : ""})`);

  return {
    id: "hidden-fragility.hardcoded-constants",
    category: "hidden-fragility",
    severity: worst >= PERVASIVE ? "high" : "medium",
    title: repeated.length === 1
      ? "A number is written directly into many formulas"
      : `${repeated.length} numbers are written directly into many formulas`,
    soWhat: "A rate or threshold typed into the formulas themselves - a VAT rate, a margin, a price - has to be found and changed in every single cell when it moves. Miss one and the workbook quietly disagrees with itself.",
    action: "Move each of these into a single labelled input cell and point the formulas at it, so changing the rate is one edit instead of forty.",
    locations,
    count: repeated.reduce((s, [, cells]) => s + cells.length, 0),
  };
};
