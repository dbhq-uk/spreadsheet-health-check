import type { Check } from "../types";
import { SUBSTANTIAL_FORMULA_COUNT } from "../types";
import { relativeTemplate } from "../r1c1";

// A well-built workbook is a few formula patterns repeated down columns: a hundred cells,
// a handful of shapes. A workbook someone has been hand-patching for years is hundreds of
// one-off formulas, each subtly different. Measure that as the ratio of distinct relativised
// formula shapes to formula cells: low means a pattern, high means a pile of exceptions.
const MANY_SHAPES = 25;
const SPRAWL_RATIO = 0.5;
const SEVERE_SHAPES = 60;
const SEVERE_RATIO = 0.7;

export const formulaSprawl: Check = (ctx) => {
  const total = ctx.formulaCells.length;
  if (total < SUBSTANTIAL_FORMULA_COUNT) return null;

  const shapes = new Set(ctx.formulaCells.map(({ addr, f }) => relativeTemplate(f, addr)));
  const unique = shapes.size;
  const ratio = unique / total;
  if (unique < MANY_SHAPES || ratio < SPRAWL_RATIO) return null;

  const pct = Math.round(ratio * 100);
  return {
    id: "key-person.formula-sprawl",
    category: "key-person",
    severity: unique >= SEVERE_SHAPES && ratio >= SEVERE_RATIO ? "high" : "medium",
    title: "The formulas are one-offs rather than a consistent pattern",
    soWhat: `${unique} of the ${total} formulas are a different shape from each other (${pct}%). A workbook built on repeating patterns can be understood by anyone; one built from hundreds of individual exceptions only makes sense to the person who added them.`,
    action: "Have the calculations rebuilt on a consistent pattern per column, so the workbook can be handed over rather than interpreted.",
    locations: [],
    count: unique,
  };
};
