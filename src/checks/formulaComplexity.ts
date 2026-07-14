import type { Check } from "../types";
import { functionsUsed, maxDepth } from "../formula";

const DEEP_NESTING = 7;         // brackets deep - beyond this a formula stops being readable
const VERY_DEEP_NESTING = 10;
const LONG_FORMULA = 250;       // characters
const VERY_LONG_FORMULA = 500;
const LOOKUP_CHAIN = 3;         // lookups stacked inside one formula

const LOOKUPS = new Set(["VLOOKUP", "HLOOKUP", "XLOOKUP", "LOOKUP", "INDEX", "MATCH"]);

export const formulaComplexity: Check = (ctx) => {
  const locations: string[] = [];
  let extreme = false;

  for (const { sheet, addr, f } of ctx.formulaCells) {
    const depth = maxDepth(f);
    const lookups = functionsUsed(f).filter(fn => LOOKUPS.has(fn)).length;
    const reasons: string[] = [];
    if (depth >= DEEP_NESTING) reasons.push(`${depth} levels deep`);
    if (f.length >= LONG_FORMULA) reasons.push(`${f.length} characters`);
    if (lookups >= LOOKUP_CHAIN) reasons.push(`${lookups} lookups chained`);
    if (reasons.length === 0) continue;
    if (depth >= VERY_DEEP_NESTING || f.length >= VERY_LONG_FORMULA) extreme = true;
    locations.push(`${sheet}!${addr} (${reasons.join(", ")})`);
  }

  if (locations.length === 0) return null;
  return {
    id: "key-person.formula-complexity",
    category: "key-person",
    severity: extreme || locations.length >= 10 ? "high" : "medium",
    title: locations.length === 1
      ? "A formula is too complex to check by reading it"
      : `${locations.length} formulas are too complex to check by reading them`,
    soWhat: "Nobody can verify a formula this long or this deeply nested at a glance, so mistakes inside them are never spotted - and only the person who wrote them can safely change them.",
    action: "Break each one into named intermediate steps in helper columns, so the calculation can be followed - and checked - one step at a time.",
    locations: locations.slice(0, 20),
    count: locations.length,
  };
};
