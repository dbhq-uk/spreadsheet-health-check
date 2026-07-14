import type { Check } from "../types";
import { SUBSTANTIAL_FORMULA_COUNT } from "../types";

// Sheets whose name says "this is the explanation".
const DOC_SHEET = /\b(notes?|readme|read me|documentation|docs?|instructions?|guide|about|help|assumptions?|methodology|definitions?|glossary|changelog|version history|control)\b/i;

// Gated on the workbook having real logic in it. A four-formula workbook has nothing to
// document, and firing on one would be the same false positive that got the single-author
// check deleted: a finding the owner can instantly dismiss, which costs you the rest.
export const undocumented: Check = (ctx) => {
  if (ctx.formulaCells.length < SUBSTANTIAL_FORMULA_COUNT) return null;
  if (ctx.hasComments) return null;
  if (ctx.sheets.some(s => DOC_SHEET.test(s.name))) return null;

  return {
    id: "key-person.undocumented",
    category: "key-person",
    severity: "low",
    title: "Nothing in the workbook explains how it works",
    soWhat: "There is no notes sheet and not a single cell comment, so everything about how these numbers are produced - and why - lives in one person's head.",
    action: "Add a notes sheet covering what each sheet is for, where the inputs come from and which assumptions are baked in. An hour of writing is the cheapest insurance here.",
    locations: [],
    count: 1,
  };
};
