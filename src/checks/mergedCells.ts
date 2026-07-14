import * as XLSX from "xlsx";
import type { Check } from "../types";

// Merged cells across the top of a sheet are a title, and harmless. Merged cells down inside
// the data are a trap: they break sorting and filtering, they make a range reference cover
// the wrong cells, and copying a formula across them fails outright. Only the merges below
// the header band are counted, so a titled sheet is not punished for having a title.
const HEADER_BAND_ROWS = 2;

export const mergedCells: Check = (ctx) => {
  const locations: string[] = [];
  let total = 0;

  for (const sheet of ctx.wb.SheetNames) {
    const merges = (ctx.wb.Sheets[sheet]?.["!merges"] ?? []) as XLSX.Range[];
    const inData = merges.filter(m => m.s.r >= HEADER_BAND_ROWS);
    total += inData.length;
    for (const m of inData.slice(0, 4)) {
      locations.push(`${sheet}!${XLSX.utils.encode_range(m)}`);
    }
  }
  if (total === 0) return null;

  return {
    id: "hidden-fragility.merged-cells",
    category: "hidden-fragility",
    severity: total >= 5 ? "medium" : "low",
    title: total === 1 ? "A merged cell sits inside the data" : `${total} merged cells sit inside the data`,
    soWhat: "Merged cells inside a data region break sorting and filtering, and make formulas that cover the range read the wrong cells - the kind of fault that produces a wrong total rather than an error message.",
    action: "Unmerge them and use centre-across-selection for the formatting instead, which looks identical and keeps every cell addressable.",
    locations: locations.slice(0, 20),
    count: total,
  };
};
