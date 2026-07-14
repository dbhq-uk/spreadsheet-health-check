import type { Check } from "../types";

export const macros: Check = (ctx) => {
  const hasVba = !!ctx.zip["xl/vbaProject.bin"] || ctx.fileFormat === "xlsm";
  if (!hasVba) return null;
  return {
    id: "key-person.macros",
    category: "key-person",
    severity: "high",
    title: "The workbook contains macros (hidden VBA code)",
    soWhat: "Logic hidden in code, usually written and understood by one person, is invisible on the page and rarely documented.",
    locations: [],
    count: 1,
  };
};
