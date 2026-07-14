import type { Check } from "../types";

export const hiddenSheets: Check = (ctx) => {
  const hidden = ctx.sheets.filter(s => s.visibility !== "visible");
  if (hidden.length === 0) return null;
  const veryHidden = hidden.some(s => s.visibility === "veryHidden");
  return {
    id: "key-person.hidden-sheets",
    category: "key-person",
    severity: veryHidden ? "high" : "medium",
    title: hidden.length === 1 ? "A sheet is hidden from view" : `${hidden.length} sheets are hidden from view`,
    soWhat: "Hidden sheets often hold the real logic or data - out of sight, so no one reviews them and errors go unnoticed.",
    action: "Unhide every sheet and look at what is in there. Anything still needed belongs in plain sight; anything that is not should be deleted rather than hidden.",
    locations: hidden.map(s => s.name),
    count: hidden.length,
  };
};
