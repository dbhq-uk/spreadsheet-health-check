import type { Check } from "../types";

const XLS_ROW_CAP = 65536;
const XLSX_ROW_CAP = 1048576;
const BIG_FILE = 15 * 1024 * 1024;

export const scaleLimits: Check = (ctx) => {
  const reasons: string[] = [];
  let severity: "medium" | "high" = "medium";
  if (ctx.fileFormat === "xls") { reasons.push(`Saved in the old .xls format, which silently stops at ${XLS_ROW_CAP.toLocaleString()} rows`); severity = "high"; }
  const maxRows = Math.max(0, ...ctx.sheets.map(s => s.rowCount));
  if (maxRows > XLSX_ROW_CAP * 0.5) { reasons.push(`A sheet is over ${maxRows.toLocaleString()} rows, approaching Excel's ceiling`); severity = "high"; }
  if (ctx.fileSizeBytes > BIG_FILE) reasons.push(`The file is ${(ctx.fileSizeBytes / 1024 / 1024).toFixed(0)}MB - large enough to be slow and fragile`);
  if (reasons.length === 0) return null;
  return {
    id: "wont-scale.limits",
    category: "wont-scale",
    severity,
    title: "This workbook is near the limits of what a spreadsheet can hold",
    soWhat: "Past these limits Excel drops data or grinds to a halt - the failure mode behind real incidents where records simply vanished.",
    locations: reasons,
    count: reasons.length,
  };
};
