import type * as XLSX from "xlsx";
import type { ParsedContext } from "./types";

// Visits the cells that exist, not the full rectangular range - on a sparse sheet near
// Excel's row ceiling the difference is the worker stalling versus finishing.
export function forEachCell(ctx: ParsedContext, cb: (sheet: string, addr: string, cell: XLSX.CellObject) => void) {
  for (const name of ctx.wb.SheetNames) {
    const ws = ctx.wb.Sheets[name];
    if (!ws) continue;
    for (const addr in ws) {
      if (addr[0] === "!") continue;
      cb(name, addr, ws[addr] as XLSX.CellObject);
    }
  }
}
