import * as XLSX from "xlsx";
import type { ParsedContext } from "./types";

export function forEachCell(ctx: ParsedContext, cb: (sheet: string, addr: string, cell: XLSX.CellObject) => void) {
  for (const name of ctx.wb.SheetNames) {
    const ws = ctx.wb.Sheets[name];
    if (!ws || !ws["!ref"]) continue;
    const range = XLSX.utils.decode_range(ws["!ref"]);
    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[addr] as XLSX.CellObject | undefined;
        if (cell) cb(name, addr, cell);
      }
    }
  }
}
