import * as XLSX from "xlsx";
import type { Check, ParsedContext } from "../types";
import { relativeTemplate } from "../r1c1";
import { functionsUsed } from "../formula";

interface Run { sheet: string; cells: { addr: string; tpl: string; f: string }[]; }
interface Placed { r: number; addr: string; tpl: string; f: string }

// A totals row is contiguous with the column it sums, so it always lands at the foot of a
// run - but it is a deliberate change of pattern, not the wrong-cell copy this check hunts.
const AGGREGATES = /^(SUM|SUMIF|SUMIFS|AVERAGE|AVERAGEIF|AVERAGEIFS|COUNT|COUNTA|COUNTIF|COUNTIFS|MIN|MAX|MEDIAN|SUBTOTAL|AGGREGATE)$/;
const isAggregateFoot = (run: Run, index: number) =>
  index === run.cells.length - 1 && functionsUsed(run.cells[index].f).some(fn => AGGREGATES.test(fn));

function columnRuns(ctx: ParsedContext): Run[] {
  // Group the formula cells collected during parse by sheet, then by column, rather than
  // re-walking the grid: proportional to the formulas that exist, not to the sheet's bounding
  // rectangle. Nested maps rather than a composite string key, so a sheet name can never
  // collide with a column index however it is punctuated.
  const bySheet = new Map<string, Map<number, Placed[]>>();
  for (const { sheet, addr, f } of ctx.formulaCells) {
    const { r, c } = XLSX.utils.decode_cell(addr);
    let columns = bySheet.get(sheet);
    if (!columns) bySheet.set(sheet, (columns = new Map()));
    let cells = columns.get(c);
    if (!cells) columns.set(c, (cells = []));
    cells.push({ r, addr, tpl: relativeTemplate(f, addr), f });
  }

  const runs: Run[] = [];
  for (const [sheet, columns] of bySheet) {
    for (const cells of columns.values()) {
      cells.sort((a, b) => a.r - b.r);
      let cur: Run["cells"] = [];
      let prevRow = Number.NaN;
      for (const { r, addr, tpl, f } of cells) {
        // A row gap means a non-formula cell (or nothing) interrupts the column: end the run.
        if (r !== prevRow + 1 && cur.length) { if (cur.length >= 3) runs.push({ sheet, cells: cur }); cur = []; }
        cur.push({ addr, tpl, f });
        prevRow = r;
      }
      if (cur.length >= 3) runs.push({ sheet, cells: cur });
    }
  }
  return runs;
}

export const inconsistentFormulas: Check = (ctx) => {
  const outliers: string[] = [];
  for (const run of columnRuns(ctx)) {
    const counts = new Map<string, number>();
    for (const { tpl } of run.cells) counts.set(tpl, (counts.get(tpl) ?? 0) + 1);
    if (counts.size < 2) continue;
    const [majTpl, majN] = [...counts.entries()].sort((a, b) => b[1] - a[1])[0];
    // only flag if there is a clear majority (>60%) and few outliers
    if (majN / run.cells.length < 0.6) continue;
    run.cells.forEach(({ addr, tpl }, i) => {
      if (tpl !== majTpl && !isAggregateFoot(run, i)) outliers.push(`${run.sheet}!${addr}`);
    });
  }
  if (outliers.length === 0) return null;
  return {
    id: "hidden-fragility.inconsistent-formulas",
    category: "hidden-fragility",
    severity: "high",
    title: outliers.length === 1 ? "A formula breaks the pattern of the cells around it" : `${outliers.length} formulas break the pattern of the cells around them`,
    soWhat: "When one cell in a column is calculated differently from its neighbours, it is usually a mistake - and the classic source of a wrong total no one spots.",
    action: "Check each flagged cell against its neighbours, then copy the correct formula across the whole column so the pattern holds.",
    locations: outliers.slice(0, 20),
    count: outliers.length,
  };
};
