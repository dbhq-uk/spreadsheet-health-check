import * as XLSX from "xlsx";
import type { Check, ParsedContext } from "../types";
import { relativeTemplate } from "../r1c1";

// Jan / Feb / Mar as three copies of the same sheet is the most recognisable "this will not
// scale" pattern there is: every structural change has to be made n times, and the month
// somebody forgets is the month the numbers go wrong.
//
// Detected structurally, not by name. A sheet's signature is its header row plus the set of
// distinct formula shapes on it, so "Jan"/"Feb" and "Region A"/"Region B" are caught alike,
// and two sheets that merely happen to be called January and February but do genuinely
// different things are not.
const SIMILAR = 0.8;
const GROUP = 3;      // a pair can be legitimate (this year / last year); three is a pattern
const FULL_YEAR = 6;

interface Sig { name: string; header: string; shapes: Set<string>; }

function signature(ctx: ParsedContext, name: string): Sig | null {
  const ws = ctx.wb.Sheets[name];
  if (!ws?.["!ref"]) return null;
  const range = XLSX.utils.decode_range(ws["!ref"]);

  // Header = the first row that carries any text, normalised.
  let header = "";
  for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + 4) && !header; r++) {
    const cells: string[] = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })] as XLSX.CellObject | undefined;
      if (cell?.t === "s" && typeof cell.v === "string") cells.push(cell.v.trim().toLowerCase());
    }
    if (cells.length >= 2) header = cells.join("|");
  }

  const shapes = new Set(
    ctx.formulaCells.filter(fc => fc.sheet === name).map(fc => relativeTemplate(fc.f, fc.addr))
  );
  if (!header && shapes.size === 0) return null;
  return { name, header, shapes };
}

function similarity(a: Sig, b: Sig): number {
  if (a.header !== b.header) return 0;
  // Same header and neither sheet calculates anything: repeated data dumps of one shape.
  if (a.shapes.size === 0 && b.shapes.size === 0) return 1;
  let shared = 0;
  for (const s of a.shapes) if (b.shapes.has(s)) shared++;
  const union = a.shapes.size + b.shapes.size - shared;
  return union === 0 ? 0 : shared / union;
}

export const duplicateSheets: Check = (ctx) => {
  const sigs = ctx.sheets.map(s => signature(ctx, s.name)).filter((s): s is Sig => s !== null);

  // Group sheets transitively: each unplaced sheet joins the first group it is similar to.
  const groups: Sig[][] = [];
  for (const sig of sigs) {
    const group = groups.find(g => g.some(member => similarity(member, sig) >= SIMILAR));
    if (group) group.push(sig);
    else groups.push([sig]);
  }

  const dupes = groups.filter(g => g.length >= GROUP).sort((a, b) => b.length - a.length);
  if (dupes.length === 0) return null;

  const biggest = dupes[0].length;
  const total = dupes.reduce((s, g) => s + g.length, 0);

  return {
    id: "wont-scale.duplicate-sheets",
    category: "wont-scale",
    severity: biggest >= FULL_YEAR ? "high" : "medium",
    title: dupes.length === 1
      ? `${biggest} sheets are near-identical copies of each other`
      : `${total} sheets are near-identical copies, in ${dupes.length} sets`,
    soWhat: "Copying a sheet per month, region or client means every fix has to be made in every copy. The one that gets missed is the one that quietly reports the wrong number, and consolidating them later is a job in itself.",
    action: "Collapse the copies into one sheet with a column for the thing that varies (the month, the region, the client), and report off that with a pivot.",
    locations: dupes.map(g => g.map(s => s.name).join(", ")),
    count: total,
  };
};
