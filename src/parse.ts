import * as XLSX from "xlsx";
import { unzipSync } from "fflate";
import type { ParsedContext, SheetInfo, FileFormat, FormulaCell } from "./types";

function detectFormat(zip: Record<string, Uint8Array>, isZip: boolean): FileFormat {
  if (!isZip) return "xls";
  if (zip["xl/vbaProject.bin"]) return "xlsm";
  if (zip["xl/workbook.xml"]) return "xlsx";
  if (Object.keys(zip).some(k => k.endsWith(".bin") && k.includes("workbook"))) return "xlsb";
  return "unknown";
}

const VIS: Record<number, SheetInfo["visibility"]> = { 0: "visible", 1: "hidden", 2: "veryHidden" };

const decode = (b: Uint8Array | undefined): string => (b ? new TextDecoder().decode(b) : "");

/**
 * Map each sheet name to its worksheet XML part. The order of <sheet> elements in
 * workbook.xml is the sheet order, but the *part* each one lives in is only knowable
 * through its r:id in workbook.xml.rels - sheet3.xml is not necessarily the third sheet.
 * Falls back to positional order if the rels part is missing or unparseable.
 */
function mapSheetXml(zip: Record<string, Uint8Array>, workbookXml: string, names: string[]): Record<string, string> {
  const out: Record<string, string> = {};
  const rels = decode(zip["xl/_rels/workbook.xml.rels"]);
  const idToTarget = new Map<string, string>();
  for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = /\bId="([^"]+)"/.exec(m[0])?.[1];
    const target = /\bTarget="([^"]+)"/.exec(m[0])?.[1];
    if (id && target) idToTarget.set(id, target.replace(/^\/?xl\//, "").replace(/^\.\//, ""));
  }
  const sheetEls = [...workbookXml.matchAll(/<sheet\b[^>]*\/?>/g)].map(m => m[0]);
  sheetEls.forEach((el, i) => {
    const name = /\bname="([^"]*)"/.exec(el)?.[1];
    const rid = /\br:id="([^"]+)"/.exec(el)?.[1];
    const target = (rid && idToTarget.get(rid)) || `worksheets/sheet${i + 1}.xml`;
    const part = zip[`xl/${target}`];
    const key = name ?? names[i];
    if (key !== undefined && part) out[key] = decode(part);
  });
  // Positional fallback for any sheet the rels walk missed.
  names.forEach((name, i) => {
    if (out[name] === undefined) {
      const part = zip[`xl/worksheets/sheet${i + 1}.xml`];
      if (part) out[name] = decode(part);
    }
  });
  return out;
}

/** OLE2/CFB - the container legacy .xls lives in. */
const OLE2 = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

export class NotASpreadsheetError extends Error {
  constructor(detail: string) {
    super(`Not a spreadsheet: ${detail}`);
    this.name = "NotASpreadsheetError";
  }
}

export function parseWorkbook(bytes: Uint8Array): ParsedContext {
  // Refuse anything that is not really a workbook, before SheetJS gets a chance to be
  // charitable about it. SheetJS will hand back an empty workbook for arbitrary bytes, and a
  // "no PK header, so it must be .xls" sniff then reports a renamed JPEG as a legacy-format
  // risk. Describing a file we never parsed is how a tool loses an owner's trust in one screen.
  const isZip = bytes[0] === 0x50 && bytes[1] === 0x4b;                       // "PK" - OOXML
  const isOle2 = OLE2.every((b, i) => bytes[i] === b);                        // legacy .xls
  if (!isZip && !isOle2) throw new NotASpreadsheetError("the file is not an Excel workbook");

  let zip: Record<string, Uint8Array> = {};
  let workbookXml = "";
  if (isZip) {
    try {
      zip = unzipSync(bytes);
    } catch {
      throw new NotASpreadsheetError("the file could not be opened");
    }
    workbookXml = decode(zip["xl/workbook.xml"]);
    const hasWorkbookPart = Object.keys(zip).some(k => /^xl\/workbook\.(xml|bin)$/.test(k));
    if (!hasWorkbookPart) throw new NotASpreadsheetError("the file is a zip, but there is no workbook inside it");
  }
  // NB: `bookProps: true` is SheetJS's props-only fast path and leaves SheetNames undefined;
  // omit it so both the workbook sheets and document properties are parsed.
  let wb: XLSX.WorkBook;
  try {
    wb = XLSX.read(bytes, { type: "array", cellFormula: true, cellNF: true, sheetStubs: true, bookVBA: true });
  } catch {
    throw new NotASpreadsheetError("the file could not be read - it may be password-protected or corrupted");
  }
  if (!wb.SheetNames?.length) throw new NotASpreadsheetError("the workbook has no sheets");
  const wbSheets = wb.Workbook?.Sheets ?? [];
  const sheetXml = mapSheetXml(zip, workbookXml, wb.SheetNames);

  const formulaCells: FormulaCell[] = [];
  let hasComments = false;

  const sheets: SheetInfo[] = wb.SheetNames.map((name, i) => {
    const ws = wb.Sheets[name];
    const ref = ws?.["!ref"];
    let rowCount = 0, colCount = 0;
    if (ref) {
      const r = XLSX.utils.decode_range(ref);
      rowCount = r.e.r + 1;
      colCount = r.e.c + 1;
    }
    // Walk the cells that exist, not the full rectangle. A sparse sheet near the million-row
    // ceiling - exactly the workbook scaleLimits exists for - would otherwise cost
    // rows x cols address encodings before the first check ever ran.
    for (const addr in ws) {
      if (addr[0] === "!") continue;
      const cell = ws[addr] as XLSX.CellObject;
      if (cell.f) formulaCells.push({ sheet: name, addr, f: cell.f });
      if (cell.c?.length) hasComments = true;
    }
    const hidden = wbSheets[i]?.Hidden ?? 0;
    return {
      name,
      visibility: VIS[hidden] ?? "visible",
      rowCount,
      colCount,
      isProtected: /<sheetProtection\b/.test(sheetXml[name] ?? ""),
    };
  });

  // Comments also live in their own OOXML parts (classic notes and threaded comments).
  if (Object.keys(zip).some(k => /^xl\/(comments\d+\.xml|threadedComments\/)/.test(k))) hasComments = true;

  return {
    wb, zip, workbookXml, sheetXml, formulaCells, hasComments,
    props: { author: (wb.Props?.Author as string) ?? null, lastModifiedBy: (wb.Props?.LastAuthor as string) ?? null },
    fileFormat: detectFormat(zip, isZip),
    fileSizeBytes: bytes.byteLength,
    sheets,
  };
}
