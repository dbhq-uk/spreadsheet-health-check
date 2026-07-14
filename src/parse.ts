import * as XLSX from "xlsx";
import { unzipSync } from "fflate";
import type { ParsedContext, SheetInfo, FileFormat } from "./types";

function detectFormat(zip: Record<string, Uint8Array>, isZip: boolean): FileFormat {
  if (!isZip) return "xls";
  if (zip["xl/vbaProject.bin"]) return "xlsm";
  if (zip["xl/workbook.xml"]) return "xlsx";
  if (Object.keys(zip).some(k => k.endsWith(".bin") && k.includes("workbook"))) return "xlsb";
  return "unknown";
}

const VIS: Record<number, SheetInfo["visibility"]> = { 0: "visible", 1: "hidden", 2: "veryHidden" };

export function parseWorkbook(bytes: Uint8Array): ParsedContext {
  const isZip = bytes[0] === 0x50 && bytes[1] === 0x4b; // "PK"
  let zip: Record<string, Uint8Array> = {};
  let workbookXml = "";
  if (isZip) {
    zip = unzipSync(bytes);
    if (zip["xl/workbook.xml"]) workbookXml = new TextDecoder().decode(zip["xl/workbook.xml"]);
  }
  // NB: `bookProps: true` is SheetJS's props-only fast path and leaves SheetNames undefined;
  // omit it so both the workbook sheets and document properties are parsed.
  const wb = XLSX.read(bytes, { type: "array", cellFormula: true, cellNF: true, sheetStubs: true, bookVBA: true });
  const wbSheets = wb.Workbook?.Sheets ?? [];
  const sheets: SheetInfo[] = wb.SheetNames.map((name, i) => {
    const ws = wb.Sheets[name];
    const ref = ws?.["!ref"];
    let rowCount = 0, colCount = 0;
    if (ref) { const r = XLSX.utils.decode_range(ref); rowCount = r.e.r + 1; colCount = r.e.c + 1; }
    const hidden = wbSheets[i]?.Hidden ?? 0;
    return { name, visibility: VIS[hidden] ?? "visible", rowCount, colCount };
  });
  return {
    wb, zip, workbookXml,
    props: { author: (wb.Props?.Author as string) ?? null, lastModifiedBy: (wb.Props?.LastAuthor as string) ?? null },
    fileFormat: detectFormat(zip, isZip),
    fileSizeBytes: bytes.byteLength,
    sheets,
  };
}
