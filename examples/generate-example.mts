// Builds examples/example-workbook.xlsx - the "try it with an example" file the website
// offers so someone can see what the report looks like without handing over their own data.
//
// It is a deliberately invented workbook for a made-up company, built to look like the real
// thing: a pricing sheet somebody has been patching for years, a copied sheet per month, a VAT
// rate typed into every line, a couple of cells quietly showing errors, and calculation turned
// off so none of it is necessarily up to date. Every fault in it is one the checks are designed
// to find, so the report it produces is a fair demonstration rather than a staged one.
//
//   npm run example
import * as XLSX from "xlsx";
import * as fs from "node:fs";
import { zipSync, unzipSync, strToU8 } from "fflate";

// SheetJS's ESM build does not bind fs itself, so XLSX.writeFile throws until it is handed
// one. Only the generators need this; the engine in src/ only ever reads a workbook.
XLSX.set_fs(fs);
import { writeFileSync, readFileSync, mkdirSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
mkdirSync(here, { recursive: true });
const file = join(here, "example-workbook.xlsx");
const MTIME = new Date("2026-01-01T00:00:00Z");

const wb = XLSX.utils.book_new();

// Pricing - the sheet the business actually runs on, and the one nobody else can safely touch.
const pricing: any[][] = [
  ["Meridian Supplies - price list and margins"],
  ["Product", "Cost", "Markup", "Net price", "Price inc VAT", "Margin %"],
];
for (let i = 0; i < 24; i++) {
  const r = i + 3;
  pricing.push([
    `Product ${String.fromCharCode(65 + (i % 26))}${i + 1}`,
    10 + i * 1.5,
    1.4,
    { t: "n", f: `B${r}*C${r}` },
    // The VAT rate, typed into all 24 rows instead of held in one cell.
    { t: "n", f: `D${r}*1.2` },
    // One row calculates its margin differently from the other 23 - the classic hidden error.
    i === 11
      ? { t: "n", f: `(D${r}-B${r})/B${r}` }
      : { t: "n", f: `(D${r}-B${r})/D${r}` },
  ]);
}
// Two cells quietly broken right now.
pricing.push(["Discontinued line", 0, 1.4, { t: "n", f: "B27*C27" }, { t: "n", f: "D27*1.2" }, { t: "e", v: 0x07 }]);
pricing.push(["Old supplier ref", { t: "e", v: 0x17 }, 1.4, 0, 0, 0]);
const wsPricing = XLSX.utils.aoa_to_sheet(pricing);
wsPricing["!merges"] = [
  { s: { r: 0, c: 0 }, e: { r: 0, c: 5 } },   // the title band - fine, and correctly ignored
  { s: { r: 26, c: 0 }, e: { r: 27, c: 0 } }, // a merge down inside the data - not fine
];
XLSX.utils.book_append_sheet(wb, wsPricing, "Pricing");

// Summary - the sheet the owner looks at, built on references nobody can follow.
const summary: any[][] = [
  ["Board summary", "", ""],
  ["Metric", "Value", "Source"],
  ["Total net", { t: "n", f: "SUM(Pricing!D3:D26)" }, "Pricing"],
  ["Total inc VAT", { t: "n", f: "SUM(Pricing!E3:E26)" }, "Pricing"],
  // Untraceable: the cell these read is assembled while they calculate.
  ["Current month", { t: "n", f: 'SUM(INDIRECT("\'"&B7&"\'!D2:D5"))' }, "Whichever sheet B7 names"],
  ["Rolling 3 month", { t: "n", f: "SUM(OFFSET(Jan!D2,0,0,3,1))" }, "Offset from Jan"],
  ["Sheet", "Mar", ""],
  // A formula nobody is going to check by reading it.
  ["Blended margin", { t: "n", f: "IF(B3>0,IF(B4>0,IF(B5>0,IF(B6>0,INDEX(Pricing!F3:F26,MATCH(MAX(Pricing!F3:F26),Pricing!F3:F26,0))*0.4+INDEX(Pricing!F3:F26,MATCH(MIN(Pricing!F3:F26),Pricing!F3:F26,0))*0.6,0),0),0),0)" }, "Do not touch"],
];
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summary), "Summary");

// One sheet per month, copied every time - the pattern that guarantees a rebuild later.
for (const month of ["Jan", "Feb", "Mar", "Apr", "May", "Jun"]) {
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ["Customer", "Units", "Unit price", "Revenue"],
    ["Ashby Ltd", 12, 24.5, { t: "n", f: "B2*C2" }],
    ["Brackley Co", 8, 24.5, { t: "n", f: "B3*C3" }],
    ["Colefax", 15, 24.5, { t: "n", f: "B4*C4" }],
    ["Dunmore", 6, 24.5, { t: "n", f: "B5*C5" }],
  ]), month);
}

// The sheets that are not meant to be seen.
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
  ["Old rates - superseded, do not delete"],
  ["VAT", 0.175],
  ["Markup", 1.35],
]), "Old_Rates");
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
  ["scratch"],
  ["do not delete - the Summary breaks", 1],
]), "Working");

wb.Workbook = {
  Sheets: wb.SheetNames.map((name) => ({
    Hidden: name === "Old_Rates" ? 1 : name === "Working" ? 2 : 0,
  })),
} as any;
wb.Props = { Author: "Sam Whitfield", LastAuthor: "Sam Whitfield" } as XLSX.FullProperties;

// Excel always caches a formula's last result, and SheetJS marks a formula cell that has no
// value as t="e" - an error cell. Without this the example workbook would open as a sea of
// #ERRORs and the error-values finding would be meaningless. The two cells that are meant to
// be errors carry a v of their own and no formula, so they are untouched.
for (const name of wb.SheetNames) {
  const ws = wb.Sheets[name];
  for (const addr of Object.keys(ws)) {
    if (addr[0] === "!") continue;
    const cell = ws[addr] as XLSX.CellObject;
    if (cell.f && cell.v === undefined) { cell.t = "n"; cell.v = 0; }
  }
}

XLSX.writeFile(wb, file, { bookType: "xlsx" });

// Automatic calculation off, so every number on screen may be stale - and the sheet the
// Pricing tab is locked, so only whoever holds the password can change how any of it works.
const files = unzipSync(new Uint8Array(readFileSync(file)));
let wbxml = new TextDecoder().decode(files["xl/workbook.xml"]);
wbxml = wbxml.includes("<calcPr")
  ? wbxml.replace("<calcPr", '<calcPr calcMode="manual" ')
  : wbxml.replace("</workbook>", '<calcPr calcMode="manual"/></workbook>');
files["xl/workbook.xml"] = strToU8(wbxml);

const sheet1 = new TextDecoder().decode(files["xl/worksheets/sheet1.xml"])
  .replace("</worksheet>", '<sheetProtection sheet="1" objects="1" scenarios="1"/></worksheet>');
files["xl/worksheets/sheet1.xml"] = strToU8(sheet1);

writeFileSync(file, zipSync(files, { mtime: MTIME }));
console.log(`example workbook written: ${file}`);
