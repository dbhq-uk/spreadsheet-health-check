// test/fixtures/generate.mts
import * as XLSX from "xlsx";
import { zipSync, unzipSync, strToU8 } from "fflate";
import { writeFileSync, readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const out = (name: string) => join(here, name);

// fflate stamps the current time into every zip entry unless told otherwise, which used to
// make the re-zipped fixtures change bytes on every run. Pinning mtime makes generation
// reproducible, so regenerating is a no-op in git unless a fixture's content actually changed.
const MTIME = new Date("2026-01-01T00:00:00Z");
const rezip = (files: Record<string, Uint8Array>) => zipSync(files, { mtime: MTIME });

function setProps(wb: XLSX.WorkBook, author: string, lastMod: string | null) {
  wb.Props = { ...(wb.Props || {}), Author: author } as XLSX.FullProperties;
  if (lastMod) (wb.Props as any).LastAuthor = lastMod;
}

/** Patch a worksheet part in place - for the OOXML elements SheetJS will not write. */
function patchSheetXml(file: string, sheet: number, insert: string) {
  const files = unzipSync(new Uint8Array(readFileSync(out(file))));
  const key = `xl/worksheets/sheet${sheet}.xml`;
  const xml = new TextDecoder().decode(files[key]).replace("</worksheet>", `${insert}</worksheet>`);
  files[key] = strToU8(xml);
  writeFileSync(out(file), rezip(files));
}

// clean.xlsx - one visible sheet, consistent formulas, two authors, no issues
{
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ["Item", "Qty", "Price", "Total"],
    ["A", 2, 10, { t: "n", f: "B2*C2" }],
    ["B", 3, 10, { t: "n", f: "B3*C3" }],
    ["C", 4, 10, { t: "n", f: "B4*C4" }],
  ]);
  XLSX.utils.book_append_sheet(wb, ws, "Sales");
  setProps(wb, "Alice", "Bob");
  XLSX.writeFile(wb, out("clean.xlsx"), { bookType: "xlsx" });
}

// errors.xlsx - cells carrying error values
{
  const wb = XLSX.utils.book_new();
  const ws: XLSX.WorkSheet = {
    "!ref": "A1:B2",
    A1: { t: "s", v: "ok" },
    B1: { t: "e", v: 0x07 }, // #DIV/0!
    A2: { t: "e", v: 0x0f }, // #VALUE!
    B2: { t: "e", v: 0x17 }, // #REF!
  };
  XLSX.utils.book_append_sheet(wb, ws, "Calc");
  setProps(wb, "Alice", "Bob");
  XLSX.writeFile(wb, out("errors.xlsx"), { bookType: "xlsx" });
}

// hidden.xlsx - one visible, one hidden, one veryHidden
{
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["visible"]]), "Main");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["hidden"]]), "Old");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["secret"]]), "Working");
  wb.Workbook = { Sheets: [{ Hidden: 0 }, { Hidden: 1 }, { Hidden: 2 }] } as any;
  setProps(wb, "Alice", "Bob");
  XLSX.writeFile(wb, out("hidden.xlsx"), { bookType: "xlsx" });
}

// inconsistent.xlsx - a column of formulas with one outlier
{
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ["n", "double"],
    [1, { t: "n", f: "A2*2" }],
    [2, { t: "n", f: "A3*2" }],
    [3, { t: "n", f: "A4*3" }], // outlier: *3 not *2
    [4, { t: "n", f: "A5*2" }],
    [5, { t: "n", f: "A6*2" }],
  ]);
  XLSX.utils.book_append_sheet(wb, ws, "Calc");
  setProps(wb, "Alice", "Bob");
  XLSX.writeFile(wb, out("inconsistent.xlsx"), { bookType: "xlsx" });
}

// totals-row.xlsx - column B ends in a legitimate SUM foot (must not flag);
// column C ends in a genuine one-off outlier (must still flag)
{
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ["n", "double", "plus one"],
    [1, { t: "n", f: "A2*2" }, { t: "n", f: "A2+1" }],
    [2, { t: "n", f: "A3*2" }, { t: "n", f: "A3+1" }],
    [3, { t: "n", f: "A4*2" }, { t: "n", f: "A4+1" }],
    [4, { t: "n", f: "A5*2" }, { t: "n", f: "A5+1" }],
    [5, { t: "n", f: "SUM(B2:B5)" }, { t: "n", f: "A6*3" }],
  ]);
  XLSX.utils.book_append_sheet(wb, ws, "Calc");
  setProps(wb, "Alice", "Bob");
  XLSX.writeFile(wb, out("totals-row.xlsx"), { bookType: "xlsx" });
}

// functions.xlsx - a consistent column of formulas using a function name (LOG10)
// regression fixture for the r1c1 tokeniser: must not be flagged as inconsistent
{
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ["n", "log"],
    [1, { t: "n", f: "LOG10(A2)*2" }],
    [2, { t: "n", f: "LOG10(A3)*2" }],
    [3, { t: "n", f: "LOG10(A4)*2" }],
    [4, { t: "n", f: "LOG10(A5)*2" }],
    [5, { t: "n", f: "LOG10(A6)*2" }],
  ]);
  XLSX.utils.book_append_sheet(wb, ws, "Calc");
  setProps(wb, "Alice", "Bob");
  XLSX.writeFile(wb, out("functions.xlsx"), { bookType: "xlsx" });
}

// circular.xlsx - iterative calc enabled + A1<->B1 cycle
{
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([[{ t: "n", f: "B1+1" }, { t: "n", f: "A1+1" }]]);
  XLSX.utils.book_append_sheet(wb, ws, "Loop");
  setProps(wb, "Alice", "Bob");
  XLSX.writeFile(wb, out("circular.xlsx"), { bookType: "xlsx" });
  // inject <calcPr iterate="1"/> into xl/workbook.xml
  const buf = readFileSync(out("circular.xlsx"));
  const files = unzipSync(new Uint8Array(buf));
  let wbxml = new TextDecoder().decode(files["xl/workbook.xml"]);
  if (!wbxml.includes("<calcPr")) {
    wbxml = wbxml.replace("</workbook>", '<calcPr iterate="1" iterateCount="100"/></workbook>');
  } else {
    wbxml = wbxml.replace("<calcPr", '<calcPr iterate="1" iterateCount="100" ');
  }
  files["xl/workbook.xml"] = strToU8(wbxml);
  writeFileSync(out("circular.xlsx"), rezip(files));
}

// external-links.xlsx - synthesise an xl/externalLinks part
{
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["ref", { t: "n", v: 1 }]]), "S");
  setProps(wb, "Alice", "Bob");
  XLSX.writeFile(wb, out("external-links.xlsx"), { bookType: "xlsx" });
  const files = unzipSync(new Uint8Array(readFileSync(out("external-links.xlsx"))));
  files["xl/externalLinks/externalLink1.xml"] = strToU8(
    '<?xml version="1.0"?><externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><externalBook/></externalLink>'
  );
  files["xl/externalLinks/_rels/externalLink1.xml.rels"] = strToU8(
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="file:///C:/Reports/prices.xlsx" TargetMode="External"/></Relationships>'
  );
  writeFileSync(out("external-links.xlsx"), rezip(files));
}

// macros.xlsm - stub vbaProject.bin entry
{
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["macro book"]]), "S");
  setProps(wb, "Alice", "Bob");
  XLSX.writeFile(wb, out("macros.xlsm"), { bookType: "xlsm" });
  const files = unzipSync(new Uint8Array(readFileSync(out("macros.xlsm"))));
  files["xl/vbaProject.bin"] = new Uint8Array([0xcf, 0xd0, 0xe0, 0x11, 0x00, 0x00]);
  writeFileSync(out("macros.xlsm"), rezip(files));
}

// legacy.xls - old BIFF8 format
{
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["old", 1]]), "S");
  setProps(wb, "Alice", "Bob");
  XLSX.writeFile(wb, out("legacy.xls"), { bookType: "xls" });
}

// key-person.xlsx - the rebuilt key-person signal: untraceable references, formulas nobody
// can read, sprawl instead of a pattern, a locked sheet, and not a word of documentation
{
  const wb = XLSX.utils.book_new();
  const rows: any[][] = [["in", "opaque", "complex", "sprawl"]];
  const FUNCS = ["SUM", "MAX", "MIN", "AVERAGE"];
  const OPS = ["+", "-", "*", "/"];
  for (let i = 0; i < 30; i++) {
    const r = i + 2;
    // Sprawl means thirty different formula *shapes*, so vary the function, the operator and
    // the range - never the constants. Every literal here stays under 13 so this column cannot
    // trip the hardcoded-constants check: the two signals must be testable in isolation.
    const span = 1 + (i % 3);
    const sprawl = `${FUNCS[i % 4]}(A${Math.max(2, r - span)}:A${r})${OPS[i % 4]}${1 + (i % 9)}`;
    rows.push([
      i + 1,
      // opaque: the reference is built at calculation time, so it cannot be read off the page
      i < 6 ? { t: "n", f: `INDIRECT("A"&${r})*OFFSET($A$1,${i},0)` } : null,
      // complex: deep nesting and a chain of lookups
      i < 4
        ? { t: "n", f: `IF(A${r}>1,IF(A${r}>2,IF(A${r}>3,IF(A${r}>4,IF(A${r}>5,IF(A${r}>6,IF(A${r}>7,INDEX($A$2:$A$31,MATCH(A${r},$A$2:$A$31,0))+INDEX($A$2:$A$31,MATCH(A${r},$A$2:$A$31,0)),0),0),0),0),0),0),0)` }
        : null,
      { t: "n", f: sprawl },
    ]);
  }
  const ws = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, "Model");
  setProps(wb, "Alice", "Bob");
  XLSX.writeFile(wb, out("key-person.xlsx"), { bookType: "xlsx" });
  patchSheetXml("key-person.xlsx", 1, '<sheetProtection sheet="1" objects="1" scenarios="1"/>');
}

// manual-calc.xlsx - automatic calculation switched off, so what is on screen may be stale
{
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ["Qty", "Price", "Total"],
    [2, 10, { t: "n", f: "A2*B2" }],
    [3, 10, { t: "n", f: "A3*B3" }],
  ]), "Sales");
  setProps(wb, "Alice", "Bob");
  XLSX.writeFile(wb, out("manual-calc.xlsx"), { bookType: "xlsx" });
  const files = unzipSync(new Uint8Array(readFileSync(out("manual-calc.xlsx"))));
  let wbxml = new TextDecoder().decode(files["xl/workbook.xml"]);
  wbxml = wbxml.includes("<calcPr")
    ? wbxml.replace("<calcPr", '<calcPr calcMode="manual" ')
    : wbxml.replace("</workbook>", '<calcPr calcMode="manual"/></workbook>');
  files["xl/workbook.xml"] = strToU8(wbxml);
  writeFileSync(out("manual-calc.xlsx"), rezip(files));
}

// hardcoded.xlsx - a VAT rate typed into every formula instead of held in one cell.
// The column-index 2 in each VLOOKUP is the false positive the check must NOT report.
{
  const wb = XLSX.utils.book_new();
  const rows: any[][] = [["Item", "Net", "Gross"]];
  for (let i = 0; i < 25; i++) {
    const r = i + 2;
    rows.push([`Item ${i + 1}`, 100 + i, { t: "n", f: `VLOOKUP(A${r},$A$2:$B$26,2,FALSE)*1.2` }]);
  }
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), "Invoice");
  setProps(wb, "Alice", "Bob");
  XLSX.writeFile(wb, out("hardcoded.xlsx"), { bookType: "xlsx" });
}

// duplicate-sheets.xlsx - the same sheet copied per month
{
  const wb = XLSX.utils.book_new();
  for (const month of ["Jan", "Feb", "Mar", "Apr"]) {
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
      ["Item", "Qty", "Price", "Total"],
      ["A", 2, 10, { t: "n", f: "B2*C2" }],
      ["B", 3, 10, { t: "n", f: "B3*C3" }],
      ["C", 4, 10, { t: "n", f: "B4*C4" }],
    ]), month);
  }
  setProps(wb, "Alice", "Bob");
  XLSX.writeFile(wb, out("duplicate-sheets.xlsx"), { bookType: "xlsx" });
}

// merged.xlsx - merged cells down inside the data region, not just a title banner
{
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ["Quarterly sales"],            // a title merge across the top: legitimate, must not fire
    ["Region", "Q1", "Q2"],
    ["North", 10, 20],
    ["South", 30, 40],
    ["Total", 40, 60],
  ]);
  ws["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 2 } },  // title band - ignored
    { s: { r: 2, c: 0 }, e: { r: 3, c: 0 } },  // inside the data - flagged
    { s: { r: 4, c: 1 }, e: { r: 4, c: 2 } },  // inside the data - flagged
  ];
  XLSX.utils.book_append_sheet(wb, ws, "Sales");
  setProps(wb, "Alice", "Bob");
  XLSX.writeFile(wb, out("merged.xlsx"), { bookType: "xlsx" });
}

// documented.xlsx - the negative control. A substantial workbook (40 formulas) that is well
// built: a Notes sheet, a cell comment, data validation on the inputs, one repeated formula
// pattern. None of the "absence of X" checks may fire on this - if they do, they would fire on
// every good workbook too, which is exactly why the single-author check was deleted.
{
  const wb = XLSX.utils.book_new();
  const rows: any[][] = [["Item", "Qty", "Price", "Total"]];
  for (let i = 0; i < 40; i++) {
    const r = i + 2;
    rows.push([`Item ${i + 1}`, i + 1, 10, { t: "n", f: `B${r}*C${r}` }]);
  }
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws.A1.c = [{ a: "Alice", t: "Enter one row per line item. Prices come from the price list." }] as any;
  XLSX.utils.book_append_sheet(wb, ws, "Sales");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ["Notes"],
    ["Qty is typed in; Price comes from the agreed price list; Total is Qty x Price."],
    ["Owner: the sales desk. Reviewed monthly."],
  ]), "Notes");
  setProps(wb, "Alice", "Bob");
  XLSX.writeFile(wb, out("documented.xlsx"), { bookType: "xlsx" });
  patchSheetXml("documented.xlsx", 1,
    '<dataValidations count="1"><dataValidation type="whole" operator="greaterThan" sqref="B2:B41"><formula1>0</formula1></dataValidation></dataValidations>');
}

console.log("fixtures generated");
