// test/fixtures/generate.mts
import * as XLSX from "xlsx";
import { zipSync, unzipSync, strToU8 } from "fflate";
import { writeFileSync, readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const out = (name: string) => join(here, name);

function setProps(wb: XLSX.WorkBook, author: string, lastMod: string | null) {
  wb.Props = { ...(wb.Props || {}), Author: author } as XLSX.FullProperties;
  if (lastMod) (wb.Props as any).LastAuthor = lastMod;
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

// single-author.xlsx - author == lastAuthor
{
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["x", 1]]), "S");
  setProps(wb, "Solo Owner", "Solo Owner");
  XLSX.writeFile(wb, out("single-author.xlsx"), { bookType: "xlsx" });
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
  writeFileSync(out("circular.xlsx"), zipSync(files));
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
  writeFileSync(out("external-links.xlsx"), zipSync(files));
}

// macros.xlsm - stub vbaProject.bin entry
{
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["macro book"]]), "S");
  setProps(wb, "Alice", "Bob");
  XLSX.writeFile(wb, out("macros.xlsm"), { bookType: "xlsm" });
  const files = unzipSync(new Uint8Array(readFileSync(out("macros.xlsm"))));
  files["xl/vbaProject.bin"] = new Uint8Array([0xcf, 0xd0, 0xe0, 0x11, 0x00, 0x00]);
  writeFileSync(out("macros.xlsm"), zipSync(files));
}

// legacy.xls - old BIFF8 format
{
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["old", 1]]), "S");
  setProps(wb, "Alice", "Bob");
  XLSX.writeFile(wb, out("legacy.xls"), { bookType: "xls" });
}

console.log("fixtures generated");
