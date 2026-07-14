export type Category = "key-person" | "broken-today" | "hidden-fragility" | "wont-scale";
export type Severity = "info" | "low" | "medium" | "high";
export type Band = "Low" | "Moderate" | "High" | "Critical";
export type FileFormat = "xlsx" | "xlsm" | "xls" | "xlsb" | "unknown";

export interface Finding {
  id: string;
  category: Category;
  severity: Severity;
  title: string;
  /** Why an owner should care, in one sentence. */
  soWhat: string;
  /** What to do about it - the concrete next step, in one sentence. */
  action: string;
  locations: string[];
  count: number;
}

export interface CategoryResult {
  category: Category;
  label: string;
  subScore: number;
  findings: Finding[];
}

export interface HealthReport {
  riskBand: Band;
  score: number;
  verdict: string;
  categories: CategoryResult[];
  findings: Finding[];
  meta: {
    sheetCount: number;
    fileFormat: FileFormat;
    author: string | null;
    lastModifiedBy: string | null;
    fileSizeBytes: number;
    formulaCount: number;
  };
}

export interface SheetInfo {
  name: string;
  visibility: "visible" | "hidden" | "veryHidden";
  rowCount: number;
  colCount: number;
  /** Sheet-level protection is on (cells locked against editing). */
  isProtected: boolean;
}

export interface FormulaCell {
  sheet: string;
  addr: string;
  /** The formula text, without a leading "=". */
  f: string;
}

export interface ParsedContext {
  wb: import("xlsx").WorkBook;
  zip: Record<string, Uint8Array>;
  props: { author: string | null; lastModifiedBy: string | null };
  fileFormat: FileFormat;
  fileSizeBytes: number;
  sheets: SheetInfo[];
  workbookXml: string;
  /** Every formula in the workbook, walked once so checks do not each re-walk the grid. */
  formulaCells: FormulaCell[];
  /** Worksheet XML by sheet name (OOXML only; empty for legacy .xls). Carries the
   *  dataValidation and sheetProtection elements SheetJS does not surface on read. */
  sheetXml: Record<string, string>;
  /** The workbook carries cell comments/notes anywhere (classic or threaded). */
  hasComments: boolean;
}

export type Check = (ctx: ParsedContext) => Finding | null;

/** A workbook with less logic than this has nothing worth documenting, validating or
 *  spreading across sheets - firing the "absence of X" checks on it would be the same
 *  false-positive trap the removed single-author check fell into. */
export const SUBSTANTIAL_FORMULA_COUNT = 20;
