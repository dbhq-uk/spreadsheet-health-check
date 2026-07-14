export type Category = "key-person" | "broken-today" | "hidden-fragility" | "wont-scale";
export type Severity = "info" | "low" | "medium" | "high";
export type Band = "Low" | "Moderate" | "High" | "Critical";
export type FileFormat = "xlsx" | "xlsm" | "xls" | "xlsb" | "unknown";

export interface Finding {
  id: string;
  category: Category;
  severity: Severity;
  title: string;
  soWhat: string;
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
  };
}

export interface SheetInfo {
  name: string;
  visibility: "visible" | "hidden" | "veryHidden";
  rowCount: number;
  colCount: number;
}

export interface ParsedContext {
  wb: import("xlsx").WorkBook;
  zip: Record<string, Uint8Array>;
  props: { author: string | null; lastModifiedBy: string | null };
  fileFormat: FileFormat;
  fileSizeBytes: number;
  sheets: SheetInfo[];
  workbookXml: string;
}

export type Check = (ctx: ParsedContext) => Finding | null;
