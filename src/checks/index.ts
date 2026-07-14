import type { Check } from "../types";
import { macros } from "./macros";
import { hiddenSheets } from "./hiddenSheets";
import { opaqueReferences } from "./opaqueReferences";
import { formulaComplexity } from "./formulaComplexity";
import { formulaSprawl } from "./formulaSprawl";
import { protectedSheets } from "./protectedSheets";
import { undocumented } from "./undocumented";
import { errorValues } from "./errorValues";
import { circularRefs } from "./circularRefs";
import { manualCalc } from "./manualCalc";
import { inconsistentFormulas } from "./inconsistentFormulas";
import { externalLinks } from "./externalLinks";
import { hardcodedConstants } from "./hardcodedConstants";
import { mergedCells } from "./mergedCells";
import { missingValidation } from "./missingValidation";
import { duplicateSheets } from "./duplicateSheets";
import { scaleLimits } from "./scaleLimits";

// NB: there is deliberately no "single author" check. creator == lastModifiedBy only
// records who saved the file last, so it fires on almost every workbook and misses a
// genuinely one-person file a colleague once opened. Key-person exposure is inferred
// from the workbook's contents below, never from its metadata. Do not reintroduce it.
export const ALL_CHECKS: Check[] = [
  // Key-person exposure - evidence that only the author can safely change this
  macros, hiddenSheets, opaqueReferences, formulaComplexity, formulaSprawl, protectedSheets, undocumented,
  // Broken today
  errorValues, circularRefs, manualCalc,
  // Hidden fragility
  inconsistentFormulas, externalLinks, hardcodedConstants, mergedCells, missingValidation,
  // Won't scale
  duplicateSheets, scaleLimits,
];
