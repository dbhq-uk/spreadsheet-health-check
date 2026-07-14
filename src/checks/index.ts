import type { Check } from "../types";
import { singleAuthor } from "./singleAuthor";
import { macros } from "./macros";
import { hiddenSheets } from "./hiddenSheets";
import { errorValues } from "./errorValues";
import { circularRefs } from "./circularRefs";
import { inconsistentFormulas } from "./inconsistentFormulas";
import { externalLinks } from "./externalLinks";
import { scaleLimits } from "./scaleLimits";

export const ALL_CHECKS: Check[] = [
  singleAuthor, macros, hiddenSheets,
  errorValues, circularRefs,
  inconsistentFormulas, externalLinks,
  scaleLimits,
];
