import type { Check } from "../types";

export const externalLinks: Check = (ctx) => {
  const parts = Object.keys(ctx.zip).filter(k => /^xl\/externalLinks\/externalLink\d+\.xml$/.test(k));
  if (parts.length === 0) return null;
  const targets: string[] = [];
  for (const k of Object.keys(ctx.zip)) {
    if (/^xl\/externalLinks\/_rels\//.test(k)) {
      const xml = new TextDecoder().decode(ctx.zip[k]);
      for (const m of xml.matchAll(/Target="([^"]+)"[^>]*TargetMode="External"/g)) {
        targets.push(m[1].replace(/^file:\/*/, ""));
      }
    }
  }
  return {
    id: "hidden-fragility.external-links",
    category: "hidden-fragility",
    severity: "medium",
    title: parts.length === 1 ? "The workbook depends on another file" : `The workbook depends on ${parts.length} other files`,
    soWhat: "It secretly pulls numbers from files that can be moved, renamed or deleted by someone who has no idea this workbook needs them.",
    action: "Either bring the data this workbook needs into it, or move both files somewhere their paths are stable and owned - a link to someone else's desktop is a matter of time.",
    locations: (targets.length ? targets : parts).slice(0, 20),
    count: parts.length,
  };
};
