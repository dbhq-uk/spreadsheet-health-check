import type { Check } from "../types";

// Sheet protection is often sensible practice (locking formula cells so users only type into
// inputs), so on its own it is weak evidence and scored low. It earns its place because the
// combination it appears in matters: a locked workbook full of macros and untraceable
// references is one nobody but the author can touch, by design.
export const protectedSheets: Check = (ctx) => {
  const locked = ctx.sheets.filter(s => s.isProtected);
  const workbookLocked = /<workbookProtection\b/.test(ctx.workbookXml);
  if (locked.length === 0 && !workbookLocked) return null;

  const locations = locked.map(s => s.name);
  if (workbookLocked) locations.push("The workbook structure itself is locked");

  return {
    id: "key-person.protected-sheets",
    category: "key-person",
    severity: "low",
    title: workbookLocked && locked.length === 0
      ? "The workbook structure is locked"
      : locked.length === 1 ? "A sheet is locked against editing" : `${locked.length} sheets are locked against editing`,
    soWhat: "Locking is often sensible, but it also means the only person who can change how this works is whoever holds the password.",
    action: "Make sure someone other than the author holds the password, and that it is written down somewhere the business controls.",
    locations,
    count: locked.length || 1,
  };
};
