import type { Check } from "../types";

export const singleAuthor: Check = (ctx) => {
  const { author, lastModifiedBy } = ctx.props;
  if (!author) return null;
  const single = !lastModifiedBy || lastModifiedBy.trim() === author.trim();
  if (!single) return null;
  return {
    id: "key-person.single-author",
    category: "key-person",
    severity: "medium",
    title: "One person appears to own this workbook",
    soWhat: "If the only person who understands it is away or leaves, the process it runs stops with them.",
    locations: [author],
    count: 1,
  };
};
