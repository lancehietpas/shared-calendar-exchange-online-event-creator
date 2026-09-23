import type { MatchMode } from "./types.js";

export const MATCH_DESCRIPTION: Record<MatchMode, string> = {
  contains:
    "case-insensitive contains, applied in the CLI. Graph does not support contains() on event.subject.",
  equals:
    "case-insensitive equals. The CLI asks Graph for subject eq when that filter is accepted, then checks again locally.",
};

export function subjectMatches(
  subject: string | null | undefined,
  query: string,
  mode: MatchMode,
): boolean {
  const haystack = (subject ?? "").trim().toLocaleLowerCase();
  const needle = query.trim().toLocaleLowerCase();
  if (!needle) {
    return false;
  }
  if (mode === "equals") {
    return haystack === needle;
  }
  return haystack.includes(needle);
}

/** OData string literal. Single quotes are escaped by doubling them. */
export function odataStringLiteral(value: string): string {
  return `'${value.replace(/'/g, "''")}'`;
}

export function subjectEqualsFilter(subject: string): string {
  return `subject eq ${odataStringLiteral(subject)}`;
}
