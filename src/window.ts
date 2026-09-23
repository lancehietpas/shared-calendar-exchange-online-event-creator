const DAY_MS = 24 * 60 * 60 * 1000;

export interface TimeWindow {
  start: string;
  end: string;
}

export function defaultCalendarWindow(now = new Date()): TimeWindow {
  return {
    start: new Date(now.getTime() - 90 * DAY_MS).toISOString(),
    end: new Date(now.getTime() + 180 * DAY_MS).toISOString(),
  };
}

/**
 * Graph calendar dateTime values often omit the offset and use more than three
 * fractional digits, for example "2026-09-23T15:00:00.0000000" with timeZone UTC.
 * Date.parse would treat a zone-less value as local time.
 */
export function parseGraphDateTime(
  value: { dateTime?: string; timeZone?: string } | null | undefined,
): number {
  if (!value?.dateTime) {
    return Number.NaN;
  }
  let text = value.dateTime.trim().replace(/(\.\d{3})\d+/, "$1");
  const hasZone = /(?:Z|[+-]\d{2}:\d{2})$/i.test(text);
  const zone = (value.timeZone ?? "").trim().toUpperCase();
  if (!hasZone) {
    // This client sends Prefer: outlook.timezone="UTC". A Windows time zone
    // name means that preference was ignored, so don't guess an offset.
    if (zone === "UTC" || zone === "GMT" || zone === "") {
      text = `${text}Z`;
    } else {
      return Number.NaN;
    }
  }
  return Date.parse(text);
}

export function parseBoundary(value: string, which: "start" | "end"): string {
  const trimmed = value.trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
    return which === "start" ? `${trimmed}T00:00:00.000Z` : `${trimmed}T23:59:59.000Z`;
  }
  const parsed = new Date(trimmed);
  if (Number.isNaN(parsed.getTime())) {
    throw new Error(`Invalid --${which} value "${value}". Use YYYY-MM-DD or an ISO-8601 datetime.`);
  }
  return parsed.toISOString();
}

export interface AuditWindow extends TimeWindow {
  truncatedFuture: boolean;
  truncatedRetention: boolean;
  empty: boolean;
  note: string | null;
}

/**
 * Purview Audit (Standard) keeps mailbox audit records for 180 days.
 * The query also cannot search future activity.
 */
export function auditSearchWindow(calendarWindow: TimeWindow, now = new Date()): AuditWindow {
  const retentionStart = new Date(now.getTime() - 180 * DAY_MS);
  let start = new Date(calendarWindow.start);
  let end = new Date(calendarWindow.end);
  let truncatedFuture = false;
  let truncatedRetention = false;

  if (end > now) {
    end = now;
    truncatedFuture = true;
  }
  if (start < retentionStart) {
    start = retentionStart;
    truncatedRetention = true;
  }

  const empty = start.getTime() >= end.getTime();
  const notes: string[] = [];
  if (truncatedFuture) {
    notes.push("The audit query end was clamped to the current time.");
  }
  if (truncatedRetention) {
    notes.push(
      "The audit query start was clamped to the past 180 days (Microsoft Purview Audit Standard retention). If this tenant has a longer audit retention policy, pass a --start/--end window inside that retention period.",
    );
  }
  if (empty) {
    notes.push("The calendar window does not overlap audit history that this tool will query.");
  }

  return {
    start: start.toISOString(),
    end: end.toISOString(),
    truncatedFuture,
    truncatedRetention,
    empty,
    note: notes.length > 0 ? notes.join(" ") : null,
  };
}
