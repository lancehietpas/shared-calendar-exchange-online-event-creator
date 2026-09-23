import { chooseBestCreator, inferCreator, readActor } from "./creator.js";
import { rankAuditHits } from "./audit.js";
import { readExtendedProperties } from "./properties.js";
import { MATCH_DESCRIPTION } from "./subject.js";
import { parseGraphDateTime } from "./window.js";
import type {
  AuditHit,
  AuditSection,
  EventReport,
  GraphEvent,
  LookupReport,
  MatchMode,
} from "./types.js";

export function buildEventReport(
  event: GraphEvent,
  calendarEmail: string,
  auditHits: AuditHit[] | null,
  hydrateError: string | null,
): EventReport {
  const extended = readExtendedProperties(event.singleValueExtendedProperties);
  const organizer = {
    name: event.organizer?.emailAddress?.name ?? null,
    email: event.organizer?.emailAddress?.address ?? null,
  };
  const inferredCreator = inferCreator({
    calendarEmail,
    organizerName: organizer.name,
    organizerEmail: organizer.email,
    extended: extended.byName,
  });
  const audit = auditHits
    ? rankAuditHits(
        auditHits,
        { subject: event.subject, createdDateTime: event.createdDateTime },
        calendarEmail,
      )
    : null;

  return {
    id: event.id ?? null,
    subject: event.subject ?? null,
    start: event.start ?? null,
    end: event.end ?? null,
    type: event.type ?? null,
    iCalUId: event.iCalUId ?? null,
    isCancelled: event.isCancelled ?? null,
    createdDateTime: event.createdDateTime ?? null,
    lastModifiedDateTime: event.lastModifiedDateTime ?? null,
    organizer,
    isOrganizer: event.isOrganizer ?? null,
    createdBy: readActor(event.createdBy),
    lastModifiedBy: readActor(event.lastModifiedBy),
    singleValueExtendedProperties: extended.list,
    extendedProperties: extended.byName,
    inferredCreator,
    audit,
    bestCreator: chooseBestCreator(inferredCreator, audit?.matches ?? []),
    hydrateError,
  };
}

export function buildReport(input: {
  calendar: string;
  subjectQuery: string;
  match: MatchMode;
  source: "calendarView" | "events";
  window: { start: string; end: string };
  truncated: boolean;
  pagesFetched: number;
  scannedEventCount: number;
  hydratedCount: number;
  audit: AuditSection;
  events: EventReport[];
}): LookupReport {
  return {
    calendar: input.calendar,
    subjectQuery: input.subjectQuery,
    match: input.match,
    matchDescription: MATCH_DESCRIPTION[input.match],
    source: input.source,
    window: input.window,
    truncated: input.truncated,
    pagesFetched: input.pagesFetched,
    scannedEventCount: input.scannedEventCount,
    matchCount: input.events.length,
    hydratedCount: input.hydratedCount,
    audit: input.audit,
    events: input.events,
  };
}

export function eventInWindow(
  event: GraphEvent,
  window: { start: string; end: string },
): boolean {
  const start = parseGraphDateTime(event.start);
  if (Number.isNaN(start)) {
    return true;
  }
  const windowStart = Date.parse(window.start);
  const windowEnd = Date.parse(window.end);
  return start >= windowStart && start <= windowEnd;
}
