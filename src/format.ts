import type { AuditHit, EventReport, LookupReport, PersonRef } from "./types.js";

export function formatReport(report: LookupReport): string {
  const lines: string[] = [
    `Calendar: ${report.calendar}`,
    `Subject query: ${report.subjectQuery}`,
    `Match: ${report.match} (${report.matchDescription})`,
    `Source: /users/{calendar}/${report.source === "calendarView" ? "calendar/calendarView" : "calendar/events"}`,
    `Window: ${report.window.start} .. ${report.window.end}`,
    `Scanned events: ${report.scannedEventCount} across ${report.pagesFetched} page(s)`,
    `Matches: ${report.matchCount}`,
  ];
  if (report.truncated) {
    lines.push("The event list stopped at --max-pages. There may be more events in the window.");
  }
  lines.push(`Audit: ${formatAuditStatus(report)}`);
  if (report.events.length === 0) {
    lines.push("", "No matching events.");
    return lines.join("\n");
  }
  for (const event of report.events) {
    lines.push("", ...formatEvent(event));
  }
  return lines.join("\n");
}

function formatAuditStatus(report: LookupReport): string {
  const audit = report.audit;
  if (audit.status === "not_requested") {
    return audit.guidance ?? "not requested";
  }
  if (audit.status === "skipped") {
    return audit.guidance ?? "skipped";
  }
  if (audit.status === "error") {
    return `failed. ${audit.error ?? "Unknown audit error."}`;
  }
  const range = audit.window ? `${audit.window.start} .. ${audit.window.end}` : "unspecified";
  return [
    `ok, query ${audit.queryId ?? "(no id)"}`,
    `keyword ${audit.keyword ?? "(none)"}`,
    `window ${range}`,
    `scanned ${audit.scannedRecordCount ?? 0} Create record(s)`,
    audit.windowNote,
  ]
    .filter(Boolean)
    .join("; ");
}

function formatEvent(event: EventReport): string[] {
  const lines = [
    "---",
    `Subject: ${show(event.subject)}`,
    `Id: ${show(event.id)}`,
    `Type: ${show(event.type)}`,
    `iCalUId: ${show(event.iCalUId)}`,
    `Start: ${formatWhen(event.start)}`,
    `End: ${formatWhen(event.end)}`,
    `Created: ${show(event.createdDateTime)}`,
    `Last modified: ${show(event.lastModifiedDateTime)}`,
    `Organizer: ${formatPerson(event.organizer)}`,
    `createdBy: ${event.createdBy ? formatPerson(event.createdBy) : "not returned by the Graph event resource"}`,
    `lastModifiedBy: ${event.lastModifiedBy ? formatPerson(event.lastModifiedBy) : "not returned by the Graph event resource"}`,
  ];
  if (event.hydrateError) {
    lines.push(`Extended properties: unavailable (${event.hydrateError})`);
  } else if (event.singleValueExtendedProperties.length === 0) {
    lines.push("singleValueExtendedProperties: none returned");
  } else {
    lines.push("singleValueExtendedProperties:");
    for (const property of event.singleValueExtendedProperties) {
      const label = property.canonicalName ?? property.id;
      lines.push(`  ${label}: ${property.value}`);
    }
  }
  lines.push(
    `Inferred creator: ${formatPerson(event.inferredCreator)}`,
    `  source: ${event.inferredCreator.source}`,
    `  confidence: ${event.inferredCreator.confidence}`,
    `  organizer is calendar mailbox: ${formatOptionalBoolean(event.inferredCreator.organizerIsCalendarMailbox)}`,
    `  note: ${event.inferredCreator.note}`,
  );
  if (event.audit) {
    lines.push(...formatAuditMatches(event));
  }
  lines.push(
    `Best creator: ${formatPerson(event.bestCreator)}`,
    `  source: ${event.bestCreator.source}`,
    `  confidence: ${event.bestCreator.confidence}`,
    `  detail: ${event.bestCreator.detail}`,
  );
  return lines;
}

function formatAuditMatches(event: EventReport): string[] {
  const audit = event.audit;
  if (!audit || audit.matches.length === 0) {
    return ["Audit matches: none for this subject"];
  }
  const lines = [`Audit matches: ${audit.matches.length}`];
  for (const hit of audit.matches.slice(0, 5)) {
    lines.push(`  - ${formatAuditHit(hit)}`);
  }
  if (audit.matches.length > 5) {
    lines.push(`  - ${audit.matches.length - 5} more not shown`);
  }
  if (audit.distinctUsers.length > 1) {
    lines.push(`  distinct Create users: ${audit.distinctUsers.join(", ")}`);
  }
  return lines;
}

function formatAuditHit(hit: AuditHit): string {
  const who = hit.logonUserDisplayName
    ? `${hit.logonUserDisplayName} <${hit.userId ?? hit.userPrincipalName ?? "unknown"}>`
    : (hit.userId ?? hit.userPrincipalName ?? "unknown user");
  return [
    who,
    hit.logonType ? `logon ${hit.logonType}` : null,
    hit.createdDateTime ?? "time unknown",
    hit.subject ? `subject "${hit.subject}"` : null,
    hit.folderPath ? `folder ${hit.folderPath}` : null,
  ]
    .filter(Boolean)
    .join("; ");
}

function formatWhen(value: { dateTime?: string; timeZone?: string } | null): string {
  if (!value?.dateTime) {
    return "(none)";
  }
  return value.timeZone ? `${value.dateTime} (${value.timeZone})` : value.dateTime;
}

function formatPerson(person: PersonRef): string {
  if (person.name && person.email) {
    return `${person.name} <${person.email}>`;
  }
  return person.name ?? person.email ?? "(unknown)";
}

function formatOptionalBoolean(value: boolean | null): string {
  if (value === null) {
    return "unknown";
  }
  return value ? "yes" : "no";
}

function show(value: string | null): string {
  return value && value.trim() ? value : "(none)";
}
