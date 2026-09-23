#!/usr/bin/env node
import { parseArgs, assertLookupArgs, HELP, UsageError } from "./args.js";
import { AUDIT_GUIDANCE, explainAuditError, searchCalendarCreates } from "./audit.js";
import { loadEnvFile, readConfig } from "./env.js";
import { hydrateEvent, listEvents, mapPool } from "./events.js";
import { formatReport } from "./format.js";
import { explainCalendarAccessError, GraphApiError, GraphClient } from "./graph.js";
import { createTokenProvider } from "./auth.js";
import { buildEventReport, buildReport, eventInWindow } from "./report.js";
import { subjectMatches } from "./subject.js";
import type { AuditHit, AuditSection, GraphEvent } from "./types.js";
import { auditSearchWindow, defaultCalendarWindow, parseBoundary } from "./window.js";

async function main(): Promise<void> {
  const args = parseArgs(process.argv.slice(2));
  if (args.help) {
    process.stdout.write(HELP);
    return;
  }
  assertLookupArgs(args);
  loadEnvFile();
  const config = readConfig();
  const auditEnabled = args.audit ?? config.auditFallback;
  const defaults = defaultCalendarWindow();
  const window = {
    start: args.start ? parseBoundary(args.start, "start") : defaults.start,
    end: args.end ? parseBoundary(args.end, "end") : defaults.end,
  };
  if (Date.parse(window.start) >= Date.parse(window.end)) {
    throw new UsageError("--start must be earlier than --end.");
  }

  const graph = new GraphClient({
    baseUrl: config.graphBaseUrl,
    getToken: createTokenProvider(config),
  });

  let listed;
  try {
    listed = await listEvents(graph, {
      graphBaseUrl: config.graphBaseUrl,
      calendarEmail: args.calendar,
      start: window.start,
      end: window.end,
      top: args.top,
      maxPages: args.maxPages,
      masters: args.masters,
      match: args.match,
      subject: args.subject,
    });
  } catch (error) {
    if (error instanceof GraphApiError) {
      throw new Error(explainCalendarAccessError(error, args.calendar));
    }
    throw error;
  }

  const matched = listed.events.filter(
    (event) =>
      subjectMatches(event.subject, args.subject, args.match) &&
      (listed.source === "calendarView" || eventInWindow(event, window)),
  );
  const toHydrate = matched.slice(0, args.maxHydrate);
  const hydrated = await mapPool(toHydrate, 3, async (event) => hydrateOne(graph, args.calendar, event));
  const skipped = matched.slice(args.maxHydrate).map((event) => ({
    event,
    error: `Skipped extended-property lookup because the match count exceeded --max-hydrate ${args.maxHydrate}.`,
  }));
  const rows = [...hydrated, ...skipped];

  const unresolved = rows.some((row) => {
    const preview = buildEventReport(row.event, args.calendar, null, row.error);
    return preview.inferredCreator.confidence === "low" || preview.inferredCreator.confidence === "none";
  });
  const audit = await loadAudit(graph, {
    enabled: auditEnabled,
    calendarEmail: args.calendar,
    keyword: args.auditKeyword,
    window,
    hasMatches: rows.length > 0,
    unresolved,
  });

  const report = buildReport({
    calendar: args.calendar,
    subjectQuery: args.subject,
    match: args.match,
    source: listed.source,
    window,
    truncated: listed.truncated,
    pagesFetched: listed.pagesFetched,
    scannedEventCount: listed.events.length,
    hydratedCount: hydrated.filter((row) => !row.error).length,
    audit: publicAudit(audit),
    events: rows.map((row) =>
      buildEventReport(row.event, args.calendar, audit.hits, row.error),
    ),
  });

  if (args.json) {
    process.stdout.write(`${JSON.stringify(report, null, 2)}\n`);
  } else {
    process.stdout.write(`${formatReport(report)}\n`);
  }
  if (audit.status === "error") {
    process.exitCode = 2;
  }
}

interface HydratedRow {
  event: GraphEvent;
  error: string | null;
}

async function hydrateOne(
  graph: GraphClient,
  calendarEmail: string,
  event: GraphEvent,
): Promise<HydratedRow> {
  if (!event.id) {
    return { event, error: "Event has no id, so creator properties were not loaded." };
  }
  try {
    const full = await hydrateEvent(graph, calendarEmail, event.id);
    return { event: { ...event, ...full }, error: null };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return { event, error: message };
  }
}

interface AuditLoad extends AuditSection {
  hits: AuditHit[] | null;
}

async function loadAudit(
  graph: GraphClient,
  input: {
    enabled: boolean;
    calendarEmail: string;
    keyword?: string;
    window: { start: string; end: string };
    hasMatches: boolean;
    unresolved: boolean;
  },
): Promise<AuditLoad> {
  if (!input.enabled) {
    return {
      status: "not_requested",
      guidance: input.unresolved
        ? AUDIT_GUIDANCE
        : "Not requested. Pass --audit to compare these events with Exchange Create audit records.",
      queryId: null,
      window: null,
      windowNote: null,
      keyword: null,
      scannedRecordCount: null,
      error: null,
      hits: null,
    };
  }
  if (!input.hasMatches) {
    return {
      status: "skipped",
      guidance: "No calendar events matched, so no audit query was sent.",
      queryId: null,
      window: null,
      windowNote: null,
      keyword: null,
      scannedRecordCount: null,
      error: null,
      hits: null,
    };
  }
  const auditWindow = auditSearchWindow(input.window);
  if (auditWindow.empty) {
    return {
      status: "skipped",
      guidance: auditWindow.note,
      queryId: null,
      window: { start: auditWindow.start, end: auditWindow.end },
      windowNote: auditWindow.note,
      keyword: input.keyword ?? input.calendarEmail,
      scannedRecordCount: null,
      error: null,
      hits: null,
    };
  }
  try {
    const result = await searchCalendarCreates(graph, {
      calendarEmail: input.calendarEmail,
      keyword: input.keyword,
      window: auditWindow,
    });
    return {
      status: "ok",
      guidance: null,
      queryId: result.queryId,
      window: { start: auditWindow.start, end: auditWindow.end },
      windowNote: auditWindow.note,
      keyword: input.keyword ?? input.calendarEmail,
      scannedRecordCount: result.scannedRecordCount,
      error: null,
      hits: result.hits,
    };
  } catch (error) {
    return {
      status: "error",
      guidance: null,
      queryId: null,
      window: { start: auditWindow.start, end: auditWindow.end },
      windowNote: auditWindow.note,
      keyword: input.keyword ?? input.calendarEmail,
      scannedRecordCount: null,
      error: explainAuditError(error),
      hits: null,
    };
  }
}

function publicAudit(audit: AuditLoad): AuditSection {
  return {
    status: audit.status,
    guidance: audit.guidance,
    queryId: audit.queryId,
    window: audit.window,
    windowNote: audit.windowNote,
    keyword: audit.keyword,
    scannedRecordCount: audit.scannedRecordCount,
    error: audit.error,
  };
}

main().catch((error: unknown) => {
  const message = error instanceof Error ? error.message : String(error);
  process.stderr.write(`${message}\n`);
  if (error instanceof UsageError) {
    process.stderr.write("Run with --help for usage.\n");
  } else if (process.env.DEBUG && error instanceof Error && error.stack) {
    process.stderr.write(`${error.stack}\n`);
  }
  process.exitCode = 1;
});
