import { GraphApiError, GraphClient } from "./graph.js";
import { FULL_PROPERTY_EXPAND, MINIMAL_PROPERTY_EXPAND } from "./properties.js";
import { subjectEqualsFilter } from "./subject.js";
import type { GraphCollection, GraphEvent, MatchMode } from "./types.js";

// calendarView ignores createdDateTime and lastModifiedDateTime when $select is
// present, so the calendarView request does not use $select. The events
// collection can use $select; creator timestamps are read again on hydrate.
const EVENT_SELECT = [
  "id",
  "subject",
  "start",
  "end",
  "organizer",
  "isOrganizer",
  "changeKey",
  "type",
  "iCalUId",
  "isCancelled",
  "webLink",
].join(",");

export interface ListEventsResult {
  events: GraphEvent[];
  source: "calendarView" | "events";
  pagesFetched: number;
  truncated: boolean;
  serverSubjectFilter: boolean;
}

export function buildListUrl(input: {
  graphBaseUrl: string;
  calendarEmail: string;
  start: string;
  end: string;
  top: number;
  masters: boolean;
  subjectEquals?: string;
  includeStartFilter?: boolean;
}): string {
  const root = input.graphBaseUrl.replace(/\/$/, "");
  const user = encodeURIComponent(input.calendarEmail);
  const path = input.masters ? "calendar/events" : "calendar/calendarView";
  const url = new URL(`${root}/users/${user}/${path}`);
  if (!input.masters) {
    url.searchParams.set("startDateTime", input.start);
    url.searchParams.set("endDateTime", input.end);
  } else {
    url.searchParams.set("$select", EVENT_SELECT);
  }
  url.searchParams.set("$top", String(input.top));
  if (input.subjectEquals) {
    url.searchParams.set("$filter", subjectEqualsFilter(input.subjectEquals));
  } else if (input.masters && input.includeStartFilter !== false) {
    url.searchParams.set("$filter", `start/dateTime ge ${odataDate(input.start)}`);
  }
  return url.toString();
}

export function buildHydratePath(calendarEmail: string, eventId: string, expand?: string): string {
  const url = new URL(
    `https://graph.microsoft.com/users/${encodeURIComponent(calendarEmail)}/events/${encodeURIComponent(eventId)}`,
  );
  if (expand) {
    // No $select here. calendarView drops createdDateTime under $select, and
    // $expand without $select still returns the default event properties.
    url.searchParams.set("$expand", expand);
  }
  return `${url.pathname}${url.search}`;
}

export async function listEvents(
  graph: GraphClient,
  input: {
    graphBaseUrl: string;
    calendarEmail: string;
    start: string;
    end: string;
    top: number;
    maxPages: number;
    masters: boolean;
    match: MatchMode;
    subject: string;
  },
): Promise<ListEventsResult> {
  const subjectEquals = input.match === "equals" ? input.subject : undefined;
  try {
    return await listPages(graph, { ...input, subjectEquals, includeStartFilter: true });
  } catch (error) {
    const canRetryWithoutFilter =
      error instanceof GraphApiError &&
      error.status === 400 &&
      (Boolean(subjectEquals) || input.masters);
    if (!canRetryWithoutFilter) {
      throw error;
    }
    return listPages(graph, {
      ...input,
      subjectEquals: undefined,
      includeStartFilter: false,
    });
  }
}

async function listPages(
  graph: GraphClient,
  input: {
    graphBaseUrl: string;
    calendarEmail: string;
    start: string;
    end: string;
    top: number;
    maxPages: number;
    masters: boolean;
    subjectEquals?: string;
    includeStartFilter: boolean;
  },
): Promise<ListEventsResult> {
  let url: string | undefined = buildListUrl(input);
  const events: GraphEvent[] = [];
  let pagesFetched = 0;
  while (url && pagesFetched < input.maxPages) {
    const page: GraphCollection<GraphEvent> = await graph.getJson<GraphCollection<GraphEvent>>(url);
    events.push(...(page.value ?? []));
    url = page["@odata.nextLink"];
    pagesFetched += 1;
  }
  return {
    events,
    source: input.masters ? "events" : "calendarView",
    pagesFetched,
    truncated: Boolean(url),
    serverSubjectFilter: Boolean(input.subjectEquals),
  };
}

export async function hydrateEvent(
  graph: GraphClient,
  calendarEmail: string,
  eventId: string,
): Promise<GraphEvent> {
  const attempts = [
    buildHydratePath(calendarEmail, eventId, FULL_PROPERTY_EXPAND),
    buildHydratePath(calendarEmail, eventId, MINIMAL_PROPERTY_EXPAND),
    buildHydratePath(calendarEmail, eventId),
  ];
  let lastError: unknown;
  for (const [index, path] of attempts.entries()) {
    try {
      return await graph.getJson<GraphEvent>(path);
    } catch (error) {
      lastError = error;
      const retryable = error instanceof GraphApiError && error.status === 400 && index < attempts.length - 1;
      if (!retryable) {
        throw error;
      }
    }
  }
  throw lastError instanceof Error ? lastError : new GraphApiError("Event lookup failed.", 0);
}

function odataDate(value: string): string {
  const withoutFraction = value.replace(/\.\d+/, "").replace(/Z$/, "");
  return `'${withoutFraction}'`;
}

export async function mapPool<T, R>(
  items: T[],
  limit: number,
  worker: (item: T, index: number) => Promise<R>,
): Promise<R[]> {
  const results = new Array<R>(items.length);
  let cursor = 0;
  async function run(): Promise<void> {
    while (cursor < items.length) {
      const index = cursor;
      cursor += 1;
      results[index] = await worker(items[index] as T, index);
    }
  }
  const workers = Math.max(1, Math.min(limit, items.length));
  await Promise.all(Array.from({ length: workers }, () => run()));
  return results;
}
