import { GraphApiError, GraphClient } from "./graph.js";
import { subjectMatches } from "./subject.js";
import type { AuditHit, RankedAudit } from "./types.js";
import type { AuditWindow } from "./window.js";

export const AUDIT_GUIDANCE = [
  "Mailbox audit fallback is off.",
  "Re-run with --audit (or AUDIT_FALLBACK=true) after an admin grants the application permission AuditLogsQuery-Exchange.Read.All and consents to it.",
  "Purview audits a calendar item Create for the mailbox owner, a delegate, and an admin when mailbox auditing is enabled.",
  "Shared-mailbox searches use the mailbox SMTP address as the keyword, not as the user filter.",
  "Audit (Standard) retention is 180 days.",
].join(" ");

export interface AuditSearchResult {
  queryId: string;
  hits: AuditHit[];
  scannedRecordCount: number;
}

/**
 * Extension point for "who created this calendar item" audit searches.
 * The CLI uses Microsoft Graph `POST /security/auditLog/queries`.
 * The Office 365 Management Activity API (Audit.Exchange content blobs) is
 * not implemented: it is a subscription feed, not a subject search.
 */
export interface CalendarCreateAuditLookup {
  searchCreates(input: {
    calendarEmail: string;
    window: AuditWindow;
    keyword?: string;
  }): Promise<AuditSearchResult>;
}

interface AuditQuery {
  id?: string;
  status?: string;
}

interface AuditRecordPage {
  value?: unknown[];
  "@odata.nextLink"?: string;
}

export async function searchCalendarCreates(
  graph: GraphClient,
  input: {
    calendarEmail: string;
    window: AuditWindow;
    keyword?: string;
    maxPages?: number;
    pollTimeoutMs?: number;
    sleep?: (ms: number) => Promise<void>;
  },
): Promise<AuditSearchResult> {
  const created = await graph.postJson<AuditQuery>("/security/auditLog/queries", {
    displayName: "shared-calendar-creator Create lookup",
    filterStartDateTime: input.window.start,
    filterEndDateTime: input.window.end,
    recordTypeFilters: ["exchangeItem"],
    operationFilters: ["Create"],
    keywordFilter: input.keyword ?? input.calendarEmail,
  });
  if (!created.id) {
    throw new GraphApiError("Audit query was created without an id.", 0);
  }

  const timeoutMs = input.pollTimeoutMs ?? 90_000;
  const sleep = input.sleep ?? delay;
  const deadline = Date.now() + timeoutMs;
  let status = created.status ?? "notStarted";
  while (!isTerminal(status)) {
    if (Date.now() >= deadline) {
      throw new GraphApiError(
        `Audit query ${created.id} was still ${status} after ${timeoutMs}ms.`,
        0,
      );
    }
    await sleep(2_000);
    const current = await graph.getJson<AuditQuery>(
      `/security/auditLog/queries/${encodeURIComponent(created.id)}`,
    );
    status = current.status ?? status;
  }
  if (status.toLowerCase() !== "succeeded") {
    throw new GraphApiError(`Audit query ${created.id} finished with status ${status}.`, 0);
  }

  const hits: AuditHit[] = [];
  let scanned = 0;
  let next: string | undefined = `/security/auditLog/queries/${encodeURIComponent(created.id)}/records`;
  let pages = 0;
  const maxPages = input.maxPages ?? 10;
  while (next && pages < maxPages) {
    const page: AuditRecordPage = await graph.getJson<AuditRecordPage>(next);
    for (const record of page.value ?? []) {
      scanned += 1;
      const hit = extractAuditHit(record);
      if ((hit.operation ?? "Create").toLowerCase() === "create") {
        hits.push(hit);
      }
    }
    next = page["@odata.nextLink"];
    pages += 1;
  }
  return { queryId: created.id, hits, scannedRecordCount: scanned };
}

export function explainAuditError(error: unknown): string {
  const message = error instanceof Error ? error.message : String(error);
  if (error instanceof GraphApiError && (error.status === 401 || error.status === 403)) {
    return [
      message,
      "Grant application permission AuditLogsQuery-Exchange.Read.All (admin consent).",
      "The Graph audit-log query API is documented at POST /security/auditLog/queries.",
      "Confirm mailbox auditing is enabled and Create is still audited for the shared mailbox.",
    ].join(" ");
  }
  return message;
}

export function extractAuditHit(record: unknown): AuditHit {
  const root = asRecord(record) ?? {};
  let data = pick(root, "auditData");
  if (typeof data === "string") {
    try {
      data = JSON.parse(data) as unknown;
    } catch {
      data = null;
    }
  }
  const auditData = asRecord(data) ?? {};
  const item = asRecord(pick(auditData, "Item", "item")) ?? {};
  const parent = asRecord(pick(item, "ParentFolder", "parentFolder")) ?? {};
  const folderPath =
    readString(pick(parent, "Path", "path")) ??
    readString(pick(auditData, "FolderPathName", "folderPathName"));
  const subject =
    readString(pick(item, "Subject", "subject")) ??
    readString(pick(auditData, "ItemSubject", "itemSubject"));

  return {
    id: readString(pick(root, "id")),
    createdDateTime:
      readString(pick(root, "createdDateTime")) ??
      readString(pick(auditData, "CreationTime", "creationTime")),
    operation: readString(pick(root, "operation")) ?? readString(pick(auditData, "Operation")),
    userId: readString(pick(root, "userId")) ?? readString(pick(auditData, "UserId")),
    userPrincipalName: readString(pick(root, "userPrincipalName")),
    logonType: formatLogonType(pick(auditData, "LogonType", "logonType")),
    logonUserDisplayName: readString(
      pick(auditData, "LogonUserDisplayName", "logonUserDisplayName"),
    ),
    mailboxOwnerUpn: readString(pick(auditData, "MailboxOwnerUPN", "mailboxOwnerUpn")),
    subject,
    folderPath,
    folderLooksLikeCalendar: folderPath ? /calendar/i.test(folderPath) : null,
    clientInfo: readString(pick(auditData, "ClientInfoString", "clientInfoString")),
    clientIp:
      readString(pick(root, "clientIp")) ?? readString(pick(auditData, "ClientIP", "clientIp")),
  };
}

export function rankAuditHits(
  hits: AuditHit[],
  event: { subject?: string | null; createdDateTime?: string | null },
  calendarEmail: string,
): RankedAudit {
  const matches = hits.filter((hit) => auditHitMatches(hit, event.subject, calendarEmail));
  const eventTime = Date.parse(event.createdDateTime ?? "");
  const ranked = [...matches].sort((left, right) => {
    const folderDelta = folderScore(right) - folderScore(left);
    if (folderDelta !== 0) {
      return folderDelta;
    }
    if (!Number.isNaN(eventTime)) {
      return distance(left, eventTime) - distance(right, eventTime);
    }
    return Date.parse(left.createdDateTime ?? "") - Date.parse(right.createdDateTime ?? "");
  });
  const distinctUsers = [
    ...new Set(
      ranked
        .map((hit) => hit.userId ?? hit.userPrincipalName)
        .filter((user): user is string => Boolean(user)),
    ),
  ];
  return {
    matches: ranked,
    best: ranked[0] ?? null,
    distinctUsers,
  };
}

function auditHitMatches(
  hit: AuditHit,
  eventSubject: string | null | undefined,
  calendarEmail: string,
): boolean {
  if (hit.operation && hit.operation.toLowerCase() !== "create") {
    return false;
  }
  if (
    hit.mailboxOwnerUpn &&
    hit.mailboxOwnerUpn.trim().toLocaleLowerCase() !== calendarEmail.trim().toLocaleLowerCase()
  ) {
    return false;
  }
  if (hit.folderLooksLikeCalendar === false) {
    return false;
  }
  if (!eventSubject?.trim() || !hit.subject?.trim()) {
    return false;
  }
  // Correlate on the full event subject. A contains search can match several
  // events; each event only keeps Create records for that exact subject.
  return subjectMatches(hit.subject, eventSubject, "equals");
}

function folderScore(hit: AuditHit): number {
  if (hit.folderLooksLikeCalendar === true) {
    return 1;
  }
  return 0;
}

function distance(hit: AuditHit, eventTime: number): number {
  const hitTime = Date.parse(hit.createdDateTime ?? "");
  if (Number.isNaN(hitTime)) {
    return Number.POSITIVE_INFINITY;
  }
  return Math.abs(hitTime - eventTime);
}

/**
 * Exchange mailbox audit LogonType values used by the unified audit log:
 * 0 Owner, 1 Admin, 2 Delegate. Strings are returned unchanged.
 */
export function formatLogonType(value: unknown): string | null {
  if (typeof value === "string" && value.trim()) {
    return value.trim();
  }
  if (typeof value === "number") {
    const names = ["Owner", "Admin", "Delegate"] as const;
    return names[value] ?? String(value);
  }
  return null;
}

function isTerminal(status: string): boolean {
  return ["succeeded", "failed", "cancelled"].includes(status.toLowerCase());
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function pick(record: Record<string, unknown>, ...names: string[]): unknown {
  const lookup = new Map(Object.keys(record).map((key) => [key.toLowerCase(), record[key]]));
  for (const name of names) {
    if (lookup.has(name.toLowerCase())) {
      return lookup.get(name.toLowerCase());
    }
  }
  return undefined;
}

function asRecord(value: unknown): Record<string, unknown> | null {
  if (value && typeof value === "object" && !Array.isArray(value)) {
    return value as Record<string, unknown>;
  }
  return null;
}

function readString(value: unknown): string | null {
  if (typeof value !== "string") {
    return null;
  }
  const trimmed = value.trim();
  return trimmed ? trimmed : null;
}
