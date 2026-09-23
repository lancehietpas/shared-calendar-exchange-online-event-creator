export type MatchMode = "contains" | "equals";

export type Confidence = "high" | "medium" | "low" | "none";

export interface GraphDateTimeZone {
  dateTime?: string;
  timeZone?: string;
}

export interface EmailAddress {
  name?: string | null;
  address?: string | null;
}

export interface GraphEvent {
  id?: string;
  subject?: string | null;
  start?: GraphDateTimeZone;
  end?: GraphDateTimeZone;
  organizer?: { emailAddress?: EmailAddress };
  isOrganizer?: boolean;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  changeKey?: string;
  type?: string;
  iCalUId?: string;
  isCancelled?: boolean;
  webLink?: string;
  createdBy?: unknown;
  lastModifiedBy?: unknown;
  singleValueExtendedProperties?: SingleValueExtendedProperty[];
}

export interface SingleValueExtendedProperty {
  id?: string;
  value?: string;
}

export interface GraphCollection<T> {
  value?: T[];
  "@odata.nextLink"?: string;
}

export interface PersonRef {
  name: string | null;
  email: string | null;
}

export interface ExtendedPropertyValue {
  id: string;
  canonicalName: string | null;
  value: string;
}

export interface InferredCreator extends PersonRef {
  source:
    | "PidTagCreatorName"
    | "PidTagSenderSmtpAddress"
    | "organizer"
    | "unresolved";
  confidence: Confidence;
  organizerIsCalendarMailbox: boolean | null;
  note: string;
}

export interface AuditHit {
  id: string | null;
  createdDateTime: string | null;
  operation: string | null;
  userId: string | null;
  userPrincipalName: string | null;
  logonType: string | null;
  logonUserDisplayName: string | null;
  mailboxOwnerUpn: string | null;
  subject: string | null;
  folderPath: string | null;
  folderLooksLikeCalendar: boolean | null;
  clientInfo: string | null;
  clientIp: string | null;
}

export interface RankedAudit {
  matches: AuditHit[];
  best: AuditHit | null;
  distinctUsers: string[];
}

export interface BestCreator extends PersonRef {
  source: string;
  confidence: Confidence;
  detail: string;
}

export interface EventReport {
  id: string | null;
  subject: string | null;
  start: GraphDateTimeZone | null;
  end: GraphDateTimeZone | null;
  type: string | null;
  iCalUId: string | null;
  isCancelled: boolean | null;
  createdDateTime: string | null;
  lastModifiedDateTime: string | null;
  organizer: PersonRef;
  isOrganizer: boolean | null;
  createdBy: PersonRef | null;
  lastModifiedBy: PersonRef | null;
  singleValueExtendedProperties: ExtendedPropertyValue[];
  extendedProperties: Record<string, string>;
  inferredCreator: InferredCreator;
  audit: RankedAudit | null;
  bestCreator: BestCreator;
  hydrateError: string | null;
}

export interface AuditSection {
  status: "not_requested" | "skipped" | "ok" | "error";
  guidance: string | null;
  queryId: string | null;
  window: { start: string; end: string } | null;
  windowNote: string | null;
  keyword: string | null;
  scannedRecordCount: number | null;
  error: string | null;
}

export interface LookupReport {
  calendar: string;
  subjectQuery: string;
  match: MatchMode;
  matchDescription: string;
  source: "calendarView" | "events";
  window: { start: string; end: string };
  truncated: boolean;
  pagesFetched: number;
  scannedEventCount: number;
  matchCount: number;
  hydratedCount: number;
  audit: AuditSection;
  events: EventReport[];
}
