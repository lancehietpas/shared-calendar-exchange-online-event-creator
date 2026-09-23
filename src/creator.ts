import type {
  AuditHit,
  BestCreator,
  Confidence,
  InferredCreator,
  PersonRef,
} from "./types.js";

export function sameText(left: string | null | undefined, right: string | null | undefined): boolean {
  const a = (left ?? "").trim().toLocaleLowerCase();
  const b = (right ?? "").trim().toLocaleLowerCase();
  return a.length > 0 && a === b;
}

export function inferCreator(input: {
  calendarEmail: string;
  organizerName?: string | null;
  organizerEmail?: string | null;
  extended: Record<string, string>;
}): InferredCreator {
  const calendarEmail = input.calendarEmail.trim();
  const organizerName = blankToNull(input.organizerName);
  const organizerEmail = blankToNull(input.organizerEmail);
  const organizerIsCalendarMailbox = organizerEmail
    ? sameText(organizerEmail, calendarEmail)
    : null;
  const creatorName = blankToNull(input.extended.PidTagCreatorName);
  const senderSmtp =
    firstDifferent(
      calendarEmail,
      input.extended.PidTagSenderSmtpAddress,
      input.extended.PidTagSentRepresentingSmtpAddress,
    ) ?? null;
  const senderName =
    blankToNull(input.extended.PidTagSenderName) ??
    blankToNull(input.extended.PidTagSentRepresentingName);

  const creatorIsMailbox =
    creatorName !== null &&
    (sameText(creatorName, calendarEmail) ||
      (organizerIsCalendarMailbox === true && sameText(creatorName, organizerName)));

  if (creatorName && !creatorIsMailbox) {
    const emailFromOrganizer =
      sameText(creatorName, organizerName) && organizerIsCalendarMailbox === false
        ? organizerEmail
        : null;
    const emailFromSender =
      senderSmtp && (!senderName || sameText(senderName, creatorName)) ? senderSmtp : null;
    const email = emailFromOrganizer ?? emailFromSender;
    return {
      name: creatorName,
      email,
      source: "PidTagCreatorName",
      confidence: email ? "high" : "medium",
      organizerIsCalendarMailbox,
      note: email
        ? "PidTagCreatorName (MAPI 0x3FF8) names someone other than the shared mailbox, and an SMTP address on the item matches that person."
        : "PidTagCreatorName (MAPI 0x3FF8) differs from the shared mailbox. On a shared calendar this is usually the display name of the person who created the item. Graph does not provide that person's SMTP address here.",
    };
  }

  if (senderSmtp) {
    return {
      name: senderName,
      email: senderSmtp,
      source: "PidTagSenderSmtpAddress",
      confidence: "medium",
      organizerIsCalendarMailbox,
      note: "PidTagSenderSmtpAddress or PidTagSentRepresentingSmtpAddress differs from the calendar mailbox. For a meeting this is often the sending mailbox, which may be the person who created the item.",
    };
  }

  if (organizerEmail && organizerIsCalendarMailbox === false) {
    return {
      name: organizerName,
      email: organizerEmail,
      source: "organizer",
      confidence: "medium",
      organizerIsCalendarMailbox,
      note: "The organizer is not the shared calendar mailbox. This looks like a meeting organized by that person rather than an appointment created directly on the shared calendar.",
    };
  }

  if (creatorName) {
    return {
      name: creatorName,
      email: organizerIsCalendarMailbox ? organizerEmail : null,
      source: "PidTagCreatorName",
      confidence: "low",
      organizerIsCalendarMailbox,
      note: "PidTagCreatorName matches the shared mailbox. Exchange often stores the mailbox as both organizer and creator when someone works in that calendar. A mailbox audit Create record can show whether a delegate did it.",
    };
  }

  return {
    name: null,
    email: null,
    source: "unresolved",
    confidence: "none",
    organizerIsCalendarMailbox,
    note: organizerIsCalendarMailbox
      ? "Graph returned the shared mailbox as organizer and no distinct creator property. The event resource has no createdBy field. Run again with --audit to search the Exchange mailbox audit log."
      : "Graph did not return PidTagCreatorName or a distinct organizer. The event resource has no createdBy field. Run again with --audit to search the Exchange mailbox audit log.",
  };
}

export function chooseBestCreator(inferred: InferredCreator, audit: AuditHit[]): BestCreator {
  if (inferred.confidence === "high" || inferred.confidence === "medium") {
    return {
      name: inferred.name,
      email: inferred.email,
      source: inferred.source,
      confidence: inferred.confidence,
      detail: inferred.note,
    };
  }

  const users = distinctAuditUsers(audit);
  if (users.length === 1) {
    const hit = audit.find((item) => auditUser(item) === users[0]) ?? audit[0];
    return {
      name: hit?.logonUserDisplayName ?? null,
      email: users[0] ?? null,
      source: "mailbox audit Create",
      confidence: audit.length > 1 ? "low" : "medium",
      detail:
        audit.length > 1
          ? `Several Exchange Create audit records match this subject. They all name ${users[0]}.`
          : "Taken from the Exchange mailbox audit Create record. UserId is the account that performed the create.",
    };
  }
  if (users.length > 1) {
    return {
      name: null,
      email: null,
      source: "mailbox audit Create",
      confidence: "low",
      detail: `Several accounts created matching items in this window: ${users.join(", ")}. Narrow --start/--end or use --match equals.`,
    };
  }

  if (inferred.confidence === "low") {
    return {
      name: inferred.name,
      email: inferred.email,
      source: inferred.source,
      confidence: "low",
      detail: inferred.note,
    };
  }

  return {
    name: null,
    email: null,
    source: "unresolved",
    confidence: "none",
    detail: inferred.note,
  };
}

export function readActor(value: unknown): PersonRef | null {
  if (!value || typeof value !== "object") {
    return null;
  }
  const record = value as Record<string, unknown>;
  const user = asRecord(record.user) ?? record;
  const email =
    readString(user.email) ??
    readString(user.address) ??
    readString(asRecord(user.emailAddress)?.address);
  const name =
    readString(user.displayName) ??
    readString(user.name) ??
    readString(asRecord(user.emailAddress)?.name);
  if (!email && !name) {
    return null;
  }
  return { name: name ?? null, email: email ?? null };
}

function firstDifferent(calendarEmail: string, ...values: Array<string | undefined>): string | null {
  for (const value of values) {
    const trimmed = blankToNull(value);
    if (trimmed && !sameText(trimmed, calendarEmail)) {
      return trimmed;
    }
  }
  return null;
}

function distinctAuditUsers(hits: AuditHit[]): string[] {
  const users = new Set<string>();
  for (const hit of hits) {
    const user = auditUser(hit);
    if (user) {
      users.add(user);
    }
  }
  return [...users];
}

function auditUser(hit: AuditHit): string | null {
  return blankToNull(hit.userId) ?? blankToNull(hit.userPrincipalName);
}

function blankToNull(value: string | null | undefined): string | null {
  const trimmed = value?.trim();
  return trimmed ? trimmed : null;
}

function readString(value: unknown): string | null {
  return typeof value === "string" ? blankToNull(value) : null;
}

function asRecord(value: unknown): Record<string, unknown> | null {
  if (value && typeof value === "object" && !Array.isArray(value)) {
    return value as Record<string, unknown>;
  }
  return null;
}
