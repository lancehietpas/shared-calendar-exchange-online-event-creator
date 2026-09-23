import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { parseArgs, UsageError } from "./args.js";
import { extractAuditHit, formatLogonType, rankAuditHits } from "./audit.js";
import { requestToken } from "./auth.js";
import { chooseBestCreator, inferCreator } from "./creator.js";
import { graphScope, parseEnvFile, readConfig } from "./env.js";
import { buildHydratePath, buildListUrl, listEvents } from "./events.js";
import { formatReport } from "./format.js";
import { GraphClient } from "./graph.js";
import { FULL_PROPERTY_EXPAND, readExtendedProperties } from "./properties.js";
import { buildEventReport } from "./report.js";
import { odataStringLiteral, subjectEqualsFilter, subjectMatches } from "./subject.js";
import { auditSearchWindow, parseBoundary, parseGraphDateTime } from "./window.js";

const calendar = "shared-calendar@example.com";

describe("subject matching", () => {
  it("matches contains without regard to case", () => {
    assert.equal(subjectMatches("Weekly Staff Meeting", "staff", "contains"), true);
    assert.equal(subjectMatches("Weekly Staff Meeting", "budget", "contains"), false);
    assert.equal(subjectMatches(null, "staff", "contains"), false);
  });

  it("matches equals on the full subject only", () => {
    assert.equal(subjectMatches("Staff Meeting", "staff meeting", "equals"), true);
    assert.equal(subjectMatches("Weekly Staff Meeting", "staff", "equals"), false);
  });

  it("escapes single quotes in OData literals", () => {
    assert.equal(odataStringLiteral("O'Brien"), "'O''Brien'");
    assert.equal(subjectEqualsFilter("O'Brien"), "subject eq 'O''Brien'");
  });
});

describe("creator inference", () => {
  it("uses PidTagCreatorName when it differs from the shared mailbox", () => {
    const inferred = inferCreator({
      calendarEmail: calendar,
      organizerName: "Community Calendar",
      organizerEmail: calendar,
      extended: { PidTagCreatorName: "Lance Smith" },
    });
    assert.equal(inferred.source, "PidTagCreatorName");
    assert.equal(inferred.name, "Lance Smith");
    assert.equal(inferred.email, null);
    assert.equal(inferred.confidence, "medium");
    assert.equal(inferred.organizerIsCalendarMailbox, true);
  });

  it("attaches SMTP when the sender address matches the creator name", () => {
    const inferred = inferCreator({
      calendarEmail: calendar,
      organizerName: "Community Calendar",
      organizerEmail: calendar,
      extended: {
        PidTagCreatorName: "Lance Smith",
        PidTagSenderName: "Lance Smith",
        PidTagSenderSmtpAddress: "lance@example.com",
      },
    });
    assert.equal(inferred.confidence, "high");
    assert.equal(inferred.email, "lance@example.com");
  });

  it("reports the organizer when the meeting is not owned by the shared mailbox", () => {
    const inferred = inferCreator({
      calendarEmail: calendar,
      organizerName: "Ada Lovelace",
      organizerEmail: "ada@example.com",
      extended: {},
    });
    assert.equal(inferred.source, "organizer");
    assert.equal(inferred.confidence, "medium");
    assert.equal(inferred.organizerIsCalendarMailbox, false);
  });

  it("stays low confidence when the creator name is the mailbox", () => {
    const inferred = inferCreator({
      calendarEmail: calendar,
      organizerName: "Community Calendar",
      organizerEmail: calendar,
      extended: { PidTagCreatorName: "Community Calendar" },
    });
    assert.equal(inferred.confidence, "low");
    const best = chooseBestCreator(inferred, [
      {
        id: "1",
        createdDateTime: "2026-09-01T15:00:00Z",
        operation: "Create",
        userId: "delegate@example.com",
        userPrincipalName: null,
        logonType: "Delegate",
        logonUserDisplayName: "Dee Delegate",
        mailboxOwnerUpn: calendar,
        subject: "Staff meeting",
        folderPath: "\\Calendar",
        folderLooksLikeCalendar: true,
        clientInfo: null,
        clientIp: null,
      },
    ]);
    assert.equal(best.source, "mailbox audit Create");
    assert.equal(best.email, "delegate@example.com");
    assert.equal(best.confidence, "medium");
  });

  it("does not invent a creator when Graph has no signal and audit is empty", () => {
    const inferred = inferCreator({
      calendarEmail: calendar,
      organizerName: "Community Calendar",
      organizerEmail: calendar,
      extended: {},
    });
    assert.equal(inferred.confidence, "none");
    assert.equal(inferred.source, "unresolved");
    const best = chooseBestCreator(inferred, []);
    assert.equal(best.confidence, "none");
  });
});

describe("audit records", () => {
  it("reads Exchange mailbox fields from auditData", () => {
    const hit = extractAuditHit({
      id: "rec-1",
      createdDateTime: "2026-09-01T15:00:01Z",
      operation: "Create",
      userId: "delegate@example.com",
      auditData: {
        LogonType: 2,
        MailboxOwnerUPN: calendar,
        ClientInfoString: "Client=OWA",
        Item: { Subject: "Staff meeting", ParentFolder: { Path: "\\Calendar" } },
      },
    });
    assert.equal(hit.userId, "delegate@example.com");
    assert.equal(hit.logonType, "Delegate");
    assert.equal(hit.subject, "Staff meeting");
    assert.equal(hit.folderLooksLikeCalendar, true);
    assert.equal(formatLogonType(0), "Owner");
    assert.equal(formatLogonType(1), "Admin");
  });

  it("parses auditData when Graph returns it as a JSON string", () => {
    const hit = extractAuditHit({
      operation: "Create",
      auditData: JSON.stringify({
        UserId: "ada@example.com",
        Item: { Subject: "Budget" },
      }),
    });
    assert.equal(hit.userId, "ada@example.com");
    assert.equal(hit.subject, "Budget");
  });

  it("correlates Create records by exact subject and mailbox", () => {
    const hits = [
      extractAuditHit({
        operation: "Create",
        userId: "ada@example.com",
        createdDateTime: "2026-09-01T15:00:00Z",
        auditData: {
          MailboxOwnerUPN: calendar,
          Item: { Subject: "Staff meeting", ParentFolder: { Path: "\\Calendar" } },
        },
      }),
      extractAuditHit({
        operation: "Create",
        userId: "other@example.com",
        auditData: {
          MailboxOwnerUPN: "someone-else@example.com",
          Item: { Subject: "Staff meeting", ParentFolder: { Path: "\\Calendar" } },
        },
      }),
      extractAuditHit({
        operation: "Create",
        userId: "pat@example.com",
        auditData: {
          MailboxOwnerUPN: calendar,
          Item: { Subject: "Staff party", ParentFolder: { Path: "\\Calendar" } },
        },
      }),
      extractAuditHit({
        operation: "Update",
        userId: "pat@example.com",
        auditData: {
          MailboxOwnerUPN: calendar,
          Item: { Subject: "Staff meeting", ParentFolder: { Path: "\\Calendar" } },
        },
      }),
    ];
    const ranked = rankAuditHits(
      hits,
      { subject: "Staff meeting", createdDateTime: "2026-09-01T15:00:02Z" },
      calendar,
    );
    assert.equal(ranked.matches.length, 1);
    assert.equal(ranked.best?.userId, "ada@example.com");
  });
});

describe("event queries", () => {
  it("builds a calendarView URL and keeps contains matching local", () => {
    const url = new URL(
      buildListUrl({
        graphBaseUrl: "https://graph.microsoft.com/v1.0",
        calendarEmail: calendar,
        start: "2026-01-01T00:00:00.000Z",
        end: "2026-02-01T00:00:00.000Z",
        top: 50,
        masters: false,
      }),
    );
    assert.equal(url.pathname, `/v1.0/users/${encodeURIComponent(calendar)}/calendar/calendarView`);
    assert.equal(url.searchParams.get("startDateTime"), "2026-01-01T00:00:00.000Z");
    assert.equal(url.searchParams.get("endDateTime"), "2026-02-01T00:00:00.000Z");
    assert.equal(url.searchParams.get("$filter"), null);
    assert.equal(url.searchParams.get("$select"), null);
  });

  it("asks Graph for subject eq in equals mode", () => {
    const url = new URL(
      buildListUrl({
        graphBaseUrl: "https://graph.microsoft.com/v1.0",
        calendarEmail: calendar,
        start: "2026-01-01T00:00:00.000Z",
        end: "2026-02-01T00:00:00.000Z",
        top: 25,
        masters: false,
        subjectEquals: "O'Brien review",
      }),
    );
    assert.equal(url.searchParams.get("$filter"), "subject eq 'O''Brien review'");
  });

  it("requests PidTagCreatorName when hydrating an event", () => {
    const path = buildHydratePath(calendar, "event id/with slash", FULL_PROPERTY_EXPAND);
    assert.match(path, /singleValueExtendedProperties/);
    assert.match(path.replace(/\+/g, " "), /String 0x3FF8/);
    assert.match(path, /event%20id%2Fwith%20slash/);
  });
});

describe("dates and config", () => {
  it("parses Graph UTC dateTime that has no offset", () => {
    const parsed = parseGraphDateTime({
      dateTime: "2026-09-23T15:00:00.0000000",
      timeZone: "UTC",
    });
    assert.equal(parsed, Date.parse("2026-09-23T15:00:00.000Z"));
  });

  it("parses date-only boundaries as UTC days", () => {
    assert.equal(parseBoundary("2026-09-01", "start"), "2026-09-01T00:00:00.000Z");
    assert.equal(parseBoundary("2026-09-01", "end"), "2026-09-01T23:59:59.000Z");
  });

  it("clamps audit search to the past and to 180 days", () => {
    const now = new Date("2026-09-23T00:00:00.000Z");
    const window = auditSearchWindow(
      { start: "2020-01-01T00:00:00.000Z", end: "2027-01-01T00:00:00.000Z" },
      now,
    );
    assert.equal(window.truncatedFuture, true);
    assert.equal(window.truncatedRetention, true);
    assert.equal(window.end, now.toISOString());
    assert.equal(window.empty, false);
  });

  it("does not override existing environment values", () => {
    const parsed = parseEnvFile(`
# comment
export TENANT_ID="tenant-from-file"
CLIENT_SECRET='secret'
NOT A LINE
`);
    assert.equal(parsed.TENANT_ID, "tenant-from-file");
    assert.equal(parsed.CLIENT_SECRET, "secret");
    assert.equal(parsed["NOT A LINE"], undefined);
  });

  it("rejects missing secrets and non-https endpoints", () => {
    assert.throws(() => readConfig({}), /TENANT_ID/);
    assert.throws(
      () =>
        readConfig({
          TENANT_ID: "t",
          CLIENT_ID: "c",
          CLIENT_SECRET: "s",
          GRAPH_BASE_URL: "http://graph.example",
        }),
      /https/,
    );
    assert.equal(graphScope("https://graph.microsoft.com/v1.0"), "https://graph.microsoft.com/.default");
  });
});

describe("cli args", () => {
  it("accepts inline flags", () => {
    const args = parseArgs([
      "--calendar=shared-calendar@example.com",
      "--subject",
      "Staff meeting",
      "--match",
      "equals",
      "--json",
      "--audit",
    ]);
    assert.equal(args.calendar, "shared-calendar@example.com");
    assert.equal(args.subject, "Staff meeting");
    assert.equal(args.match, "equals");
    assert.equal(args.json, true);
    assert.equal(args.audit, true);
  });

  it("rejects unknown flags and bad match modes", () => {
    assert.throws(() => parseArgs(["--nope"]), UsageError);
    assert.throws(() => parseArgs(["--match", "regex"]), UsageError);
  });
});

describe("graph client", () => {
  it("redacts the client secret from token errors", async () => {
    const fetchImpl: typeof fetch = async (_url, init) => {
      const body = init?.body instanceof URLSearchParams ? init.body : new URLSearchParams();
      assert.equal(body.get("grant_type"), "client_credentials");
      assert.equal(body.get("scope"), "https://graph.microsoft.com/.default");
      return new Response(
        JSON.stringify({
          error: "invalid_client",
          error_description: "rejected super-secret-value",
        }),
        { status: 401 },
      );
    };
    await assert.rejects(
      () =>
        requestToken(
          {
            tenantId: "tenant",
            clientId: "client",
            clientSecret: "super-secret-value",
            graphBaseUrl: "https://graph.microsoft.com/v1.0",
            authorityHost: "https://login.microsoftonline.com",
            auditFallback: false,
          },
          fetchImpl,
        ),
      (error: unknown) => {
        assert.match(error instanceof Error ? error.message : "", /\[redacted\]/);
        assert.doesNotMatch(error instanceof Error ? error.message : "", /super-secret-value/);
        return true;
      },
    );
  });

  it("refuses to follow a nextLink on another host", async () => {
    let calls = 0;
    const fetchImpl: typeof fetch = async () => {
      calls += 1;
      return new Response(
        JSON.stringify({
          value: [{ id: "1", subject: "Staff meeting" }],
          "@odata.nextLink": "https://evil.example/next",
        }),
        { status: 200, headers: { "Content-Type": "application/json" } },
      );
    };
    const graph = new GraphClient({
      baseUrl: "https://graph.microsoft.com/v1.0",
      getToken: async () => "token",
      fetchImpl,
    });
    await assert.rejects(
      () =>
        listEvents(graph, {
          graphBaseUrl: "https://graph.microsoft.com/v1.0",
          calendarEmail: calendar,
          start: "2026-01-01T00:00:00.000Z",
          end: "2026-02-01T00:00:00.000Z",
          top: 10,
          maxPages: 5,
          masters: false,
          match: "contains",
          subject: "staff",
        }),
      /Refusing to send the Graph token/,
    );
    assert.equal(calls, 1);
  });
});

describe("report text", () => {
  it("shows organizer, missing createdBy, and extended properties", () => {
    const extended = readExtendedProperties([
      { id: "String 0x3ff8", value: "Lance Smith" },
      { id: "String 0x3FFA", value: "Lance Smith" },
    ]);
    assert.equal(extended.byName.PidTagCreatorName, "Lance Smith");
    const event = buildEventReport(
      {
        id: "AAMkAGExample",
        subject: "Staff meeting",
        start: { dateTime: "2026-09-23T15:00:00.0000000", timeZone: "UTC" },
        end: { dateTime: "2026-09-23T16:00:00.0000000", timeZone: "UTC" },
        organizer: { emailAddress: { name: "Community Calendar", address: calendar } },
        createdDateTime: "2026-09-01T15:00:00Z",
        lastModifiedDateTime: "2026-09-01T15:05:00Z",
        type: "singleInstance",
        singleValueExtendedProperties: [
          { id: "String 0x3ff8", value: "Lance Smith" },
          { id: "String 0x3FFA", value: "Lance Smith" },
        ],
      },
      calendar,
      null,
      null,
    );
    const text = formatReport({
      calendar,
      subjectQuery: "staff",
      match: "contains",
      matchDescription: "case-insensitive contains",
      source: "calendarView",
      window: { start: "2026-06-25T00:00:00.000Z", end: "2027-03-22T00:00:00.000Z" },
      truncated: false,
      pagesFetched: 1,
      scannedEventCount: 4,
      matchCount: 1,
      hydratedCount: 1,
      audit: {
        status: "not_requested",
        guidance: "Not requested.",
        queryId: null,
        window: null,
        windowNote: null,
        keyword: null,
        scannedRecordCount: null,
        error: null,
      },
      events: [event],
    });
    assert.match(text, /Subject: Staff meeting/);
    assert.match(text, /Organizer: Community Calendar <shared-calendar@example.com>/);
    assert.match(text, /createdBy: not returned by the Graph event resource/);
    assert.match(text, /PidTagCreatorName: Lance Smith/);
    assert.match(text, /Best creator: Lance Smith/);
  });
});
