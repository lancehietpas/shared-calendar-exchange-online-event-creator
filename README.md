# Shared calendar event creator

Find who originally created an event on a shared Exchange Online calendar.

The tool signs in as an app (client credentials), reads the shared mailbox calendar with Microsoft Graph, and reports the organizer plus the MAPI properties that usually name the person who created the item. When those properties only show the shared mailbox, you can ask it to search the Exchange mailbox audit log for the `Create` operation.

```bash
npm install
npm run build
npm start -- --calendar shared-calendar@yourtenant.org --subject "staff meeting"
```

## What you will see

For each match: subject, Graph id, start, end, organizer name and email, `createdBy`, `lastModifiedBy`, and `singleValueExtendedProperties`. The Graph event resource does not include `createdBy` or `lastModifiedBy`. Those lines stay empty unless a response happens to contain them. The useful creator fields are the extended properties below.

| Property | Graph id | Meaning |
| --- | --- | --- |
| `PidTagCreatorName` | `String 0x3FF8` | Display name of the message creator ([MS-OXCMSG](https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcmsg/35135646-c41d-444f-b5de-a92062f18cd9)) |
| `PidTagLastModifierName` | `String 0x3FFA` | Display name of the last person to change the item |
| `PidTagSenderSmtpAddress` | `String 0x5D01` | SMTP address of the sending mailbox |
| `PidTagSentRepresentingSmtpAddress` | `String 0x5D02` | SMTP address of the represented sender |

On a shared calendar, Graph usually sets `organizer` to the **mailbox**, not the delegate who typed the appointment. That is expected. `PidTagCreatorName` is the first place a different person shows up. It is a display name, not an account id. `--audit` is how you get the account (`UserId`) from the mailbox audit log.

Example (values are illustrative):

```text
Calendar: shared-calendar@yourtenant.org
Subject query: staff
Match: contains (case-insensitive contains, applied in the CLI. Graph does not support contains() on event.subject.)
Source: /users/{calendar}/calendar/calendarView
Window: 2026-06-25T00:00:00.000Z .. 2027-03-22T00:00:00.000Z
Scanned events: 12 across 1 page(s)
Matches: 1
Audit: Not requested. Pass --audit to compare these events with Exchange Create audit records.

---
Subject: Staff meeting
Id: AAMkAGExample
Type: singleInstance
iCalUId: (none)
Start: 2026-09-23T15:00:00.0000000 (UTC)
End: 2026-09-23T16:00:00.0000000 (UTC)
Created: 2026-09-01T15:00:00Z
Last modified: 2026-09-01T15:05:00Z
Organizer: Community Calendar <shared-calendar@yourtenant.org>
createdBy: not returned by the Graph event resource
lastModifiedBy: not returned by the Graph event resource
singleValueExtendedProperties:
  PidTagCreatorName: Lance Smith
  PidTagLastModifierName: Lance Smith
Inferred creator: Lance Smith
  source: PidTagCreatorName
  confidence: medium
  organizer is calendar mailbox: yes
  note: PidTagCreatorName (MAPI 0x3FF8) differs from the shared mailbox. On a shared calendar this is usually the display name of the person who created the item. Graph does not provide that person's SMTP address here.
Best creator: Lance Smith
  source: PidTagCreatorName
  confidence: medium
  detail: PidTagCreatorName (MAPI 0x3FF8) differs from the shared mailbox. On a shared calendar this is usually the display name of the person who created the item. Graph does not provide that person's SMTP address here.
```

`--json` prints the same report as JSON on stdout. Warnings and errors go to stderr.

## Register the Entra app

Use the Entra admin center for the tenant that hosts the shared mailbox (work tenant or a personal Microsoft 365 tenant).

1. **Identity > Applications > App registrations > New registration.**
   - Name: `shared-calendar-creator`
   - Supported account types: **Accounts in this organizational directory only**
   - Redirect URI: leave empty. This is a background app.
2. Copy **Application (client) ID** into `CLIENT_ID` and **Directory (tenant) ID** into `TENANT_ID`.
3. **Certificates & secrets > New client secret.** Copy the secret **Value** into `CLIENT_SECRET` immediately. Store it in `.env` or your shell, not in git.
4. **API permissions > Add a permission > Microsoft Graph > Application permissions.**
   - Add **`Calendars.Read`**.
   - Admin consent it (**Grant admin consent**).
5. Add **`AuditLogsQuery-Exchange.Read.All`** only if you will use `--audit`. Consent that too.

### Which calendar permission

Checked against [List events](https://learn.microsoft.com/en-us/graph/api/user-list-events?view=graph-rest-1.0) and [List calendar events](https://learn.microsoft.com/en-us/graph/api/calendar-list-events?view=graph-rest-1.0):

| Permission | Application (app-only) | Use it? |
| --- | --- | --- |
| `Calendars.ReadBasic` | Yes. Least privileged for listing events. | Not enough here. Basic read can omit body, attachments, and extensions. This tool reads legacy extended properties. |
| `Calendars.Read` | Yes. | **Use this.** |
| `Calendars.ReadWrite` | Yes. | Only if something else must edit events. This CLI does not write. |
| `Calendars.Read.Shared` | **No.** Delegated only. | Does not apply to client credentials. |

Application permissions are tenant-wide until you scope them. Do that before you try the CLI against a production tenant.

### Limit the app to the shared mailbox

Two different mechanisms exist. They are not interchangeable.

**Application access policy** narrows Microsoft Graph application permissions that were consented in Entra. This is the right follow-up after you grant `Calendars.Read`.

[Application access policies](https://learn.microsoft.com/en-us/exchange/permissions-exo/application-access-policies) and [Limit mailbox access](https://learn.microsoft.com/en-us/graph/auth-limit-mailbox-access):

```powershell
Connect-ExchangeOnline
# Mail-enabled security group whose only member is the shared mailbox.
New-ApplicationAccessPolicy `
  -AppId "<client-id>" `
  -PolicyScopeGroupId "calendar-creator-scope@yourtenant.org" `
  -AccessRight RestrictAccess `
  -Description "Calendar creator lookup"
Test-ApplicationAccessPolicy -Identity "shared-calendar@yourtenant.org" -AppId "<client-id>"
```

`RestrictAccess` means the app may access mailboxes in that group and no others. The policy can take time to apply. `Test-ApplicationAccessPolicy` is the check. A mailbox outside the group returns HTTP 403 from Graph.

**Exchange RBAC for Applications** is Microsoft's newer model ([RBAC for Applications](https://learn.microsoft.com/en-us/exchange/permissions-exo/application-rbac)). It grants `Application Calendars.Read` to the app's service principal for a mailbox scope. Those grants are **added to** Entra permissions. They do not shrink a tenant-wide `Calendars.Read` consent. Use RBAC instead of the Entra calendar permission, not on top of an unscoped consent, if you want the scope to be the only grant:

```powershell
Connect-ExchangeOnline
# Object ID is the Enterprise application object id, not the app registration object id.
New-ServicePrincipal -AppId "<client-id>" -ObjectId "<enterprise-app-object-id>" -DisplayName "shared-calendar-creator"
New-ManagementScope -Name "SharedCalendarMailbox" -RecipientRestrictionFilter "PrimarySmtpAddress -eq 'shared-calendar@yourtenant.org'"
New-ManagementRoleAssignment -App "<enterprise-app-object-id>" -Role "Application Calendars.Read" -CustomResourceScope "SharedCalendarMailbox"
Test-ServicePrincipalAuthorization -Identity "<client-id>" -Resource "shared-calendar@yourtenant.org"
```

Outlook sharing ("share this calendar") is not required. App-only Graph reads the mailbox itself.

## Configure and run

Node.js 20 or newer.

```bash
cp .env.example .env
# edit .env — TENANT_ID, CLIENT_ID, CLIENT_SECRET
npm install
npm start -- --calendar shared-calendar@yourtenant.org --subject "staff meeting"
```

`.env` is gitignored. The process environment wins over `.env` when both set the same name.

| Variable | Required | Purpose |
| --- | --- | --- |
| `TENANT_ID` | yes | Directory (tenant) id |
| `CLIENT_ID` | yes | Application (client) id |
| `CLIENT_SECRET` | yes | Client secret value |
| `GRAPH_BASE_URL` | no | Default `https://graph.microsoft.com/v1.0` |
| `AUTHORITY_HOST` | no | Default `https://login.microsoftonline.com` |
| `AUDIT_FALLBACK` | no | `true` turns on the audit search when you do not pass `--audit` or `--no-audit` |

The token request is the [client credentials flow](https://learn.microsoft.com/en-us/graph/auth-v2-service). Scope is `https://graph.microsoft.com/.default` unless `GRAPH_BASE_URL` points at another Graph host, in which case the scope uses that host.

### Flags

```text
--calendar <email>       Shared mailbox SMTP address (primary SMTP is safest)
--subject <text>         Subject to find
--match contains         Default. Case-insensitive substring, applied locally.
--match equals           Case-insensitive full subject. Also sent as subject eq.
--start <date>           YYYY-MM-DD or ISO-8601. Default: 90 days ago (UTC)
--end <date>             YYYY-MM-DD or ISO-8601. Default: 180 days ahead (UTC)
--masters                Use /calendar/events instead of calendarView
--audit                  Search Purview for Exchange Create records
--no-audit               Do not search audit logs
--audit-keyword <text>   Audit keyword. Default: the mailbox address
--json                   JSON on stdout
--top <n>                Page size 1-100 (default 50)
--max-pages <n>          Stop after this many pages (default 20)
--max-hydrate <n>        How many matches to load extended properties for (default 25)
```

`npm start` rebuilds TypeScript first. After `npm run build` you can also run `node dist/cli.js`.

Exit codes: `0` query finished (including zero matches), `1` usage or Graph calendar failure, `2` calendar query finished but the audit search failed.

## How the lookup works

1. Get an app-only token.
2. Read the default calendar.
   - Default: [`calendarView`](https://learn.microsoft.com/en-us/graph/api/calendar-list-calendarview?view=graph-rest-1.0) for the date window. Recurring series come back as instances. That request does not use `$select`, because [calendarView omits `createdDateTime` and `lastModifiedDateTime` when `$select` is set](https://learn.microsoft.com/en-us/graph/api/user-list-calendarview?view=graph-rest-1.0).
   - `--masters`: [`/calendar/events`](https://learn.microsoft.com/en-us/graph/api/user-list-events?view=graph-rest-1.0), which returns single meetings and series masters, not expanded occurrences. Results are then limited to the window when `start.dateTime` parses.
3. Match the subject.
   - **`contains` (default):** case-insensitive substring in this program. [Graph's event examples filter with `startswith` / `eq`](https://learn.microsoft.com/en-us/graph/api/calendar-list-events?view=graph-rest-1.0). `contains()` on `event.subject` is not supported (Graph returns an unsupported filter operator). Every event in the window is scanned, up to `--max-pages`.
   - **`equals`:** case-insensitive equality. The request includes `$filter=subject eq '...'` (single quotes doubled, per OData). If Graph rejects that filter, the tool scans the window and compares locally. Graph string equality is case-insensitive; the local check is too.
4. For each match, `GET /users/{mailbox}/events/{id}` with `$expand=singleValueExtendedProperties(...)`. Graph returns an extended property only when the expand filter names its id ([extended properties](https://learn.microsoft.com/en-us/graph/api/resources/extended-properties-overview?view=graph-rest-1.0), [get singleValueLegacyExtendedProperty](https://learn.microsoft.com/en-us/graph/api/singlevaluelegacyextendedproperty-get?view=graph-rest-1.0)).
5. Decide a best creator:
   - Distinct `PidTagCreatorName`, plus a matching SMTP address when one is present.
   - Otherwise that display name alone (medium confidence).
   - Otherwise a sender SMTP address that is not the shared mailbox.
   - Otherwise `organizer`, when it is not the shared mailbox. That pattern is a meeting someone else organized, not an appointment typed on the shared calendar.
   - If the only name is the mailbox itself, confidence stays low and the audit path is the way to see a delegate.

`Prefer: outlook.timezone="UTC"` is sent so start and end come back in UTC.

## Audit fallback

Turn it on with `--audit` when the organizer and `PidTagCreatorName` are both the shared mailbox, or when you want the account that the audit log recorded.

Implemented call ([Create auditLogQuery](https://learn.microsoft.com/en-us/graph/api/security-auditcoreroot-post-auditlogqueries?view=graph-rest-1.0)):

```http
POST https://graph.microsoft.com/v1.0/security/auditLog/queries
Content-Type: application/json

{
  "displayName": "shared-calendar-creator Create lookup",
  "filterStartDateTime": "<window start>",
  "filterEndDateTime": "<window end>",
  "recordTypeFilters": ["exchangeItem"],
  "operationFilters": ["Create"],
  "keywordFilter": "<shared mailbox SMTP>"
}
```

The tool polls `GET /security/auditLog/queries/{id}` until `succeeded`, then reads `GET /security/auditLog/queries/{id}/records`. A hit is kept when:

- the operation is `Create`
- `MailboxOwnerUPN` is the `--calendar` address, when that field is present
- the folder path is missing or contains `Calendar` (a path such as `\Contacts` is dropped)
- the item subject equals the event subject, ignoring case

The closest audit time to the event `createdDateTime` is listed first. Graph event ids are not the Exchange audit item ids, so the tool does not join on id.

`keywordFilter` is the free-text keyword. For a shared mailbox, put the SMTP address in the keyword, not in the user filter ([Search mailbox activities](https://learn.microsoft.com/en-us/purview/audit-log-search-for-mailbox-activities)). `--audit-keyword` overrides it. The query is limited to the past 180 days, which is [Audit (Standard) retention](https://learn.microsoft.com/en-us/purview/audit-log-retention-policies). Future dates are clamped to now. If the tenant keeps logs longer, pass a `--start`/`--end` pair inside that retention.

### Permissions and auditing you need for `--audit`

1. Application permission **`AuditLogsQuery-Exchange.Read.All`**, admin consented. The [audit query permission table](https://learn.microsoft.com/en-us/graph/api/security-auditcoreroot-post-auditlogqueries?view=graph-rest-1.0) is split by Microsoft 365 service. Exchange mailbox records need the Exchange permission, not only `AuditLogsQuery-Entra.Read.All`.
2. Mailbox auditing enabled. In Exchange Online it is on by default. [Created mailbox item](https://learn.microsoft.com/en-us/purview/audit-log-activities) (`Create`) is written when an item is created in Calendar, Contacts, Notes, or Tasks. Message creation is not audited. Owner, delegate, and admin creates are covered when those audit sets still include `Create`.
3. Confirm with Exchange Online PowerShell:

```powershell
Get-OrganizationConfig | Format-List AuditDisabled
Get-Mailbox shared-calendar@yourtenant.org | Format-List AuditEnabled,DefaultAuditSet,AuditOwner,AuditDelegate,AuditAdmin
```

`AuditDisabled` should be `False`. `Create` should appear in the owner, delegate, and admin sets. If it was removed:

```powershell
Set-OrganizationConfig -AuditDisabled $false
Set-Mailbox shared-calendar@yourtenant.org -AuditEnabled $true
Set-Mailbox shared-calendar@yourtenant.org -AuditOwner @{Add="Create"} -AuditDelegate @{Add="Create"} -AuditAdmin @{Add="Create"}
```

Audit records are not retroactive. A create that happened while auditing was off will not appear.

### What is not implemented

The [Office 365 Management Activity API](https://learn.microsoft.com/en-us/office/office-365-management-api/office-365-management-activity-api-reference) can also deliver `Audit.Exchange` blobs (`POST .../activity/feed/subscriptions/start?contentType=Audit.Exchange`). That API needs the Office 365 Management APIs application permission **`ActivityFeed.Read`**, and it is a subscription feed rather than a subject search. This version does not call `manage.office.com`. `CalendarCreateAuditLookup` in `src/audit.ts` is the interface for that client if you add one later. The Graph audit-log query above is the search that `--audit` runs.

## Troubleshooting

| Symptom | What to check |
| --- | --- |
| Token error `invalid_client` | Secret value (not the secret id), tenant id, and client id. A new secret is required after expiry. |
| Graph 403 | Admin consent for `Calendars.Read`. If a policy exists, the mailbox must be in the scope group. `Test-ApplicationAccessPolicy`. |
| Graph 404 | `--calendar` is the mailbox's primary SMTP address, and the mailbox has a calendar. |
| Zero matches | Widen `--start`/`--end`, or switch `--match`. The default window is 90 days back and 180 days forward. |
| Organizer is only the shared mailbox | Expected. Read `PidTagCreatorName`. If that is also the mailbox, run `--audit`. |
| Audit 403 | `AuditLogsQuery-Exchange.Read.All` consented. |
| Audit query stays empty | Auditing disabled, `Create` removed from the mailbox audit set, the event is older than retention, or the keyword does not match. Search the same mailbox in Purview with the SMTP address in **Keywords**. |
| List stopped early | Raise `--max-pages`. The report says when the page cap was hit. |

## Development

```bash
npm test
```

Tests use Node's built-in test runner. They do not call Microsoft.

See `AGENTS.md` for the auth model and the Graph calls.
