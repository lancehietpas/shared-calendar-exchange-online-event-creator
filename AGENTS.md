# Notes for agents

The supported way to find who created a shared-calendar event is the PowerShell 7 script `Find-SharedCalendarEventCreator.ps1`. Lance runs it locally on his Mac:

```bash
pwsh ./Find-SharedCalendarEventCreator.ps1 -Calendar shared-calendar@yourtenant.org -Subject "staff meeting"
```

Do not ask for `CLIENT_SECRET`, and do not put a real secret in the repo, a PR, or a chat reply. Secrets stay in his environment or a local `.env` (gitignored). `.env.example` lists the names only.

`src/` is an optional TypeScript mirror of the same Graph calls. It is not the user-facing tool. Prefer changing the PowerShell script when behavior changes, and keep the two from drifting if you touch both.

## Auth model

App-only client credentials. There is no signed-in user and no delegated token.

- Token: `POST https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token`
- `grant_type=client_credentials`
- `scope` is `{graph origin}/.default` (commercial default `https://graph.microsoft.com/.default`)
- `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET`

`Calendars.Read.Shared` is a delegated permission. It does not apply here. Application permission for reading mailbox calendars is `Calendars.Read`. `Calendars.ReadBasic` can omit extensions, so this lookup expects `Calendars.Read` because it reads legacy extended properties.

Tenant-wide `Calendars.Read` can be narrowed with an Exchange application access policy. Exchange RBAC for Applications does not narrow an Entra application permission that was already consented. See the README.

## What Graph returns

Default query:

`GET /users/{calendar}/calendar/calendarView?startDateTime&endDateTime`

`-Masters` uses `GET /users/{calendar}/calendar/events` instead. calendarView is requested without `$select` because that parameter drops `createdDateTime`.

The event resource has **no** `createdBy` or `lastModifiedBy`. On a shared calendar, `organizer` is often the shared mailbox even when a delegate created the item.

Creator evidence, in order:

1. `singleValueExtendedProperties`, especially MAPI `PidTagCreatorName` (`String 0x3FF8`) and `PidTagLastModifierName` (`String 0x3FFA`). Graph only returns extended properties when `$expand` names their ids.
2. Sender SMTP properties `String 0x5D01` and `String 0x5D02` when they differ from the mailbox.
3. `organizer`, when that address is not the shared mailbox.
4. Optional Purview audit search: `POST /security/auditLog/queries` with `recordTypeFilters: ["exchangeItem"]` and `operationFilters: ["Create"]`. Off unless `-Audit` or `AUDIT_FALLBACK=true`. Needs application permission `AuditLogsQuery-Exchange.Read.All`.

Subject `Contains` is case-insensitive and applied in the script. Graph does not support `contains()` on `event.subject`. `Equals` sends `subject eq '...'` and checks again locally. Audit rows are tied to an event only when the audit item subject equals that event subject.

The Office 365 Management Activity API (`ActivityFeed.Read`, `Audit.Exchange` blobs) is not called. It is a subscription feed, not a subject search.

## Checks

```bash
pwsh -File ./tests/Find-SharedCalendarEventCreator.Tests.ps1
```

PowerShell tests do not call Microsoft. Do not put tenant secrets in tests or fixtures.
