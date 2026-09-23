# Notes for agents

This CLI answers one question: who created an event on a shared Exchange Online calendar.

## Auth model

App-only client credentials. There is no signed-in user and no delegated token.

- Token: `POST https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token`
- `grant_type=client_credentials`
- `scope` is `{graph origin}/.default` (commercial default `https://graph.microsoft.com/.default`)
- Secrets are `TENANT_ID`, `CLIENT_ID`, and `CLIENT_SECRET`
- Load them from the environment or a local `.env`. Never commit `.env` or a real secret. `.env.example` is the template.

`Calendars.Read.Shared` is a delegated permission. It does not apply to this app. Application permission for reading mailbox calendars is `Calendars.Read` (or `Calendars.ReadWrite` if write is ever added). `Calendars.ReadBasic` is the least privileged list-events permission and can omit body, attachments, and extensions, so this tool expects `Calendars.Read` because it reads legacy extended properties.

Tenant-wide `Calendars.Read` can be narrowed with an Exchange application access policy. Exchange RBAC for Applications is a separate grant. It does not narrow an Entra application permission that was already consented. See the README.

## What Graph returns

Default query:

`GET /users/{calendar}/calendar/calendarView?startDateTime&endDateTime`

`--masters` uses `GET /users/{calendar}/calendar/events` instead.

The event resource has **no** `createdBy` or `lastModifiedBy`. On a shared calendar, `organizer` is often the shared mailbox even when a delegate created the item.

Creator evidence, in order:

1. `singleValueExtendedProperties`, especially MAPI `PidTagCreatorName` (`String 0x3FF8`) and `PidTagLastModifierName` (`String 0x3FFA`). Graph only returns extended properties when `$expand` names their ids.
2. Sender SMTP properties `String 0x5D01` and `String 0x5D02` when they differ from the mailbox.
3. `organizer`, when that address is not the shared mailbox.
4. Optional Purview audit search: `POST /security/auditLog/queries` with `recordTypeFilters: ["exchangeItem"]` and `operationFilters: ["Create"]`. Off unless `--audit` or `AUDIT_FALLBACK=true`. Needs application permission `AuditLogsQuery-Exchange.Read.All`.

Subject `contains` is case-insensitive and applied in this process. Graph does not support `contains()` on `event.subject`. `equals` sends `subject eq '...'` and checks again locally. Audit rows are tied to an event only when the audit item subject equals that event subject.

The Office 365 Management Activity API (`ActivityFeed.Read`, `Audit.Exchange` blobs) is not called. `CalendarCreateAuditLookup` in `src/audit.ts` is the seam if a blob-feed client is added later.

## Checks

```bash
npm install
npm test
npm start -- --help
```

Unit tests do not call Microsoft. Do not put tenant secrets in tests or fixtures.
