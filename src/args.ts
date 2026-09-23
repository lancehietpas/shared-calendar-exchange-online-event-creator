import type { MatchMode } from "./types.js";

export class UsageError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "UsageError";
  }
}

export interface CliArgs {
  help: boolean;
  calendar?: string;
  subject?: string;
  match: MatchMode;
  start?: string;
  end?: string;
  json: boolean;
  audit?: boolean;
  masters: boolean;
  top: number;
  maxPages: number;
  maxHydrate: number;
  auditKeyword?: string;
}

export const HELP = `Find who created an event on a shared Exchange Online calendar.

Usage:
  npm start -- --calendar <mailbox> --subject <text>
  node dist/cli.js --calendar <mailbox> --subject <text>

Required:
  --calendar <email>     Shared calendar mailbox SMTP address
  --subject <text>       Event subject to find

Matching:
  --match contains       Case-insensitive substring (default). Applied locally.
  --match equals         Case-insensitive full subject match

Window:
  --start <date>         YYYY-MM-DD or ISO-8601. Default: 90 days ago (UTC)
  --end <date>           YYYY-MM-DD or ISO-8601. Default: 180 days ahead (UTC)
  --masters              Read /calendar/events (series masters and single
                         instances) instead of calendarView occurrences

Audit:
  --audit                Search Purview for Exchange Create records
  --no-audit             Do not search audit logs
  --audit-keyword <text> Keyword sent to the audit query. Default: mailbox SMTP

Output:
  --json                 Print the JSON report on stdout
  --top <n>              Page size, 1-100 (default 50)
  --max-pages <n>        Maximum list pages, 1-100 (default 20)
  --max-hydrate <n>      Maximum events to load creator properties for (default 25)
  --help                 Show this help

Auth comes from TENANT_ID, CLIENT_ID, and CLIENT_SECRET. See .env.example.
`;

export function parseArgs(argv: string[]): CliArgs {
  const args: CliArgs = {
    help: false,
    match: "contains",
    json: false,
    masters: false,
    top: 50,
    maxPages: 20,
    maxHydrate: 25,
  };

  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index] ?? "";
    if (token === "--") {
      continue;
    }
    const { name, inline } = splitFlag(token);
    const take = (label: string): string => {
      if (inline !== undefined) {
        return inline;
      }
      const value = argv[index + 1];
      if (!value || value.startsWith("--")) {
        throw new UsageError(`Missing value for ${label}.`);
      }
      index += 1;
      return value;
    };

    switch (name) {
      case "--help":
      case "-h":
        args.help = true;
        break;
      case "--json":
        args.json = true;
        break;
      case "--audit":
        args.audit = true;
        break;
      case "--no-audit":
        args.audit = false;
        break;
      case "--masters":
        args.masters = true;
        break;
      case "--calendar":
        args.calendar = take("--calendar");
        break;
      case "--subject":
        args.subject = take("--subject");
        break;
      case "--match": {
        const mode = take("--match");
        if (mode !== "contains" && mode !== "equals") {
          throw new UsageError('--match must be "contains" or "equals".');
        }
        args.match = mode;
        break;
      }
      case "--start":
        args.start = take("--start");
        break;
      case "--end":
        args.end = take("--end");
        break;
      case "--audit-keyword":
        args.auditKeyword = take("--audit-keyword");
        break;
      case "--top":
        args.top = parseBounded(take("--top"), "--top", 1, 100);
        break;
      case "--max-pages":
        args.maxPages = parseBounded(take("--max-pages"), "--max-pages", 1, 100);
        break;
      case "--max-hydrate":
        args.maxHydrate = parseBounded(take("--max-hydrate"), "--max-hydrate", 1, 100);
        break;
      default:
        throw new UsageError(`Unknown argument "${token}".`);
    }
  }

  return args;
}

export function assertLookupArgs(args: CliArgs): asserts args is CliArgs & {
  calendar: string;
  subject: string;
} {
  if (!args.calendar?.trim() || !args.subject?.trim()) {
    throw new UsageError("Both --calendar and --subject are required.");
  }
  if (!args.calendar.includes("@")) {
    throw new UsageError("--calendar must be an email address.");
  }
  args.calendar = args.calendar.trim();
  args.subject = args.subject.trim();
}

function splitFlag(token: string): { name: string; inline?: string } {
  if (!token.startsWith("--") && token !== "-h") {
    throw new UsageError(`Unexpected argument "${token}". Flags use --name.`);
  }
  const separator = token.indexOf("=");
  if (separator === -1) {
    return { name: token };
  }
  return { name: token.slice(0, separator), inline: token.slice(separator + 1) };
}

function parseBounded(value: string, label: string, min: number, max: number): number {
  if (!/^\d+$/.test(value)) {
    throw new UsageError(`${label} must be an integer from ${min} to ${max}.`);
  }
  const parsed = Number(value);
  if (parsed < min || parsed > max) {
    throw new UsageError(`${label} must be an integer from ${min} to ${max}.`);
  }
  return parsed;
}
