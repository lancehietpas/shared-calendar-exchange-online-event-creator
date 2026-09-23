import { readFileSync, existsSync } from "node:fs";
import { resolve } from "node:path";

export function parseEnvFile(contents: string): Record<string, string> {
  const values: Record<string, string> = {};
  for (const rawLine of contents.split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line || line.startsWith("#")) {
      continue;
    }
    const withoutExport = line.startsWith("export ")
      ? line.slice("export ".length).trim()
      : line;
    const separator = withoutExport.indexOf("=");
    if (separator <= 0) {
      continue;
    }
    const key = withoutExport.slice(0, separator).trim();
    let value = withoutExport.slice(separator + 1).trim();
    if (
      value.length >= 2 &&
      ((value.startsWith('"') && value.endsWith('"')) ||
        (value.startsWith("'") && value.endsWith("'")))
    ) {
      value = value.slice(1, -1);
    }
    if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(key)) {
      values[key] = value;
    }
  }
  return values;
}

/** Apply a .env file without overriding variables already set in the environment. */
export function loadEnvFile(
  env: NodeJS.ProcessEnv = process.env,
  filePath = resolve(process.cwd(), ".env"),
): void {
  if (!existsSync(filePath)) {
    return;
  }
  const parsed = parseEnvFile(readFileSync(filePath, "utf8"));
  for (const [key, value] of Object.entries(parsed)) {
    if (env[key] === undefined || env[key] === "") {
      env[key] = value;
    }
  }
}

export interface AppConfig {
  tenantId: string;
  clientId: string;
  clientSecret: string;
  graphBaseUrl: string;
  authorityHost: string;
  auditFallback: boolean;
}

export function graphScope(graphBaseUrl: string): string {
  return `${new URL(graphBaseUrl).origin}/.default`;
}

export function readConfig(env: NodeJS.ProcessEnv = process.env): AppConfig {
  const required = ["TENANT_ID", "CLIENT_ID", "CLIENT_SECRET"] as const;
  const missing = required.filter((name) => !env[name]?.trim());
  if (missing.length > 0) {
    throw new Error(
      `Missing ${missing.join(", ")}. Copy .env.example to .env or export the variables. Do not commit the secret.`,
    );
  }

  const graphBaseUrl = (env.GRAPH_BASE_URL?.trim() || "https://graph.microsoft.com/v1.0").replace(
    /\/$/,
    "",
  );
  const authorityHost = (
    env.AUTHORITY_HOST?.trim() || "https://login.microsoftonline.com"
  ).replace(/\/$/, "");
  assertHttps("GRAPH_BASE_URL", graphBaseUrl);
  assertHttps("AUTHORITY_HOST", authorityHost);

  return {
    tenantId: env.TENANT_ID!.trim(),
    clientId: env.CLIENT_ID!.trim(),
    clientSecret: env.CLIENT_SECRET!.trim(),
    graphBaseUrl,
    authorityHost,
    auditFallback: parseBool(env.AUDIT_FALLBACK),
  };
}

function assertHttps(name: string, value: string): void {
  let url: URL;
  try {
    url = new URL(value);
  } catch {
    throw new Error(`${name} is not a valid URL.`);
  }
  if (url.protocol !== "https:") {
    throw new Error(`${name} must use https.`);
  }
}

function parseBool(value: string | undefined): boolean {
  if (!value) {
    return false;
  }
  return ["1", "true", "yes", "on"].includes(value.trim().toLowerCase());
}
