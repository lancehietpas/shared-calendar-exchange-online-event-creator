import { graphScope, type AppConfig } from "./env.js";

interface TokenResponse {
  access_token?: string;
  expires_in?: number;
  error?: string;
  error_description?: string;
}

export function createTokenProvider(
  config: AppConfig,
  fetchImpl: typeof fetch = fetch,
): () => Promise<string> {
  let cached: { token: string; expiresAt: number } | null = null;
  return async () => {
    if (cached && cached.expiresAt > Date.now() + 60_000) {
      return cached.token;
    }
    const next = await requestToken(config, fetchImpl);
    cached = next;
    return next.token;
  };
}

export async function requestToken(
  config: AppConfig,
  fetchImpl: typeof fetch = fetch,
): Promise<{ token: string; expiresAt: number }> {
  const url = `${config.authorityHost}/${encodeURIComponent(config.tenantId)}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: config.clientId,
    client_secret: config.clientSecret,
    scope: graphScope(config.graphBaseUrl),
    grant_type: "client_credentials",
  });

  let response: Response;
  try {
    response = await fetchImpl(url, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body,
      signal: AbortSignal.timeout(30_000),
    });
  } catch (error) {
    throw new Error(`Token request failed: ${errorMessage(error)}`);
  }

  const text = await response.text();
  const parsed = parseJson(text);
  if (!response.ok || !parsed?.access_token) {
    const description = redact(
      parsed?.error_description || parsed?.error || text.slice(0, 400),
      config.clientSecret,
    );
    throw new Error(
      `Token request failed (${response.status}). Check TENANT_ID, CLIENT_ID, and CLIENT_SECRET. ${description}`,
    );
  }

  const expiresIn = typeof parsed.expires_in === "number" ? parsed.expires_in : 3600;
  return {
    token: parsed.access_token,
    expiresAt: Date.now() + expiresIn * 1000,
  };
}

function parseJson(text: string): TokenResponse | null {
  if (!text) {
    return null;
  }
  try {
    return JSON.parse(text) as TokenResponse;
  } catch {
    return null;
  }
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function redact(text: string, secret: string): string {
  if (!secret) {
    return text;
  }
  return text.split(secret).join("[redacted]");
}
