export class GraphApiError extends Error {
  readonly status: number;
  readonly code: string | undefined;

  constructor(message: string, status: number, code?: string) {
    super(message);
    this.name = "GraphApiError";
    this.status = status;
    this.code = code;
  }
}

export class GraphClient {
  private readonly baseUrl: string;
  private readonly getToken: () => Promise<string>;
  private readonly fetchImpl: typeof fetch;

  constructor(options: {
    baseUrl: string;
    getToken: () => Promise<string>;
    fetchImpl?: typeof fetch;
  }) {
    this.baseUrl = options.baseUrl.replace(/\/$/, "");
    this.getToken = options.getToken;
    this.fetchImpl = options.fetchImpl ?? fetch;
  }

  getJson<T>(pathOrUrl: string): Promise<T> {
    return this.requestJson<T>("GET", pathOrUrl);
  }

  postJson<T>(path: string, body: unknown): Promise<T> {
    return this.requestJson<T>("POST", path, body);
  }

  private async requestJson<T>(method: string, pathOrUrl: string, body?: unknown): Promise<T> {
    const url = this.resolve(pathOrUrl);
    let attempt = 0;
    while (true) {
      attempt += 1;
      const token = await this.getToken();
      const headers: Record<string, string> = {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
        Prefer: 'outlook.timezone="UTC"',
      };
      if (body !== undefined) {
        headers["Content-Type"] = "application/json";
      }

      let response: Response;
      try {
        response = await this.fetchImpl(url, {
          method,
          headers,
          body: body === undefined ? undefined : JSON.stringify(body),
          signal: AbortSignal.timeout(30_000),
        });
      } catch (error) {
        throw new GraphApiError(
          `Graph request failed: ${error instanceof Error ? error.message : String(error)}`,
          0,
        );
      }

      if ((response.status === 429 || response.status === 503) && attempt < 4) {
        await delay(retryDelayMs(response, attempt));
        continue;
      }

      const text = await response.text();
      const parsed = parseJson(text);
      if (!response.ok) {
        const errorBody = asRecord(asRecord(parsed)?.error);
        const code = typeof errorBody?.code === "string" ? errorBody.code : undefined;
        const message =
          typeof errorBody?.message === "string" ? errorBody.message : text.slice(0, 500);
        throw new GraphApiError(
          `Graph ${response.status}${code ? ` ${code}` : ""}: ${message}`.trim(),
          response.status,
          code,
        );
      }
      return (parsed ?? {}) as T;
    }
  }

  private resolve(pathOrUrl: string): string {
    if (pathOrUrl.startsWith("https://")) {
      const target = new URL(pathOrUrl);
      const allowed = new URL(this.baseUrl);
      if (target.origin !== allowed.origin) {
        throw new GraphApiError(`Refusing to send the Graph token to ${target.origin}.`, 0);
      }
      return pathOrUrl;
    }
    const path = pathOrUrl.startsWith("/") ? pathOrUrl : `/${pathOrUrl}`;
    return `${this.baseUrl}${path}`;
  }
}

function retryDelayMs(response: Response, attempt: number): number {
  const retryAfter = Number(response.headers.get("retry-after"));
  if (Number.isFinite(retryAfter) && retryAfter > 0) {
    return Math.min(retryAfter * 1000, 10_000);
  }
  return Math.min(1000 * attempt, 10_000);
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function parseJson(text: string): unknown {
  if (!text) {
    return null;
  }
  try {
    return JSON.parse(text) as unknown;
  } catch {
    return null;
  }
}

function asRecord(value: unknown): Record<string, unknown> | null {
  if (value && typeof value === "object" && !Array.isArray(value)) {
    return value as Record<string, unknown>;
  }
  return null;
}

export function explainCalendarAccessError(error: GraphApiError, calendarEmail: string): string {
  if (error.status === 401) {
    return `${error.message} The app token was rejected. Confirm admin consent was granted and the secret is current.`;
  }
  if (error.status === 403) {
    return [
      error.message,
      `Application Calendars.Read is required, with admin consent.`,
      `If an Exchange application access policy is in place, ${calendarEmail} must be in the policy group.`,
      `Check with: Test-ApplicationAccessPolicy -Identity ${calendarEmail} -AppId <client-id>`,
    ].join(" ");
  }
  if (error.status === 404) {
    return `${error.message} The mailbox or calendar was not found. Use the shared mailbox primary SMTP address for --calendar.`;
  }
  return error.message;
}
