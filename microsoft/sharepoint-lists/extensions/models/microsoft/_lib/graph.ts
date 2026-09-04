// Microsoft Graph API Helper
// Shared fetch utilities for the Microsoft models in this repo.
//
// Adapted from @webframp/microsoft/teams (Sean Escriva,
// https://github.com/webframp/swamp-extensions), licensed under the Apache
// License 2.0. Changed from the original: the paginated helper returns a
// truncated flag alongside the accumulated items, the page cap is a parameter
// rather than a module constant, and requests retry throttled and transient
// responses honouring Retry-After.

export const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

/** Shape of a Graph collection response. */
export interface GraphListResponse<T> {
  value: T[];
  "@odata.nextLink"?: string;
  "@odata.count"?: number;
}

/** A non-2xx Graph response, carrying the HTTP status and Graph error code. */
export class GraphApiError extends Error {
  constructor(
    public readonly statusCode: number,
    public readonly graphCode: string,
    message: string,
  ) {
    super(message);
    this.name = "GraphApiError";
  }
}

// ---------------------------------------------------------------------------
// Retry policy
// ---------------------------------------------------------------------------

/**
 * Statuses worth retrying. 429 is Graph's throttle; SharePoint list endpoints
 * throttle aggressively, and a long paginated walk is exactly the shape that
 * trips it. 503/504 are transient backend faults that Graph's own guidance
 * says to retry.
 */
const RETRYABLE_STATUS = new Set([429, 503, 504]);

/** How hard to retry a throttled or transient response. */
export interface RetryOptions {
  /** Retries after the first attempt. Default 3; 0 disables retrying. */
  maxRetries?: number;
  /** Ceiling on any single wait, in milliseconds. Default 30_000. */
  maxDelayMs?: number;
  /** Injectable for tests, so a retry path does not really sleep. */
  sleepFn?: (ms: number) => Promise<void>;
}

const defaultSleep = (ms: number) =>
  new Promise<void>((resolve) => setTimeout(resolve, ms));

/**
 * How long to wait before retrying. `Retry-After` wins when present — Graph
 * sends it in seconds, but the HTTP spec also permits a date, so both are
 * handled. Without it, back off exponentially from one second.
 */
export function retryDelayMs(
  response: Response,
  attempt: number,
  maxDelayMs: number,
): number {
  const header = response.headers.get("Retry-After");

  if (header !== null) {
    const seconds = Number(header);
    if (Number.isFinite(seconds) && seconds >= 0) {
      return Math.min(seconds * 1000, maxDelayMs);
    }
    const until = Date.parse(header);
    if (!Number.isNaN(until)) {
      return Math.min(Math.max(until - Date.now(), 0), maxDelayMs);
    }
  }

  return Math.min(1000 * 2 ** attempt, maxDelayMs);
}

// ---------------------------------------------------------------------------
// Single-resource request
// ---------------------------------------------------------------------------

/**
 * Make a single Graph API request and return the parsed response body.
 * Throws GraphApiError on non-2xx responses. A 204 (returned by DELETE and
 * some PATCH calls) resolves to an empty object.
 */
export async function graphRequest<T>(
  accessToken: string,
  method: string,
  path: string,
  body?: unknown,
  extraHeaders?: Record<string, string>,
  fetchFn: typeof fetch = fetch,
  retry: RetryOptions = {},
): Promise<T> {
  const url = path.startsWith("https://") ? path : `${GRAPH_BASE}${path}`;
  const headers: Record<string, string> = {
    "Authorization": `Bearer ${accessToken}`,
    "Content-Type": "application/json",
    ...extraHeaders,
  };

  const maxRetries = retry.maxRetries ?? 3;
  const maxDelayMs = retry.maxDelayMs ?? 30_000;
  const sleepFn = retry.sleepFn ?? defaultSleep;

  for (let attempt = 0;; attempt++) {
    const response = await fetchFn(url, {
      method,
      headers,
      body: body !== undefined ? JSON.stringify(body) : undefined,
    });

    if (RETRYABLE_STATUS.has(response.status) && attempt < maxRetries) {
      const delayMs = retryDelayMs(response, attempt, maxDelayMs);
      // Drain the body before discarding the response so the connection is
      // released rather than leaked for the lifetime of the walk.
      await response.text().catch(() => {});
      await sleepFn(delayMs);
      continue;
    }

    if (response.status === 204) {
      // No content — return empty object.
      return {} as T;
    }

    // Parse only after the status is known to be final, and tolerate a body
    // that is not JSON: an error from a proxy or gateway is often HTML, and
    // reporting the status beats a SyntaxError from the parser.
    const raw = await response.text();
    let data: Record<string, unknown>;
    try {
      data = (raw ? JSON.parse(raw) : {}) as Record<string, unknown>;
    } catch {
      if (!response.ok) {
        throw new GraphApiError(
          response.status,
          "non_json_response",
          `Graph API error ${response.status}: ${
            raw.slice(0, 200) || response.statusText
          }`,
        );
      }
      throw new GraphApiError(
        response.status,
        "non_json_response",
        `Graph returned a non-JSON body for ${method} ${path}`,
      );
    }

    if (!response.ok) {
      const err = data["error"] as
        | { code?: string; message?: string }
        | undefined;
      throw new GraphApiError(
        response.status,
        String(err?.code ?? "unknown"),
        String(err?.message ?? `Graph API error ${response.status}`),
      );
    }

    return data as T;
  }
}

// ---------------------------------------------------------------------------
// Paginated list request
// ---------------------------------------------------------------------------

/** Accumulated pages plus whether the page cap cut the walk short. */
export interface PaginatedResult<T> {
  items: T[];
  truncated: boolean;
}

/**
 * Fetch pages of a Graph API list endpoint, following @odata.nextLink up to
 * maxPages (default 20). Returns items and whether results were truncated.
 */
export async function graphRequestPaginated<T>(
  accessToken: string,
  path: string,
  params?: Record<string, string>,
  extraHeaders?: Record<string, string>,
  fetchFn: typeof fetch = fetch,
  maxPages: number = 20,
  retry: RetryOptions = {},
): Promise<PaginatedResult<T>> {
  const allItems: T[] = [];

  let url: string;
  if (path.startsWith("https://")) {
    url = path;
    if (params && Object.keys(params).length > 0) {
      const qs = new URLSearchParams(params).toString();
      url = `${url}${url.includes("?") ? "&" : "?"}${qs}`;
    }
  } else {
    const qs = params ? `?${new URLSearchParams(params).toString()}` : "";
    url = `${GRAPH_BASE}${path}${qs}`;
  }

  let pages = 0;
  while (pages < maxPages) {
    const page = await graphRequest<GraphListResponse<T>>(
      accessToken,
      "GET",
      url,
      undefined,
      extraHeaders,
      fetchFn,
      retry,
    );

    allItems.push(...(page.value ?? []));
    pages++;

    if (!page["@odata.nextLink"]) {
      return { items: allItems, truncated: false };
    }

    url = page["@odata.nextLink"];
  }

  return { items: allItems, truncated: true };
}
