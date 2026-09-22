// Microsoft Graph Authentication Helper
// Delegated OAuth2 with device code flow (public client) and silent token refresh.
//
// Adapted from @webframp/microsoft/teams (Sean Escriva,
// https://github.com/webframp/swamp-extensions), licensed under the Apache
// License 2.0. That license continues to govern the portions derived from it;
// the Apache-2.0 notice above is retained per its section 4(b).
//
// Changed from the original: scopes are a parameter rather than a module
// constant, because a refresh token is bound to the scopes it was issued with,
// and a tenant that will consent to Sites.Read.All but not Sites.ReadWrite.All
// needs to bootstrap with a narrower set.

const TOKEN_ENDPOINT_BASE = "https://login.microsoftonline.com";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

/** Stored credentials for one public-client app registration. */
export interface MicrosoftCredentials {
  tenantId: string;
  clientId: string;
  refreshToken: string;
  scopes: string;
}

/** Token endpoint success payload. */
export interface TokenResponse {
  access_token: string;
  refresh_token: string;
  expires_in: number;
  token_type: string;
  scope: string;
}

/** Device code endpoint success payload. */
export interface DeviceCodeResponse {
  device_code: string;
  user_code: string;
  verification_uri: string;
  expires_in: number;
  interval: number;
  message: string;
}

/**
 * Default delegated scopes for reading and writing SharePoint list items.
 *
 * `Sites.ReadWrite.All` is delegated, not application: every call still runs as
 * the signed-in user, so the app can never reach a list that user cannot
 * already open, and item-level permissions and list validation apply normally.
 * That is the main reason this model uses device code flow rather than the
 * app-only client credentials the other SharePoint extensions use — an
 * app-only `Sites.ReadWrite.All` is tenant-wide write to every site.
 *
 * For a read-only deployment, override with
 * `offline_access User.Read Sites.Read.All` and the write methods will fail at
 * Graph with `accessDenied` rather than silently succeeding.
 */
export const DEFAULT_SHAREPOINT_SCOPES = [
  "offline_access",
  "User.Read",
  "Sites.ReadWrite.All",
].join(" ");

// ---------------------------------------------------------------------------
// Response parsing
// ---------------------------------------------------------------------------

/**
 * Read a token-endpoint response body as JSON, tolerating a body that is not
 * JSON at all.
 *
 * A proxy or gateway in front of login.microsoftonline.com answers with HTML,
 * and `response.json()` on that throws a SyntaxError that hides the status
 * that actually explains the failure. Surfacing the status as a
 * MicrosoftAuthError keeps a 502 from being reported as a parser fault.
 * `_lib/graph.ts` handles Graph responses the same way.
 */
async function readTokenBody(
  response: Response,
  operation: string,
): Promise<Record<string, unknown>> {
  const raw = await response.text();
  try {
    return (raw ? JSON.parse(raw) : {}) as Record<string, unknown>;
  } catch {
    throw new MicrosoftAuthError(
      "non_json_response",
      `${operation} returned a non-JSON body (HTTP ${response.status} ${
        response.statusText || "no status text"
      }): ${raw.slice(0, 200)}`,
    );
  }
}

// ---------------------------------------------------------------------------
// Token refresh (public client — no client_secret)
// ---------------------------------------------------------------------------

/**
 * Exchange a refresh token for a new access token.
 * Public client flow: no client_secret is sent.
 * Throws MicrosoftAuthError with code "invalid_grant" when the refresh token
 * has expired (90-day inactivity or password change) — callers should direct
 * the user to re-run `bootstrap`.
 */
export async function refreshAccessToken(
  creds: MicrosoftCredentials,
  fetchFn: typeof fetch = fetch,
): Promise<TokenResponse> {
  const url = `${TOKEN_ENDPOINT_BASE}/${creds.tenantId}/oauth2/v2.0/token`;

  const body = new URLSearchParams({
    grant_type: "refresh_token",
    client_id: creds.clientId,
    refresh_token: creds.refreshToken,
    scope: creds.scopes,
  });

  const response = await fetchFn(url, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString(),
  });

  const data = await readTokenBody(response, "Token refresh");

  if (!response.ok || data["error"]) {
    const errorCode = String(data["error"] ?? "unknown");
    const errorDesc = String(
      data["error_description"] ?? "Token refresh failed",
    );

    if (errorCode === "invalid_grant") {
      throw new MicrosoftAuthError(
        "invalid_grant",
        "Refresh token has expired or been revoked. " +
          "Re-run the `bootstrap` method to re-authenticate via device code flow.",
      );
    }

    throw new MicrosoftAuthError(errorCode, errorDesc);
  }

  return data as unknown as TokenResponse;
}

// ---------------------------------------------------------------------------
// Device code flow (public client — no client_secret)
// ---------------------------------------------------------------------------

/**
 * Initiate a device code flow and return the DeviceCodeResponse.
 * The caller should display `response.message` to the user, then poll
 * `pollDeviceCode()` until a token is returned.
 */
export async function initiateDeviceCode(
  tenantId: string,
  clientId: string,
  scopes: string,
  fetchFn: typeof fetch = fetch,
): Promise<DeviceCodeResponse> {
  const url = `${TOKEN_ENDPOINT_BASE}/${tenantId}/oauth2/v2.0/devicecode`;

  const body = new URLSearchParams({ client_id: clientId, scope: scopes });

  const response = await fetchFn(url, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString(),
  });

  const data = await readTokenBody(response, "Device code initiation");

  if (!response.ok) {
    throw new MicrosoftAuthError(
      String(data["error"] ?? "device_code_error"),
      String(
        data["error_description"] ?? "Failed to initiate device code flow",
      ),
    );
  }

  return data as unknown as DeviceCodeResponse;
}

/**
 * Poll the token endpoint until the user completes device code authentication.
 * Public client: no client_secret in the poll request.
 * Returns the token once granted, or throws on permanent failure.
 */
export async function pollDeviceCode(
  tenantId: string,
  clientId: string,
  deviceCode: string,
  intervalSeconds: number,
  timeoutMs: number,
  fetchFn: typeof fetch = fetch,
  sleepFn: (ms: number) => Promise<void> = (ms) =>
    new Promise((r) => setTimeout(r, ms)),
): Promise<TokenResponse> {
  const url = `${TOKEN_ENDPOINT_BASE}/${tenantId}/oauth2/v2.0/token`;
  const deadline = Date.now() + timeoutMs;

  while (Date.now() < deadline) {
    await sleepFn(intervalSeconds * 1000);

    const body = new URLSearchParams({
      grant_type: "urn:ietf:params:oauth:grant-type:device_code",
      client_id: clientId,
      device_code: deviceCode,
    });

    const response = await fetchFn(url, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString(),
    });

    const data = await readTokenBody(response, "Device code polling");

    if (response.ok && data["access_token"]) {
      return data as unknown as TokenResponse;
    }

    const errorCode = String(data["error"] ?? "");

    if (errorCode === "authorization_pending") {
      continue;
    }

    if (errorCode === "slow_down") {
      intervalSeconds += 5;
      continue;
    }

    throw new MicrosoftAuthError(
      errorCode,
      String(
        data["error_description"] ?? "Device code authentication failed",
      ),
    );
  }

  throw new MicrosoftAuthError(
    "device_code_expired",
    "Device code authentication timed out. Re-run `bootstrap` to try again.",
  );
}

// ---------------------------------------------------------------------------
// Error type
// ---------------------------------------------------------------------------

/** An OAuth2 failure, carrying Entra's own error code. */
export class MicrosoftAuthError extends Error {
  constructor(
    public readonly code: string,
    message: string,
  ) {
    super(message);
    this.name = "MicrosoftAuthError";
  }
}
