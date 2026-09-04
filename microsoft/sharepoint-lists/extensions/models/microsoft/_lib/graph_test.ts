// Graph helper tests — retry policy and error-body handling.
// SPDX-License-Identifier: Apache-2.0

import { assertEquals, assertRejects } from "jsr:@std/assert@1.0.19";
import {
  GraphApiError,
  graphRequest,
  graphRequestPaginated,
  retryDelayMs,
} from "./graph.ts";

/** Collects the delays a retry path asked for, without ever sleeping. */
function recordingSleep() {
  const waited: number[] = [];
  return {
    waited,
    sleepFn: (ms: number) => {
      waited.push(ms);
      return Promise.resolve();
    },
  };
}

function json(
  body: unknown,
  status = 200,
  headers: HeadersInit = {},
): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "content-type": "application/json", ...headers },
  });
}

Deno.test("retryDelayMs: Retry-After in seconds wins over backoff", () => {
  const r = new Response(null, {
    status: 429,
    headers: { "Retry-After": "7" },
  });
  assertEquals(retryDelayMs(r, 0, 30_000), 7000);
});

Deno.test("retryDelayMs: Retry-After is capped by maxDelayMs", () => {
  const r = new Response(null, {
    status: 429,
    headers: { "Retry-After": "600" },
  });
  assertEquals(retryDelayMs(r, 0, 30_000), 30_000);
});

Deno.test("retryDelayMs: an HTTP-date Retry-After is honoured", () => {
  const when = new Date(Date.now() + 5000).toUTCString();
  const r = new Response(null, {
    status: 503,
    headers: { "Retry-After": when },
  });
  const ms = retryDelayMs(r, 0, 30_000);
  // Second-granularity date, so allow a little slack either way.
  assertEquals(ms >= 3500 && ms <= 5500, true, `got ${ms}`);
});

Deno.test("retryDelayMs: no header falls back to exponential backoff", () => {
  const r = new Response(null, { status: 429 });
  assertEquals(retryDelayMs(r, 0, 30_000), 1000);
  assertEquals(retryDelayMs(r, 1, 30_000), 2000);
  assertEquals(retryDelayMs(r, 2, 30_000), 4000);
  assertEquals(retryDelayMs(r, 9, 30_000), 30_000);
});

Deno.test("graphRequest: a 429 is retried and then succeeds", async () => {
  const { waited, sleepFn } = recordingSleep();
  let calls = 0;
  const result = await graphRequest<{ ok: boolean }>(
    "token",
    "GET",
    "/me",
    undefined,
    undefined,
    () => {
      calls++;
      return Promise.resolve(
        calls === 1
          ? json({ error: { code: "activityLimitReached" } }, 429, {
            "Retry-After": "2",
          })
          : json({ ok: true }),
      );
    },
    { sleepFn },
  );

  assertEquals(result.ok, true);
  assertEquals(calls, 2);
  assertEquals(waited, [2000]);
});

Deno.test("graphRequest: retries are bounded and the last error surfaces", async () => {
  const { waited, sleepFn } = recordingSleep();
  let calls = 0;
  const err = await assertRejects(
    () =>
      graphRequest("token", "GET", "/me", undefined, undefined, () => {
        calls++;
        return Promise.resolve(
          json({
            error: { code: "activityLimitReached", message: "slow down" },
          }, 429),
        );
      }, { maxRetries: 2, sleepFn }),
    GraphApiError,
  );

  assertEquals((err as GraphApiError).statusCode, 429);
  assertEquals(calls, 3); // initial attempt + 2 retries
  assertEquals(waited, [1000, 2000]);
});

Deno.test("graphRequest: maxRetries 0 disables retrying", async () => {
  let calls = 0;
  await assertRejects(
    () =>
      graphRequest("token", "GET", "/me", undefined, undefined, () => {
        calls++;
        return Promise.resolve(json({ error: { code: "throttled" } }, 429));
      }, { maxRetries: 0 }),
    GraphApiError,
  );
  assertEquals(calls, 1);
});

Deno.test("graphRequest: a 404 is not retried", async () => {
  let calls = 0;
  await assertRejects(
    () =>
      graphRequest("token", "GET", "/me", undefined, undefined, () => {
        calls++;
        return Promise.resolve(json({ error: { code: "itemNotFound" } }, 404));
      }),
    GraphApiError,
  );
  assertEquals(calls, 1);
});

Deno.test("graphRequest: a non-JSON error body reports the status, not a SyntaxError", async () => {
  const err = await assertRejects(
    () =>
      graphRequest(
        "token",
        "GET",
        "/me",
        undefined,
        undefined,
        () =>
          Promise.resolve(
            new Response("<html><body>502 Bad Gateway</body></html>", {
              status: 502,
              headers: { "content-type": "text/html" },
            }),
          ),
      ),
    GraphApiError,
  );
  const e = err as GraphApiError;
  assertEquals(e.statusCode, 502);
  assertEquals(e.graphCode, "non_json_response");
  assertEquals(e.message.includes("502"), true);
});

Deno.test("graphRequestPaginated: the retry policy reaches every page", async () => {
  const { waited, sleepFn } = recordingSleep();
  const responses = [
    // page 1 throttles once, then yields a nextLink
    json({ error: { code: "throttled" } }, 429, { "Retry-After": "1" }),
    json({
      value: [{ id: "a" }],
      "@odata.nextLink": "https://graph.microsoft.com/v1.0/next",
    }),
    // page 2 throttles once too
    json({ error: { code: "throttled" } }, 429, { "Retry-After": "3" }),
    json({ value: [{ id: "b" }] }),
  ];
  let i = 0;

  const result = await graphRequestPaginated<{ id: string }>(
    "token",
    "/things",
    undefined,
    undefined,
    () => Promise.resolve(responses[i++]),
    20,
    { sleepFn },
  );

  assertEquals(result.items.map((x) => x.id), ["a", "b"]);
  assertEquals(result.truncated, false);
  // Both pages retried, proving the policy is forwarded per-request.
  assertEquals(waited, [1000, 3000]);
});
