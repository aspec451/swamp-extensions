// SharePoint Lists Model Tests
// SPDX-License-Identifier: Apache-2.0
// deno-lint-ignore-file no-import-prefix

import {
  assertEquals,
  assertExists,
  assertNotEquals,
  assertRejects,
  assertStringIncludes,
} from "jsr:@std/assert@1";
import {
  createModelTestContext,
  withMockedFetch,
} from "@systeminit/swamp-testing";
import { model } from "./sharepoint_lists.ts";

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------

const SITE_HOST_PATH = "contoso.sharepoint.com:/sites/Test";
const SITE_ID = "site-123";
const LIST_ID = "11111111-2222-3333-4444-555555555555";

const GLOBAL_ARGS = {
  tenantId: "test-tenant",
  clientId: "test-client",
  refreshToken: "test-refresh-token",
  siteHostPath: SITE_HOST_PATH,
  scopes: "offline_access User.Read Sites.Read.All",
};

const DEFINITION = {
  id: "test-id",
  name: "test-sp-lists",
  version: 1,
  tags: {},
};

const TOKEN_OK = {
  access_token: "test-access-token",
  refresh_token: "rotated-refresh-token",
  expires_in: 3600,
  token_type: "Bearer",
  scope: "offline_access User.Read Sites.Read.All",
};

const SITE = {
  id: SITE_ID,
  name: "Test",
  displayName: "Test Site",
  webUrl: "https://contoso.sharepoint.com/sites/Test",
};

const TRACKER_LIST = {
  id: LIST_ID,
  name: "ProjectTracker0",
  displayName: "Project Tracker",
  description: "Work items",
  webUrl: "https://contoso.sharepoint.com/sites/Test/Lists/ProjectTracker0",
  lastModifiedDateTime: "2026-09-01T10:00:00Z",
  list: { template: "genericList", hidden: false },
};

const HIDDEN_LIST = {
  id: "99999999-0000-0000-0000-000000000000",
  name: "FormServerTemplates",
  displayName: "Form Templates",
  webUrl: "https://contoso.sharepoint.com/sites/Test/FormServerTemplates",
  list: { template: "documentLibrary", hidden: true },
};

function json(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "content-type": "application/json" },
  });
}

function graphError(code: string, message: string, status: number): Response {
  return json({ error: { code, message } }, status);
}

/**
 * Route Entra and Graph calls by decoded pathname.
 *
 * Matching on the pathname rather than the whole URL keeps the tests
 * indifferent to query-string encoding: the model builds some URLs by hand
 * (literal `$select`) and others through URLSearchParams (which percent-encodes
 * `$` to `%24`), and both are valid for Graph.
 */
function handlerFor(
  extra: (path: string, u: URL, req: Request) => Response | undefined,
): (req: Request) => Response {
  return (req: Request) => {
    const u = new URL(req.url);
    const path = decodeURIComponent(u.pathname);

    if (path.endsWith("/oauth2/v2.0/token")) return json(TOKEN_OK);
    if (path === `/v1.0/sites/${SITE_HOST_PATH}`) return json(SITE);

    const response = extra(path, u, req);
    if (response) return response;

    return graphError(
      "unexpectedCall",
      `No mock route for ${req.method} ${path}`,
      500,
    );
  };
}

/** The site + list lookups every list-scoped method performs first. */
function withListResolved(
  extra: (path: string, u: URL, req: Request) => Response | undefined,
): (req: Request) => Response {
  return handlerFor((path, u, req) => {
    if (path === `/v1.0/sites/${SITE_ID}/lists/${LIST_ID}`) {
      return json(TRACKER_LIST);
    }
    return extra(path, u, req);
  });
}

// deno-lint-ignore no-explicit-any
type ExecCtx = any;

// ---------------------------------------------------------------------------
// Model Export Structure Tests
// ---------------------------------------------------------------------------

Deno.test("sharepoint-list model: has correct type", () => {
  assertEquals(model.type, "@aspec451/microsoft/sharepoint-lists");
});

Deno.test("sharepoint-list model: has valid version format", () => {
  const versionPattern = /^\d{4}\.\d{2}\.\d{2}\.\d+$/;
  assertEquals(versionPattern.test(model.version), true);
});

Deno.test("sharepoint-list model: has globalArguments with credentials and site", () => {
  assertExists(model.globalArguments);
  const shape = model.globalArguments.shape;
  assertExists(shape.tenantId);
  assertExists(shape.clientId);
  assertExists(shape.refreshToken);
  assertExists(shape.siteHostPath);
  assertExists(shape.scopes);
});

Deno.test("sharepoint-list model: refreshToken is marked sensitive", () => {
  const meta = model.globalArguments.shape.refreshToken.meta();
  assertEquals(meta?.sensitive, true);
});

Deno.test("sharepoint-list model: has required resources", () => {
  assertExists(model.resources);
  assertExists(model.resources.lists);
  assertExists(model.resources.listSchema);
  assertExists(model.resources.listItems);
  assertExists(model.resources.itemWrite);
  assertExists(model.resources.listChanges);
  assertExists(model.resources.listBaseline);
  assertExists(model.resources.bootstrap);
});

Deno.test("sharepoint-list model: has required methods", () => {
  assertExists(model.methods);
  assertExists(model.methods.bootstrap);
  assertExists(model.methods.list_lists);
  assertExists(model.methods.get_list);
  assertExists(model.methods.list_items);
  assertExists(model.methods.list_changes);
  assertExists(model.methods.create_item);
  assertExists(model.methods.update_item);
  assertExists(model.methods.delete_item);
});

Deno.test("sharepoint-list model: list_items has filter and paging arguments", () => {
  const shape = model.methods.list_items.arguments.shape;
  assertExists(shape.list);
  assertExists(shape.filter);
  assertExists(shape.orderBy);
  assertExists(shape.select);
  assertExists(shape.top);
  assertExists(shape.maxPages);
});

Deno.test("sharepoint-list model: list_changes has scope, tracking, and baseline arguments", () => {
  const shape = model.methods.list_changes.arguments.shape;
  assertExists(shape.list);
  assertExists(shape.scopeColumn);
  assertExists(shape.scopeValue);
  assertExists(shape.filter);
  assertExists(shape.trackFields);
  assertExists(shape.ignoreFields);
  assertExists(shape.contextFields);
  assertExists(shape.summarizeColumn);
  assertExists(shape.updateBaseline);
  assertExists(shape.top);
  assertExists(shape.maxPages);
});

Deno.test("sharepoint-list model: list_changes consumes its window by default", () => {
  // A scheduled run that leaves the baseline in place re-reports the same
  // changes forever, so the default has to be the consuming one.
  assertEquals(
    model.methods.list_changes.arguments.parse({ list: "x" }).updateBaseline,
    true,
  );
});

Deno.test("sharepoint-list model: write methods require a fields object", () => {
  assertExists(model.methods.create_item.arguments.shape.fields);
  assertExists(model.methods.update_item.arguments.shape.fields);
  assertExists(model.methods.update_item.arguments.shape.itemId);
});

Deno.test("sharepoint-list model: delete_item has a confirm guard argument", () => {
  assertExists(model.methods.delete_item.arguments.shape.confirm);
});

// ---------------------------------------------------------------------------
// bootstrap
// ---------------------------------------------------------------------------

Deno.test("bootstrap: completes device code flow and persists the rotated refresh token", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  let tokenCalls = 0;
  const { result } = await withMockedFetch(
    (req: Request) => {
      const path = new URL(req.url).pathname;
      if (path.endsWith("/oauth2/v2.0/devicecode")) {
        return json({
          device_code: "dc-1",
          user_code: "ABC-DEF",
          verification_uri: "https://microsoft.com/devicelogin",
          expires_in: 900,
          interval: 0,
          message: "Enter ABC-DEF",
        });
      }
      if (path.endsWith("/oauth2/v2.0/token")) {
        tokenCalls++;
        // First poll is still pending — exercises the retry branch.
        if (tokenCalls === 1) {
          return json({ error: "authorization_pending" }, 400);
        }
        return json(TOKEN_OK);
      }
      return graphError("unexpectedCall", path, 500);
    },
    () => model.methods.bootstrap.execute({}, context as ExecCtx),
  );

  assertEquals(result.dataHandles.length, 1);
  assertEquals(tokenCalls, 2);

  const resources = getWrittenResources();
  assertEquals(resources[0].specName, "bootstrap");
  assertEquals(resources[0].data.status, "authenticated");
  // The token stored must be the freshly issued one, not the input token.
  assertEquals(resources[0].data.refreshToken, "rotated-refresh-token");
});

Deno.test("bootstrap: surfaces an expired refresh token as re-run guidance", async () => {
  const { context } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  await withMockedFetch(
    () => json({ error: "invalid_grant" }, 400),
    async () => {
      const err = await assertRejects(
        () =>
          model.methods.list_lists.execute(
            { includeHidden: false, maxPages: 20 },
            context as ExecCtx,
          ),
        Error,
      );
      assertStringIncludes(err.message, "bootstrap");
    },
  );
});

Deno.test("bootstrap: a token response with no refresh token fails rather than reporting success", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: { ...GLOBAL_ARGS, scopes: "User.Read Sites.Read.All" },
    definition: DEFINITION,
  });

  await withMockedFetch(
    (req: Request) => {
      const path = new URL(req.url).pathname;
      if (path.endsWith("/oauth2/v2.0/devicecode")) {
        return json({
          device_code: "dc-1",
          user_code: "ABC-DEF",
          verification_uri: "https://microsoft.com/devicelogin",
          expires_in: 900,
          interval: 0,
          message: "Enter ABC-DEF",
        });
      }
      if (path.endsWith("/oauth2/v2.0/token")) {
        // What Entra answers when the scope set omits offline_access: an
        // access token, and nothing durable.
        return json({
          access_token: "test-access-token",
          expires_in: 3600,
          token_type: "Bearer",
          scope: "User.Read Sites.Read.All",
        });
      }
      return graphError("unexpectedCall", path, 500);
    },
    async () => {
      const err = await assertRejects(
        () => model.methods.bootstrap.execute({}, context as ExecCtx),
        Error,
      );
      assertStringIncludes(err.message, "offline_access");
    },
  );

  // And nothing is persisted: a record saying "authenticated" with no token
  // to store is worse than no record at all.
  assertEquals(getWrittenResources().length, 0);
});

// ---------------------------------------------------------------------------
// list_lists
// ---------------------------------------------------------------------------

Deno.test("list_lists: returns visible lists and filters hidden ones by default", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  const { result } = await withMockedFetch(
    handlerFor((path) =>
      path === `/v1.0/sites/${SITE_ID}/lists`
        ? json({ value: [TRACKER_LIST, HIDDEN_LIST] })
        : undefined
    ),
    () =>
      model.methods.list_lists.execute(
        { includeHidden: false, maxPages: 20 },
        context as ExecCtx,
      ),
  );

  assertEquals(result.dataHandles.length, 1);
  const data = getWrittenResources()[0].data;
  assertEquals(data.siteId, SITE_ID);
  assertEquals(data.totalFetched, 1);
  assertEquals(data.truncated, false);
  const lists = data.lists as Array<Record<string, unknown>>;
  assertEquals(lists.length, 1);
  assertEquals(lists[0].displayName, "Project Tracker");
  assertEquals(lists[0].template, "genericList");
});

Deno.test("list_lists: includeHidden keeps system lists", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  await withMockedFetch(
    handlerFor((path) =>
      path === `/v1.0/sites/${SITE_ID}/lists`
        ? json({ value: [TRACKER_LIST, HIDDEN_LIST] })
        : undefined
    ),
    () =>
      model.methods.list_lists.execute(
        { includeHidden: true, maxPages: 20 },
        context as ExecCtx,
      ),
  );

  const data = getWrittenResources()[0].data;
  assertEquals(data.totalFetched, 2);
  assertEquals(data.includedHidden, true);
});

// ---------------------------------------------------------------------------
// get_list — column discovery and list resolution
// ---------------------------------------------------------------------------

Deno.test("get_list: resolves a list by display name after the direct lookup 404s", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  let directLookups = 0;
  await withMockedFetch(
    handlerFor((path) => {
      // Graph only accepts the GUID or the URL name here, so a display name misses.
      if (path === `/v1.0/sites/${SITE_ID}/lists/Project Tracker`) {
        directLookups++;
        return graphError(
          "itemNotFound",
          "Requested site could not be found",
          404,
        );
      }
      if (path === `/v1.0/sites/${SITE_ID}/lists`) {
        return json({ value: [HIDDEN_LIST, TRACKER_LIST] });
      }
      if (path === `/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/columns`) {
        return json({
          value: [
            { name: "Title", displayName: "Title", required: true, text: {} },
            {
              name: "Status",
              displayName: "Status",
              choice: { choices: ["Open", "Closed"] },
            },
            {
              name: "Created",
              displayName: "Created",
              readOnly: true,
              dateTime: {},
            },
          ],
        });
      }
      return undefined;
    }),
    () =>
      model.methods.get_list.execute(
        { list: "Project Tracker", includeHidden: false, maxPages: 20 },
        context as ExecCtx,
      ),
  );

  assertEquals(directLookups, 1);
  const data = getWrittenResources()[0].data;
  assertEquals((data.list as Record<string, unknown>).id, LIST_ID);

  const columns = data.columns as Array<Record<string, unknown>>;
  assertEquals(columns.length, 3);
  assertEquals(columns[1].type, "choice");
  assertEquals(columns[1].choices, ["Open", "Closed"]);
  assertEquals(columns[2].type, "dateTime");
  assertEquals(columns[2].readOnly, true);

  // Read-only columns must not be advertised as writable.
  assertEquals(data.writableColumns, ["Title", "Status"]);
});

Deno.test("get_list: a list that matches nothing names the lists that do exist", async () => {
  const { context } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  await withMockedFetch(
    handlerFor((path) => {
      if (path === `/v1.0/sites/${SITE_ID}/lists/Nope`) {
        return graphError("itemNotFound", "not found", 404);
      }
      if (path === `/v1.0/sites/${SITE_ID}/lists`) {
        return json({ value: [TRACKER_LIST] });
      }
      return undefined;
    }),
    async () => {
      const err = await assertRejects(
        () =>
          model.methods.get_list.execute(
            { list: "Nope", includeHidden: false, maxPages: 20 },
            context as ExecCtx,
          ),
        Error,
      );
      assertStringIncludes(err.message, "Project Tracker");
    },
  );
});

// ---------------------------------------------------------------------------
// list_items — reads, filtering, truncation
// ---------------------------------------------------------------------------

Deno.test("list_items: returns items with fields expanded", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  const { result } = await withMockedFetch(
    withListResolved((path) =>
      path === `/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items`
        ? json({
          value: [
            {
              id: "1",
              createdDateTime: "2026-08-01T09:00:00Z",
              lastModifiedDateTime: "2026-08-02T09:00:00Z",
              webUrl: "https://contoso.sharepoint.com/item/1",
              createdBy: { user: { displayName: "A Person" } },
              lastModifiedBy: { user: { displayName: "B Person" } },
              fields: { Title: "First", Status: "Open" },
            },
          ],
        })
        : undefined
    ),
    () =>
      model.methods.list_items.execute(
        { list: LIST_ID, top: 200, maxPages: 20 },
        context as ExecCtx,
      ),
  );

  assertEquals(result.dataHandles.length, 1);
  const data = getWrittenResources()[0].data;
  assertEquals(data.listId, LIST_ID);
  assertEquals(data.listName, "Project Tracker");
  assertEquals(data.totalFetched, 1);
  assertEquals(data.truncated, false);

  const items = data.items as Array<Record<string, unknown>>;
  assertEquals(items[0].fields, { Title: "First", Status: "Open" });
  assertEquals(items[0].createdBy, "A Person");
  assertEquals(items[0].lastModifiedBy, "B Person");
});

Deno.test("list_items: a filter opts in to the non-indexed-column Prefer header", async () => {
  const { context } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  const { calls } = await withMockedFetch(
    withListResolved((path) =>
      path === `/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items`
        ? json({ value: [] })
        : undefined
    ),
    () =>
      model.methods.list_items.execute(
        {
          list: LIST_ID,
          filter: "fields/Status eq 'Open'",
          orderBy: "fields/Created desc",
          top: 200,
          maxPages: 20,
        },
        context as ExecCtx,
      ),
  );

  const itemsCall = calls.find((c) => c.url.includes("/items"));
  assertExists(itemsCall);
  assertEquals(
    itemsCall.headers["prefer"],
    "HonorNonIndexedQueriesWarningMayFailRandomly",
  );
  const query = new URL(itemsCall.url).searchParams;
  assertEquals(query.get("$filter"), "fields/Status eq 'Open'");
  assertEquals(query.get("$orderby"), "fields/Created desc");
  assertEquals(query.get("$expand"), "fields");
});

Deno.test("list_items: no filter means no Prefer header", async () => {
  const { context } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  const { calls } = await withMockedFetch(
    withListResolved((path) =>
      path === `/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items`
        ? json({ value: [] })
        : undefined
    ),
    () =>
      model.methods.list_items.execute(
        { list: LIST_ID, top: 200, maxPages: 20 },
        context as ExecCtx,
      ),
  );

  const itemsCall = calls.find((c) => c.url.includes("/items"));
  assertExists(itemsCall);
  assertEquals(itemsCall.headers["prefer"], undefined);
});

Deno.test("list_items: select scopes the expanded fields", async () => {
  const { context } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  const { calls } = await withMockedFetch(
    withListResolved((path) =>
      path === `/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items`
        ? json({ value: [] })
        : undefined
    ),
    () =>
      model.methods.list_items.execute(
        { list: LIST_ID, select: "Title,Status", top: 200, maxPages: 20 },
        context as ExecCtx,
      ),
  );

  const itemsCall = calls.find((c) => c.url.includes("/items"));
  assertExists(itemsCall);
  assertEquals(
    new URL(itemsCall.url).searchParams.get("$expand"),
    "fields($select=Title,Status)",
  );
});

Deno.test("list_items: hitting the page cap sets truncated and warns that the count is a floor", async () => {
  const { context, getWrittenResources, getLogsByLevel } =
    createModelTestContext({
      globalArgs: GLOBAL_ARGS,
      definition: DEFINITION,
    });

  await withMockedFetch(
    withListResolved((path) =>
      path === `/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items`
        ? json({
          value: [{ id: "1", fields: { Title: "First" } }],
          "@odata.nextLink":
            `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items?$skiptoken=abc`,
        })
        : undefined
    ),
    () =>
      model.methods.list_items.execute(
        { list: LIST_ID, top: 1, maxPages: 1 },
        context as ExecCtx,
      ),
  );

  const data = getWrittenResources()[0].data;
  assertEquals(data.truncated, true);
  assertEquals(data.totalFetched, 1);

  // Regression guard: the runtime logger is LogTape, whose method is warn()
  // while the level it records is "warning". Calling logger.warning() here
  // throws a TypeError instead of warning, and this is the only path that
  // reaches it.
  const warnings = getLogsByLevel("warning");
  assertEquals(warnings.length, 1);
  assertStringIncludes(warnings[0].message, "maxPages");
});

// ---------------------------------------------------------------------------
// Write guards
// ---------------------------------------------------------------------------

Deno.test("create_item: empty fields is refused before any network call", async () => {
  const { context } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  const { calls } = await withMockedFetch(
    () => graphError("unexpectedCall", "should not be reached", 500),
    async () => {
      const err = await assertRejects(
        () =>
          model.methods.create_item.execute(
            { list: LIST_ID, fields: {} },
            context as ExecCtx,
          ),
        Error,
      );
      assertStringIncludes(err.message, "fields is empty");
    },
  );

  assertEquals(calls.length, 0);
});

Deno.test("update_item: empty fields is refused before any network call", async () => {
  const { context } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  const { calls } = await withMockedFetch(
    () => graphError("unexpectedCall", "should not be reached", 500),
    async () => {
      await assertRejects(
        () =>
          model.methods.update_item.execute(
            { list: LIST_ID, itemId: "1", fields: {} },
            context as ExecCtx,
          ),
        Error,
        "no-op",
      );
    },
  );

  assertEquals(calls.length, 0);
});

Deno.test("delete_item: refuses without confirm and makes no network call", async () => {
  const { context } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  const { calls } = await withMockedFetch(
    () => graphError("unexpectedCall", "should not be reached", 500),
    async () => {
      const err = await assertRejects(
        () =>
          model.methods.delete_item.execute(
            { list: LIST_ID, itemId: "7", confirm: false },
            context as ExecCtx,
          ),
        Error,
      );
      assertStringIncludes(err.message, "confirm=true");
    },
  );

  assertEquals(calls.length, 0);
});

// ---------------------------------------------------------------------------
// Writes
// ---------------------------------------------------------------------------

Deno.test("create_item: posts fields and records the created item", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  const { calls } = await withMockedFetch(
    withListResolved((path, _u, req) =>
      path === `/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items` &&
        req.method === "POST"
        ? json({
          id: "42",
          webUrl: "https://contoso.sharepoint.com/item/42",
          fields: { Title: "New request", Status: "Open" },
        })
        : undefined
    ),
    () =>
      model.methods.create_item.execute(
        { list: LIST_ID, fields: { Title: "New request", Status: "Open" } },
        context as ExecCtx,
      ),
  );

  const post = calls.find((c) =>
    c.method === "POST" && c.url.includes("/items")
  );
  assertExists(post);
  assertExists(post.body);
  assertEquals(JSON.parse(post.body), {
    fields: { Title: "New request", Status: "Open" },
  });

  const data = getWrittenResources()[0].data;
  assertEquals(data.operation, "create");
  assertEquals(data.itemId, "42");
  assertEquals(data.fields, { Title: "New request", Status: "Open" });
});

Deno.test("update_item: patches the fields subresource without a wrapper object", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  const { calls } = await withMockedFetch(
    withListResolved((path, _u, req) =>
      path === `/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items/7/fields` &&
        req.method === "PATCH"
        ? json({ Title: "First", Status: "Closed" })
        : undefined
    ),
    () =>
      model.methods.update_item.execute(
        { list: LIST_ID, itemId: "7", fields: { Status: "Closed" } },
        context as ExecCtx,
      ),
  );

  const patch = calls.find((c) => c.method === "PATCH");
  assertExists(patch);
  assertExists(patch.body);
  // Graph's /items/{id}/fields endpoint takes the field map directly.
  assertEquals(JSON.parse(patch.body), { Status: "Closed" });

  const data = getWrittenResources()[0].data;
  assertEquals(data.operation, "update");
  assertEquals(data.itemId, "7");
  assertEquals(data.fields, { Title: "First", Status: "Closed" });
});

Deno.test("update_item: a 204 PATCH keeps the submitted fields in the audit record", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  await withMockedFetch(
    withListResolved((path, _u, req) =>
      path === `/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items/7/fields` &&
        req.method === "PATCH"
        // Graph may answer a field PATCH with 204 and no body.
        ? new Response(null, { status: 204 })
        : undefined
    ),
    () =>
      model.methods.update_item.execute(
        { list: LIST_ID, itemId: "7", fields: { Status: "Closed" } },
        context as ExecCtx,
      ),
  );

  // The itemWrite resource is the record of what changed, so an empty body
  // must not erase it — the submitted fields stand in.
  const data = getWrittenResources()[0].data;
  assertEquals(data.fields, { Status: "Closed" });
});

Deno.test("get_list: hitting the column page cap sets truncated and warns", async () => {
  const { context, getWrittenResources, getLogs } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  await withMockedFetch(
    withListResolved((path, _u, _req) =>
      path === `/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/columns`
        ? json({
          value: [{ name: "Title", text: {} }],
          "@odata.nextLink":
            `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/columns?$skiptoken=2`,
        })
        : undefined
    ),
    () =>
      model.methods.get_list.execute(
        { list: LIST_ID, includeHidden: false, maxPages: 1 },
        context as ExecCtx,
      ),
  );

  const data = getWrittenResources()[0].data;
  assertEquals(data.truncated, true);
  assertEquals((data.columns as unknown[]).length, 1);

  const warned = getLogs().some((l: { message: string }) =>
    l.message.includes("maxPages cap")
  );
  assertEquals(warned, true, "expected a warning that columns are incomplete");
});

Deno.test("get_list: an exhausted column walk reports truncated false", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  await withMockedFetch(
    withListResolved((path, _u, _req) =>
      path === `/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/columns`
        ? json({ value: [{ name: "Title", text: {} }] })
        : undefined
    ),
    () =>
      model.methods.get_list.execute(
        { list: LIST_ID, includeHidden: false, maxPages: 20 },
        context as ExecCtx,
      ),
  );

  assertEquals(getWrittenResources()[0].data.truncated, false);
});

Deno.test("delete_item: reads the item before deleting and keeps the pre-image", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  const { calls } = await withMockedFetch(
    withListResolved((path, _u, req) => {
      const itemPath = `/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items/7`;
      if (path === itemPath && req.method === "GET") {
        return json({
          id: "7",
          webUrl: "https://contoso.sharepoint.com/item/7",
          fields: { Title: "Doomed", Status: "Open" },
        });
      }
      if (path === itemPath && req.method === "DELETE") {
        return new Response(null, { status: 204 });
      }
      return undefined;
    }),
    () =>
      model.methods.delete_item.execute(
        { list: LIST_ID, itemId: "7", confirm: true },
        context as ExecCtx,
      ),
  );

  // The read must precede the delete, or there is no record of what was lost.
  const readIndex = calls.findIndex((c) =>
    c.method === "GET" && c.url.includes("/items/7")
  );
  const deleteIndex = calls.findIndex((c) => c.method === "DELETE");
  assertEquals(readIndex >= 0, true);
  assertEquals(deleteIndex >= 0, true);
  assertEquals(readIndex < deleteIndex, true);

  const data = getWrittenResources()[0].data;
  assertEquals(data.operation, "delete");
  assertEquals(data.itemId, "7");
  assertEquals(data.fields, { Title: "Doomed", Status: "Open" });
});

// ---------------------------------------------------------------------------
// Error context
// ---------------------------------------------------------------------------

Deno.test("delete_item: an already-deleted item succeeds as alreadyAbsent", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  const { calls } = await withMockedFetch(
    withListResolved((path, _u, req) =>
      path === `/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items/7` &&
        req.method === "GET"
        ? graphError("itemNotFound", "Item not found", 404)
        : undefined
    ),
    () =>
      model.methods.delete_item.execute(
        { list: LIST_ID, itemId: "7", confirm: true },
        context as ExecCtx,
      ),
  );

  // Deleting what is already gone is a no-op that reports success, not a
  // failure — the requested end state already holds.
  const data = getWrittenResources()[0].data;
  assertEquals(data.operation, "delete");
  assertEquals(data.alreadyAbsent, true);
  assertEquals(data.deleteConfirmed, true);
  assertEquals(data.fields, {});
  // Nothing was destroyed on the way to that conclusion.
  assertEquals(calls.some((c) => c.method === "DELETE"), false);
});

Deno.test("delete_item: a non-404 read failure still fails the delete", async () => {
  const { context } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  await withMockedFetch(
    withListResolved((path, _u, req) =>
      path === `/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items/7` &&
        req.method === "GET"
        ? graphError("accessDenied", "Forbidden", 403)
        : undefined
    ),
    async () => {
      const err = await assertRejects(
        () =>
          model.methods.delete_item.execute(
            { list: LIST_ID, itemId: "7", confirm: true },
            context as ExecCtx,
          ),
        Error,
      );
      // A 403 must not be mistaken for "already gone".
      assertStringIncludes(err.message, "before delete");
      assertStringIncludes(err.message, "Forbidden");
    },
  );
});

Deno.test("delete_item: the pre-image is persisted before the DELETE is issued", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  await withMockedFetch(
    withListResolved((path, _u, req) => {
      const itemPath = `/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items/7`;
      if (path === itemPath && req.method === "GET") {
        return json({ id: "7", fields: { Title: "Doomed" } });
      }
      if (path === itemPath && req.method === "DELETE") {
        return new Response(null, { status: 204 });
      }
      return undefined;
    }),
    () =>
      model.methods.delete_item.execute(
        { list: LIST_ID, itemId: "7", confirm: true },
        context as ExecCtx,
      ),
  );

  const written = getWrittenResources();
  // Two versions: the unconfirmed pre-image, then the confirmed record. An
  // interruption between them leaves the pre-image behind rather than nothing.
  assertEquals(written.length, 2);
  assertEquals(written[0].data.deleteConfirmed, false);
  assertEquals(written[0].data.fields, { Title: "Doomed" });
  assertEquals(written[1].data.deleteConfirmed, true);
  assertEquals(written[1].data.fields, { Title: "Doomed" });
  // Both are versions of one instance, not two competing records.
  assertEquals(written[0].name, written[1].name);
});

Deno.test("errors carry the operation and the Graph error message", async () => {
  const { context } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  await withMockedFetch(
    withListResolved((path) =>
      path === `/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items`
        ? graphError("accessDenied", "Access denied", 403)
        : undefined
    ),
    async () => {
      const err = await assertRejects(
        () =>
          model.methods.list_items.execute(
            { list: LIST_ID, top: 200, maxPages: 20 },
            context as ExecCtx,
          ),
        Error,
      );
      assertStringIncludes(err.message, "Failed to list items");
      assertStringIncludes(err.message, "Access denied");
    },
  );
});

// ---------------------------------------------------------------------------
// Instance naming
// ---------------------------------------------------------------------------

Deno.test("create_item: a long list name cannot truncate the item ID out of the instance name", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  // 42 characters — long enough that naive `create-${list}-${id}` slugging
  // would cut the ID off the end and land both items on one instance.
  const longList = {
    ...TRACKER_LIST,
    displayName: "Client Onboarding Requests Archive 2024 Q1",
  };

  let posts = 0;
  await withMockedFetch(
    handlerFor((path, _u, req) => {
      if (path === `/v1.0/sites/${SITE_ID}/lists/${LIST_ID}`) {
        return json(longList);
      }
      if (
        path === `/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items` &&
        req.method === "POST"
      ) {
        posts++;
        return json({
          id: posts === 1 ? "12" : "4711",
          fields: { Title: "Row" },
        });
      }
      return undefined;
    }),
    async () => {
      for (const _ of [0, 1]) {
        await model.methods.create_item.execute(
          { list: LIST_ID, fields: { Title: "Row" } },
          context as ExecCtx,
        );
      }
    },
  );

  const written = getWrittenResources();
  assertEquals(written.length, 2);
  assertEquals(written[0].data.itemId, "12");
  assertEquals(written[1].data.itemId, "4711");
  assertStringIncludes(written[0].name, "-12");
  assertStringIncludes(written[1].name, "-4711");
  // Two items, two records — not one record with a second version.
  assertNotEquals(written[0].name, written[1].name);
});

Deno.test("list_items: two lists sharing a 48-character name prefix write to distinct instances", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  const LIST_B_ID = "22222222-3333-4444-5555-666666666666";
  const prefix = "Regional Field Service Requests Archive Southern Division ";
  const listA = { ...TRACKER_LIST, displayName: `${prefix}Alpha` };
  const listB = {
    ...TRACKER_LIST,
    id: LIST_B_ID,
    displayName: `${prefix}Beta`,
  };

  await withMockedFetch(
    handlerFor((path) => {
      if (path === `/v1.0/sites/${SITE_ID}/lists/${LIST_ID}`) {
        return json(listA);
      }
      if (path === `/v1.0/sites/${SITE_ID}/lists/${LIST_B_ID}`) {
        return json(listB);
      }
      if (path.endsWith("/items")) return json({ value: [] });
      return undefined;
    }),
    async () => {
      for (const list of [LIST_ID, LIST_B_ID]) {
        await model.methods.list_items.execute(
          { list, top: 200, maxPages: 20 },
          context as ExecCtx,
        );
      }
    },
  );

  const written = getWrittenResources();
  assertEquals(written.length, 2);
  assertEquals(written[0].data.listId, LIST_ID);
  assertEquals(written[1].data.listId, LIST_B_ID);
  // Snapshots of two different lists must not share one version history.
  assertNotEquals(written[0].name, written[1].name);
});

// ---------------------------------------------------------------------------
// Error Attribution Tests
// ---------------------------------------------------------------------------

Deno.test("resolveList: a 403 on the direct lookup fails rather than enumerating the site", async () => {
  const { context } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  let enumerations = 0;

  await withMockedFetch(
    handlerFor((path) => {
      if (path === `/v1.0/sites/${SITE_ID}/lists/Project Tracker`) {
        return graphError("accessDenied", "Access denied", 403);
      }
      if (path === `/v1.0/sites/${SITE_ID}/lists`) {
        enumerations++;
        return json({ value: [TRACKER_LIST] });
      }
      return undefined;
    }),
    async () => {
      const err = await assertRejects(
        () =>
          model.methods.list_items.execute(
            { list: "Project Tracker", top: 200, maxPages: 20 },
            context as ExecCtx,
          ),
        Error,
      );
      // The permission failure has to be reported as itself. Recasting it as
      // a missing list — which is what falling through to the display-name
      // search does — sends the user hunting for a list that is right there.
      assertStringIncludes(err.message, "Access denied");
    },
  );

  assertEquals(enumerations, 0);
});

Deno.test("resolveList: a 404 on the direct lookup still falls back to the display-name search", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  await withMockedFetch(
    handlerFor((path) => {
      if (path === `/v1.0/sites/${SITE_ID}/lists/Project Tracker`) {
        return graphError("itemNotFound", "Not found", 404);
      }
      if (path === `/v1.0/sites/${SITE_ID}/lists`) {
        return json({ value: [TRACKER_LIST] });
      }
      if (path === `/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items`) {
        return json({ value: [] });
      }
      return undefined;
    }),
    () =>
      model.methods.list_items.execute(
        { list: "Project Tracker", top: 200, maxPages: 20 },
        context as ExecCtx,
      ),
  );

  assertEquals(getWrittenResources()[0].data.listId, LIST_ID);
});

Deno.test("auth: a non-JSON token endpoint body reports the HTTP status, not a parse error", async () => {
  const { context } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  await withMockedFetch(
    () =>
      new Response("<html><body>502 Bad Gateway</body></html>", {
        status: 502,
        statusText: "Bad Gateway",
        headers: { "content-type": "text/html" },
      }),
    async () => {
      const err = await assertRejects(
        () =>
          model.methods.list_lists.execute(
            { includeHidden: false, maxPages: 20 },
            context as ExecCtx,
          ),
        Error,
      );
      // A gateway in front of login.microsoftonline.com answers with HTML.
      // The status is the part that explains the failure; a SyntaxError from
      // the JSON parser hides it.
      assertStringIncludes(err.message, "502");
      assertStringIncludes(err.message, "non-JSON");
    },
  );
});

// ---------------------------------------------------------------------------
// Pre-flight Check Tests
// ---------------------------------------------------------------------------

const WRITE_SCOPES = "offline_access User.Read Sites.ReadWrite.All";

/** Run one check against the given global arguments. */
function runCheck(
  name: keyof typeof model.checks,
  globalArgs: Record<string, unknown>,
) {
  const { context } = createModelTestContext({
    globalArgs: globalArgs as typeof GLOBAL_ARGS,
    definition: DEFINITION,
  });
  return model.checks[name].execute(context as ExecCtx);
}

Deno.test("checks: the write methods are guarded by scope and live credential checks", () => {
  assertExists(model.checks);

  // Every mutating method must be covered by the live credential check —
  // a write that fails at Graph has already resolved the site and the list.
  const live = model.checks["credentials-resolve-site"];
  assertEquals(live.labels, ["live"]);
  assertEquals(live.appliesTo, ["create_item", "update_item", "delete_item"]);

  const scope = model.checks["write-scope-granted"];
  assertEquals(scope.labels, ["policy"]);
  assertEquals(scope.appliesTo, ["create_item", "update_item", "delete_item"]);

  // The locator check is cheap and applies everywhere, so it carries no
  // appliesTo at all.
  assertEquals(model.checks["site-locator-format"].labels, ["policy"]);
  assertEquals(
    "appliesTo" in model.checks["site-locator-format"],
    false,
  );
});

Deno.test("check site-locator-format: accepts a host:path locator and a bare host", async () => {
  for (
    const siteHostPath of [
      SITE_HOST_PATH,
      "contoso.sharepoint.com",
      "contoso.sharepoint.com,aaaa,bbbb",
    ]
  ) {
    const result = await runCheck("site-locator-format", {
      ...GLOBAL_ARGS,
      siteHostPath,
    });
    assertEquals(result.pass, true, `rejected ${siteHostPath}`);
  }
});

Deno.test("check site-locator-format: rejects a pasted browser URL", async () => {
  const result = await runCheck("site-locator-format", {
    ...GLOBAL_ARGS,
    siteHostPath: "https://contoso.sharepoint.com/sites/Test",
  });

  assertEquals(result.pass, false);
  assertStringIncludes(result.errors![0], "browser URL");
});

Deno.test("check site-locator-format: rejects a path with no colon separating the host", async () => {
  const result = await runCheck("site-locator-format", {
    ...GLOBAL_ARGS,
    siteHostPath: "contoso.sharepoint.com/sites/Test",
  });

  assertEquals(result.pass, false);
  assertStringIncludes(result.errors![0], "hostname");
});

Deno.test("check site-locator-format: rejects surrounding whitespace", async () => {
  const result = await runCheck("site-locator-format", {
    ...GLOBAL_ARGS,
    siteHostPath: ` ${SITE_HOST_PATH} `,
  });

  assertEquals(result.pass, false);
  assertStringIncludes(result.errors![0], "whitespace");
});

Deno.test("check write-scope-granted: read-only scopes fail before a write is attempted", async () => {
  // GLOBAL_ARGS is deliberately a read-only deployment (Sites.Read.All).
  const result = await runCheck("write-scope-granted", GLOBAL_ARGS);

  assertEquals(result.pass, false);
  assertStringIncludes(result.errors![0], "Sites.ReadWrite.All");
});

Deno.test("check write-scope-granted: a write scope passes, fully-qualified or bare", async () => {
  for (
    const scopes of [
      WRITE_SCOPES,
      "offline_access https://graph.microsoft.com/Sites.ReadWrite.All",
      "offline_access Sites.FullControl.All",
    ]
  ) {
    const result = await runCheck("write-scope-granted", {
      ...GLOBAL_ARGS,
      scopes,
    });
    assertEquals(result.pass, true, `rejected ${scopes}`);
  }
});

Deno.test("check credentials-resolve-site: an expired refresh token fails with bootstrap guidance", async () => {
  await withMockedFetch(
    () => json({ error: "invalid_grant" }, 400),
    async () => {
      const result = await runCheck("credentials-resolve-site", {
        ...GLOBAL_ARGS,
        scopes: WRITE_SCOPES,
      });

      assertEquals(result.pass, false);
      assertStringIncludes(result.errors![0], "bootstrap");
    },
  );
});

Deno.test("check credentials-resolve-site: an unresolvable site is reported as the locator, not the credentials", async () => {
  await withMockedFetch(
    (req: Request) => {
      const path = decodeURIComponent(new URL(req.url).pathname);
      if (path.endsWith("/oauth2/v2.0/token")) return json(TOKEN_OK);
      return graphError(
        "itemNotFound",
        "Requested site could not be found",
        404,
      );
    },
    async () => {
      const result = await runCheck("credentials-resolve-site", {
        ...GLOBAL_ARGS,
        scopes: WRITE_SCOPES,
      });

      assertEquals(result.pass, false);
      // Credentials were fine; saying so keeps the user from re-running
      // bootstrap over what is actually a typo in siteHostPath.
      assertStringIncludes(result.errors![0], "did not resolve");
    },
  );
});

Deno.test("check credentials-resolve-site: valid credentials and a resolvable site pass", async () => {
  await withMockedFetch(
    handlerFor(() => undefined),
    async () => {
      const result = await runCheck("credentials-resolve-site", {
        ...GLOBAL_ARGS,
        scopes: WRITE_SCOPES,
      });

      assertEquals(result.pass, true);
    },
  );
});

// ---------------------------------------------------------------------------
// list_changes
// ---------------------------------------------------------------------------

const ITEMS_PATH = `/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items`;
const COLUMNS_PATH = `/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/columns`;

/** Instance names `changeInstances` derives for "Project Tracker". */
const CHANGES_ALL = "project-tracker-all-rows";
const BASELINE_ALL = "baseline-project-tracker-all-rows";
const BASELINE_OPEN = "baseline-project-tracker-status-open";

/** The column definitions the diff reads its display names from. */
const COLUMNS = {
  value: [
    { name: "Title", displayName: "Item" },
    { name: "Status", displayName: "Request Status" },
    { name: "Notes", displayName: "Notes" },
  ],
};

/** The arguments zod would have defaulted, since execute is called directly. */
const CHANGE_ARGS = {
  list: LIST_ID,
  ignoreFields: [] as string[],
  contextFields: [] as string[],
  updateBaseline: true,
  top: 1000,
  maxPages: 20,
};

/** One Graph list item, with the version SharePoint carries in its fields. */
function spItem(
  id: string,
  fields: Record<string, unknown>,
  lastModified = "2026-09-22T09:00:00Z",
  version = "3.0",
): Record<string, unknown> {
  return {
    id,
    lastModifiedDateTime: lastModified,
    lastModifiedBy: { user: { displayName: "B Person" } },
    fields: { id, _UIVersionString: version, ...fields },
  };
}

/** A stored baseline, as a previous clean run would have written it. */
function storedBaseline(
  rows: Array<{
    itemId: string;
    fields: Record<string, string>;
    lastModified?: string;
    version?: string;
  }>,
  overrides: Record<string, unknown> = {},
): Record<string, unknown> {
  return {
    listId: LIST_ID,
    listName: "Project Tracker",
    scopeColumn: null,
    scopeValue: null,
    fieldPolicy: "track=*;ignore=",
    capturedAt: "2026-09-20T07:30:00Z",
    rowCount: rows.length,
    rows: rows.map((row) => ({
      itemId: row.itemId,
      title: row.fields.Title ?? null,
      lastModified: row.lastModified ?? "2026-09-19T12:00:00Z",
      lastModifiedBy: "B Person",
      version: row.version ?? "2.0",
      itemUrl: null,
      fields: row.fields,
    })),
    fetchedAt: "2026-09-20T07:30:00Z",
    ...overrides,
  };
}

/**
 * Site, list, and column-label routes, with the item collection supplied and
 * single-item GETs (the removal probe) routed separately.
 */
function withChangeRoutes(
  items: (u: URL) => Response,
  probe?: (itemId: string) => Response | undefined,
): (req: Request) => Response {
  return withListResolved((path, u) => {
    if (path === COLUMNS_PATH) return json(COLUMNS);
    if (path === ITEMS_PATH) return items(u);
    if (path.startsWith(`${ITEMS_PATH}/`)) {
      return probe?.(path.slice(ITEMS_PATH.length + 1));
    }
    return undefined;
  });
}

/** The written resource for one spec, or undefined when nothing was written. */
function written(
  resources: Array<{ specName: string; name: string; data: unknown }>,
  specName: string,
): { specName: string; name: string; data: Record<string, unknown> } {
  const match = resources.find((r) => r.specName === specName);
  assertExists(match, `no ${specName} resource was written`);
  return match as {
    specName: string;
    name: string;
    data: Record<string, unknown>;
  };
}

Deno.test("list_changes: scopeColumn without scopeValue is refused before any network call", async () => {
  const { context } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  const { calls } = await withMockedFetch(
    () => graphError("unexpectedCall", "should not be reached", 500),
    async () => {
      const err = await assertRejects(
        () =>
          model.methods.list_changes.execute(
            { ...CHANGE_ARGS, scopeColumn: "Status" },
            context as ExecCtx,
          ),
        Error,
      );
      assertStringIncludes(err.message, "go together");
    },
  );

  assertEquals(calls.length, 0);
});

Deno.test("list_changes: a display name as scopeColumn is refused before any network call", async () => {
  const { context } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  const { calls } = await withMockedFetch(
    () => graphError("unexpectedCall", "should not be reached", 500),
    async () => {
      const err = await assertRejects(
        () =>
          model.methods.list_changes.execute(
            {
              ...CHANGE_ARGS,
              scopeColumn: "Assigned To",
              scopeValue: "Platform Team",
            },
            context as ExecCtx,
          ),
        Error,
      );
      assertStringIncludes(err.message, "internal column");
      assertStringIncludes(err.message, "Assigned_x0020_To");
    },
  );

  assertEquals(calls.length, 0);
});

Deno.test("list_changes: a first run establishes the baseline and reports nothing as changed", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  const { result } = await withMockedFetch(
    withChangeRoutes(() =>
      json({
        value: [
          spItem("1", { Title: "First", Status: "Open" }),
          spItem("2", { Title: "Second", Status: "Confirmed" }),
        ],
      })
    ),
    () =>
      model.methods.list_changes.execute(
        { ...CHANGE_ARGS },
        context as ExecCtx,
      ),
  );

  // The changes report first, so a consumer reading dataHandles[0] gets the
  // report rather than the machine state behind it.
  assertEquals(result.dataHandles.length, 2);
  assertEquals(result.dataHandles[0].name, CHANGES_ALL);
  assertEquals(result.dataHandles[1].name, BASELINE_ALL);

  const changes = written(getWrittenResources(), "listChanges");
  assertEquals(changes.name, CHANGES_ALL);
  assertEquals(changes.data.firstRun, true);
  assertEquals(changes.data.since, null);
  assertEquals(changes.data.sinceGapHours, null);
  assertEquals(changes.data.baselineInstance, BASELINE_ALL);
  assertEquals(changes.data.baselineUpdated, true);
  assertEquals(changes.data.baselineHeldReason, null);
  assertEquals(changes.data.added, []);
  assertEquals(changes.data.modified, []);
  assertEquals(changes.data.removed, []);

  const totals = changes.data.totals as Record<string, number>;
  assertEquals(totals.inScope, 2);
  assertEquals(totals.baselineRows, 0);
  assertEquals(totals.changed, 0);

  const baseline = written(getWrittenResources(), "listBaseline");
  assertEquals(baseline.name, BASELINE_ALL);
  assertEquals(baseline.data.rowCount, 2);
  assertEquals(baseline.data.fieldPolicy, "track=*;ignore=");
  const rows = baseline.data.rows as Array<Record<string, unknown>>;
  // The version is corroboration of a bump and is kept; it is not a field.
  assertEquals(rows[0].version, "3.0");
  assertEquals(rows[0].fields, { Title: "First", Status: "Open" });
});

Deno.test("list_changes: a field change is reported with the column's display name", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
    storedResources: {
      [BASELINE_ALL]: storedBaseline([
        { itemId: "1", fields: { Title: "First", Status: "Open" } },
      ]),
    },
  });

  await withMockedFetch(
    withChangeRoutes(() =>
      json({ value: [spItem("1", { Title: "First", Status: "Confirmed" })] })
    ),
    () =>
      model.methods.list_changes.execute(
        { ...CHANGE_ARGS },
        context as ExecCtx,
      ),
  );

  const data = written(getWrittenResources(), "listChanges").data;
  assertEquals(data.firstRun, false);
  assertEquals(data.since, "2026-09-20T07:30:00Z");

  const modified = data.modified as Array<Record<string, unknown>>;
  assertEquals(modified.length, 1);
  assertEquals(modified[0].itemId, "1");
  assertEquals(modified[0].changes, [
    {
      field: "Status",
      label: "Request Status",
      from: "Open",
      to: "Confirmed",
      kind: "changed",
    },
  ]);

  const totals = data.totals as Record<string, number>;
  assertEquals(totals.fieldChanges, 1);
  assertEquals(totals.touched, 0);
  assertEquals(totals.changed, 1);
});

Deno.test("list_changes: an added row carries every tracked column as a set", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
    storedResources: {
      [BASELINE_ALL]: storedBaseline([
        {
          itemId: "1",
          fields: { Title: "First", Status: "Open" },
          lastModified: "2026-09-19T12:00:00Z",
          version: "2.0",
        },
      ]),
    },
  });

  await withMockedFetch(
    withChangeRoutes(() =>
      json({
        value: [
          spItem(
            "1",
            { Title: "First", Status: "Open" },
            "2026-09-19T12:00:00Z",
            "2.0",
          ),
          spItem("2", { Title: "Second", Status: "New" }),
        ],
      })
    ),
    () =>
      model.methods.list_changes.execute(
        { ...CHANGE_ARGS },
        context as ExecCtx,
      ),
  );

  const data = written(getWrittenResources(), "listChanges").data;
  const added = data.added as Array<Record<string, unknown>>;
  assertEquals(added.length, 1);
  assertEquals(added[0].itemId, "2");
  assertEquals(added[0].title, "Second");
  // Display-name order, so the row reads the way the form does: Item, then
  // Request Status.
  assertEquals(added[0].changes, [
    { field: "Title", label: "Item", from: null, to: "Second", kind: "set" },
    {
      field: "Status",
      label: "Request Status",
      from: null,
      to: "New",
      kind: "set",
    },
  ]);
  assertEquals((data.modified as unknown[]).length, 0);
  assertEquals((data.touched as unknown[]).length, 0);
  assertEquals((data.totals as Record<string, number>).changed, 1);
});

Deno.test("list_changes: a bumped row with nothing moved is touched, not modified", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
    storedResources: {
      [BASELINE_ALL]: storedBaseline([
        { itemId: "1", fields: { Title: "First", Status: "Open" } },
      ]),
    },
  });

  await withMockedFetch(
    withChangeRoutes(() =>
      json({
        value: [
          spItem(
            "1",
            { Title: "First", Status: "Open" },
            "2026-09-22T09:00:00Z",
            "4.0",
          ),
        ],
      })
    ),
    () =>
      model.methods.list_changes.execute(
        { ...CHANGE_ARGS },
        context as ExecCtx,
      ),
  );

  const data = written(getWrittenResources(), "listChanges").data;
  const touched = data.touched as Array<Record<string, unknown>>;
  assertEquals(touched.length, 1);
  assertEquals(touched[0].changes, []);
  assertEquals((data.modified as unknown[]).length, 0);
  // A bump with nothing visible behind it is not something to read.
  assertEquals((data.totals as Record<string, number>).changed, 0);
});

Deno.test("list_changes: a row edited out of the scope is rescoped, not deleted", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
    storedResources: {
      [BASELINE_OPEN]: storedBaseline(
        [
          { itemId: "1", fields: { Title: "First", Status: "Open" } },
          { itemId: "9", fields: { Title: "Ninth", Status: "Open" } },
        ],
        { scopeColumn: "Status", scopeValue: "Open" },
      ),
    },
  });

  await withMockedFetch(
    withChangeRoutes(
      () =>
        json({
          value: [
            spItem(
              "1",
              { Title: "First", Status: "Open" },
              "2026-09-19T12:00:00Z",
              "2.0",
            ),
          ],
        }),
      (itemId) =>
        itemId === "9"
          ? json(spItem("9", { Title: "Ninth", Status: "Closed" }))
          : undefined,
    ),
    () =>
      model.methods.list_changes.execute(
        { ...CHANGE_ARGS, scopeColumn: "Status", scopeValue: "Open" },
        context as ExecCtx,
      ),
  );

  const data = written(getWrittenResources(), "listChanges").data;
  const removed = data.removed as Array<Record<string, unknown>>;
  assertEquals(removed.length, 1);
  assertEquals(removed[0].itemId, "9");
  assertEquals(removed[0].fate, "rescoped");
  assertEquals(removed[0].nowInScope, "Closed");
  assertEquals(removed[0].lastKnownFields, { Title: "Ninth", Status: "Open" });
  // A row that left the scope is real news, so the baseline still advances.
  assertEquals(data.baselineUpdated, true);
  assertEquals((data.totals as Record<string, number>).changed, 1);
});

Deno.test("list_changes: an item that no longer resolves is reported as deleted", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
    storedResources: {
      [BASELINE_ALL]: storedBaseline([
        { itemId: "1", fields: { Title: "First", Status: "Open" } },
        { itemId: "9", fields: { Title: "Ninth", Status: "Open" } },
      ]),
    },
  });

  await withMockedFetch(
    withChangeRoutes(
      () =>
        json({
          value: [
            spItem(
              "1",
              { Title: "First", Status: "Open" },
              "2026-09-19T12:00:00Z",
              "2.0",
            ),
          ],
        }),
      (itemId) =>
        itemId === "9"
          ? graphError("itemNotFound", "Item does not exist", 404)
          : undefined,
    ),
    () =>
      model.methods.list_changes.execute(
        { ...CHANGE_ARGS },
        context as ExecCtx,
      ),
  );

  const data = written(getWrittenResources(), "listChanges").data;
  const removed = data.removed as Array<Record<string, unknown>>;
  assertEquals(removed.length, 1);
  assertEquals(removed[0].fate, "deleted");
  assertEquals(removed[0].nowInScope, null);
  // The baseline's copy is all that is left of a deleted row.
  assertEquals(removed[0].lastKnownFields, { Title: "Ninth", Status: "Open" });
  assertEquals(data.baselineUpdated, true);
});

Deno.test("list_changes: a missing row that is still there holds the baseline back", async () => {
  const { context, getWrittenResources, getLogsByLevel } =
    createModelTestContext({
      globalArgs: GLOBAL_ARGS,
      definition: DEFINITION,
      storedResources: {
        [BASELINE_ALL]: storedBaseline([
          { itemId: "1", fields: { Title: "First", Status: "Open" } },
          { itemId: "9", fields: { Title: "Ninth", Status: "Open" } },
        ]),
      },
    });

  const { result } = await withMockedFetch(
    withChangeRoutes(
      () =>
        json({
          value: [
            spItem(
              "1",
              { Title: "First", Status: "Open" },
              "2026-09-19T12:00:00Z",
              "2.0",
            ),
          ],
        }),
      (itemId) =>
        itemId === "9"
          ? json(spItem("9", { Title: "Ninth", Status: "Open" }))
          : undefined,
    ),
    () =>
      model.methods.list_changes.execute(
        { ...CHANGE_ARGS },
        context as ExecCtx,
      ),
  );

  const data = written(getWrittenResources(), "listChanges").data;
  const removed = data.removed as Array<Record<string, unknown>>;
  assertEquals(removed[0].fate, "still-present");

  const totals = data.totals as Record<string, number>;
  assertEquals(totals.stillPresent, 1);
  // A short fetch is news about the fetch, not about the row.
  assertEquals(totals.changed, 0);

  assertEquals(data.baselineUpdated, false);
  assertStringIncludes(String(data.baselineHeldReason), "came back short");
  assertEquals(result.dataHandles.length, 1);
  assertEquals(getWrittenResources().length, 1);
  assertStringIncludes(
    getLogsByLevel("warning")[0].message,
    "Baseline not replaced",
  );
});

Deno.test("list_changes: mass absence is flagged suspect and skips the probes", async () => {
  const rows = Array.from({ length: 12 }, (_, i) => ({
    itemId: String(i + 1),
    fields: { Title: `Row ${i + 1}`, Status: "Open" },
  }));

  const { context, getWrittenResources, getLogsByLevel } =
    createModelTestContext({
      globalArgs: GLOBAL_ARGS,
      definition: DEFINITION,
      storedResources: { [BASELINE_ALL]: storedBaseline(rows) },
    });

  const { calls } = await withMockedFetch(
    withChangeRoutes(() =>
      json({
        value: [
          spItem(
            "1",
            { Title: "Row 1", Status: "Open" },
            "2026-09-19T12:00:00Z",
            "2.0",
          ),
        ],
      })
    ),
    () =>
      model.methods.list_changes.execute(
        { ...CHANGE_ARGS },
        context as ExecCtx,
      ),
  );

  const data = written(getWrittenResources(), "listChanges").data;
  assertEquals(data.suspectRemoval, true);
  assertEquals((data.removed as unknown[]).length, 11);
  // Deliberately skipped, which is a different claim from a probe that ran
  // and could not tell.
  for (const row of data.removed as Array<Record<string, unknown>>) {
    assertEquals(row.fate, "unprobed");
  }

  // A run that cannot see its own population must not spend a Graph call per
  // absent row proving rows it cannot see are still there.
  assertEquals(calls.filter((c) => /\/items\/\d+/.test(c.url)).length, 0);

  assertEquals(data.baselineUpdated, false);
  assertStringIncludes(String(data.baselineHeldReason), "failed filter");
  assertEquals(getWrittenResources().length, 1);
  assertEquals(getLogsByLevel("error").length, 1);
});

Deno.test("list_changes: hitting the page cap holds the baseline back", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  await withMockedFetch(
    withChangeRoutes(() =>
      json({
        value: [spItem("1", { Title: "First", Status: "Open" })],
        "@odata.nextLink":
          `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items?$skiptoken=abc`,
      })
    ),
    () =>
      model.methods.list_changes.execute(
        { ...CHANGE_ARGS, maxPages: 1 },
        context as ExecCtx,
      ),
  );

  const data = written(getWrittenResources(), "listChanges").data;
  assertEquals(data.truncated, true);
  assertEquals(data.baselineUpdated, false);
  assertStringIncludes(String(data.baselineHeldReason), "1-page cap");
  // A partial population must not become the baseline: the rows it missed
  // would read as deletions on the next run.
  assertEquals(getWrittenResources().length, 1);
});

Deno.test("list_changes: updateBaseline=false leaves the window unconsumed", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
    storedResources: {
      [BASELINE_ALL]: storedBaseline([
        { itemId: "1", fields: { Title: "First", Status: "Open" } },
      ]),
    },
  });

  await withMockedFetch(
    withChangeRoutes(() =>
      json({ value: [spItem("1", { Title: "First", Status: "Confirmed" })] })
    ),
    () =>
      model.methods.list_changes.execute(
        { ...CHANGE_ARGS, updateBaseline: false },
        context as ExecCtx,
      ),
  );

  const data = written(getWrittenResources(), "listChanges").data;
  assertEquals((data.modified as unknown[]).length, 1);
  assertEquals(data.baselineUpdated, false);
  assertStringIncludes(String(data.baselineHeldReason), "updateBaseline=false");
  assertEquals(getWrittenResources().length, 1);
});

Deno.test("list_changes: a baseline captured under different tracked columns is not diffed", async () => {
  const { context, getWrittenResources, getLogsByLevel } =
    createModelTestContext({
      globalArgs: GLOBAL_ARGS,
      definition: DEFINITION,
      storedResources: {
        [BASELINE_ALL]: storedBaseline(
          [{ itemId: "1", fields: { Status: "Open" } }],
          { fieldPolicy: "track=Status;ignore=" },
        ),
      },
    });

  await withMockedFetch(
    withChangeRoutes(() =>
      json({ value: [spItem("1", { Title: "First", Status: "Open" })] })
    ),
    () =>
      model.methods.list_changes.execute(
        { ...CHANGE_ARGS },
        context as ExecCtx,
      ),
  );

  const data = written(getWrittenResources(), "listChanges").data;
  // Diffing across policies would report Title as newly set on every row.
  assertEquals(data.firstRun, true);
  assertEquals((data.totals as Record<string, number>).baselineRows, 0);
  assertEquals((data.modified as unknown[]).length, 0);
  assertStringIncludes(getLogsByLevel("warning")[0].message, "re-establishing");
});

Deno.test("list_changes: a baseline from a different list is not diffed", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
    storedResources: {
      [BASELINE_ALL]: storedBaseline(
        [{ itemId: "1", fields: { Title: "First", Status: "Open" } }],
        { listId: "00000000-0000-0000-0000-000000000000" },
      ),
    },
  });

  await withMockedFetch(
    withChangeRoutes(() =>
      json({ value: [spItem("1", { Title: "First", Status: "Open" })] })
    ),
    () =>
      model.methods.list_changes.execute(
        { ...CHANGE_ARGS },
        context as ExecCtx,
      ),
  );

  const data = written(getWrittenResources(), "listChanges").data;
  assertEquals(data.firstRun, true);
  assertEquals((data.totals as Record<string, number>).baselineRows, 0);
});

Deno.test("list_changes: the scope and the caller filter are grouped, and quotes escaped", async () => {
  const { context } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  const { calls } = await withMockedFetch(
    withChangeRoutes(() => json({ value: [] })),
    () =>
      model.methods.list_changes.execute(
        {
          ...CHANGE_ARGS,
          scopeColumn: "Status",
          scopeValue: "O'Brien",
          filter: "fields/Archived ne 1",
        },
        context as ExecCtx,
      ),
  );

  const itemsCall = calls.find((c) => c.url.includes("/items?"));
  assertExists(itemsCall);
  const query = new URL(itemsCall.url).searchParams;
  // Unparenthesised, an `or` inside the caller's filter would regroup against
  // the scope and widen the population.
  assertEquals(
    query.get("$filter"),
    "(fields/Status eq 'O''Brien') and (fields/Archived ne 1)",
  );
  assertEquals(
    itemsCall.headers["prefer"],
    "HonorNonIndexedQueriesWarningMayFailRandomly",
  );
});

Deno.test("list_changes: an unscoped run sends no filter and no Prefer header", async () => {
  const { context } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  const { calls } = await withMockedFetch(
    withChangeRoutes(() => json({ value: [] })),
    () =>
      model.methods.list_changes.execute(
        { ...CHANGE_ARGS },
        context as ExecCtx,
      ),
  );

  const itemsCall = calls.find((c) => c.url.includes("/items?"));
  assertExists(itemsCall);
  assertEquals(new URL(itemsCall.url).searchParams.get("$filter"), null);
  assertEquals(itemsCall.headers["prefer"], undefined);
});

Deno.test("list_changes: rows Graph returned outside the scope are dropped and the ignored filter reported", async () => {
  const { context, getWrittenResources, getLogsByLevel } =
    createModelTestContext({
      globalArgs: GLOBAL_ARGS,
      definition: DEFINITION,
    });

  await withMockedFetch(
    withChangeRoutes(() =>
      json({
        value: [
          spItem("1", { Title: "First", Status: "Open" }),
          spItem("2", { Title: "Second", Status: "Open" }),
          spItem("3", { Title: "Third", Status: "Closed" }),
        ],
      })
    ),
    () =>
      model.methods.list_changes.execute(
        { ...CHANGE_ARGS, scopeColumn: "Status", scopeValue: "Open" },
        context as ExecCtx,
      ),
  );

  const data = written(getWrittenResources(), "listChanges").data;
  assertEquals(data.filterHonored, false);
  const totals = data.totals as Record<string, number>;
  assertEquals(totals.scanned, 3);
  assertEquals(totals.inScope, 2);
  assertStringIncludes(
    getLogsByLevel("warning")[0].message,
    "ignored the filter",
  );
});

Deno.test("list_changes: a redacted span in the baseline is a blind spot, not a change", async () => {
  const link = "see https://teams.microsoft.com/l/";
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
    storedResources: {
      [BASELINE_ALL]: storedBaseline([
        {
          itemId: "1",
          fields: { Title: "First", Notes: `${link}***/thread` },
        },
      ]),
    },
  });

  await withMockedFetch(
    withChangeRoutes(() =>
      json({
        value: [
          spItem(
            "1",
            { Title: "First", Notes: `${link}19:abc-tenant-guid/thread` },
            "2026-09-19T12:00:00Z",
            "2.0",
          ),
        ],
      })
    ),
    () =>
      model.methods.list_changes.execute(
        { ...CHANGE_ARGS },
        context as ExecCtx,
      ),
  );

  const data = written(getWrittenResources(), "listChanges").data;
  // Without this the column reports a change on every run, forever.
  assertEquals((data.modified as unknown[]).length, 0);
  assertEquals((data.touched as unknown[]).length, 0);
  assertEquals(data.redactionBlindFields, ["Notes"]);
});

Deno.test("list_changes: an edit outside the redacted span is still caught", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
    storedResources: {
      [BASELINE_ALL]: storedBaseline([
        {
          itemId: "1",
          fields: {
            Title: "First",
            Notes: "see https://teams.microsoft.com/l/***/thread",
          },
        },
      ]),
    },
  });

  await withMockedFetch(
    withChangeRoutes(() =>
      json({
        value: [
          spItem(
            "1",
            {
              Title: "First",
              Notes: "REVIEWED https://teams.microsoft.com/l/19:abc/thread",
            },
            "2026-09-19T12:00:00Z",
            "2.0",
          ),
        ],
      })
    ),
    () =>
      model.methods.list_changes.execute(
        { ...CHANGE_ARGS },
        context as ExecCtx,
      ),
  );

  const data = written(getWrittenResources(), "listChanges").data;
  const modified = data.modified as Array<Record<string, unknown>>;
  assertEquals(modified.length, 1);
  const changes = modified[0].changes as Array<Record<string, unknown>>;
  assertEquals(changes[0].field, "Notes");
  assertEquals(changes[0].kind, "changed");
  assertEquals(data.redactionBlindFields, []);
});

Deno.test("list_changes: trackFields narrows the diff to the named columns", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
    storedResources: {
      [BASELINE_ALL]: storedBaseline(
        [{ itemId: "1", fields: { Status: "Open" } }],
        { fieldPolicy: "track=Status;ignore=" },
      ),
    },
  });

  await withMockedFetch(
    withChangeRoutes(() =>
      json({
        value: [
          spItem("1", { Title: "Renamed", Status: "Open", Notes: "new note" }),
        ],
      })
    ),
    () =>
      model.methods.list_changes.execute(
        { ...CHANGE_ARGS, trackFields: ["Status"] },
        context as ExecCtx,
      ),
  );

  const data = written(getWrittenResources(), "listChanges").data;
  assertEquals(data.firstRun, false);
  // Title and Notes both moved, and neither is tracked — so the row is a bump
  // with nothing visible behind it.
  assertEquals((data.modified as unknown[]).length, 0);
  assertEquals((data.touched as unknown[]).length, 1);
  const baseline = written(getWrittenResources(), "listBaseline").data;
  const rows = baseline.rows as Array<Record<string, unknown>>;
  assertEquals(rows[0].fields, { Status: "Open" });
});

Deno.test("list_changes: an ignored column does not report an edit", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
    storedResources: {
      [BASELINE_ALL]: storedBaseline(
        [{ itemId: "1", fields: { Title: "First", Status: "Open" } }],
        { fieldPolicy: "track=*;ignore=Notes" },
      ),
    },
  });

  await withMockedFetch(
    withChangeRoutes(() =>
      json({
        value: [
          spItem(
            "1",
            {
              Title: "First",
              Status: "Open",
              Notes: "rewritten by a workflow",
            },
            "2026-09-19T12:00:00Z",
            "2.0",
          ),
        ],
      })
    ),
    () =>
      model.methods.list_changes.execute(
        { ...CHANGE_ARGS, ignoreFields: ["Notes"] },
        context as ExecCtx,
      ),
  );

  const data = written(getWrittenResources(), "listChanges").data;
  assertEquals((data.modified as unknown[]).length, 0);
  assertEquals((data.touched as unknown[]).length, 0);
});

Deno.test("list_changes: consecutive runs diff against the run before", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  // First run: nothing stored, so this establishes the baseline through the
  // real write path rather than a fixture — which is what proves the baseline
  // a run writes is one the next run will accept.
  await withMockedFetch(
    withChangeRoutes(() =>
      json({
        value: [
          spItem(
            "1",
            { Title: "First", Status: "Open" },
            "2026-09-21T08:00:00Z",
          ),
        ],
      })
    ),
    () =>
      model.methods.list_changes.execute(
        { ...CHANGE_ARGS },
        context as ExecCtx,
      ),
  );

  await withMockedFetch(
    withChangeRoutes(() =>
      json({
        value: [
          spItem(
            "1",
            { Title: "First", Status: "Confirmed" },
            "2026-09-22T09:00:00Z",
            "4.0",
          ),
        ],
      })
    ),
    () =>
      model.methods.list_changes.execute(
        { ...CHANGE_ARGS },
        context as ExecCtx,
      ),
  );

  const reports = getWrittenResources().filter((r) =>
    r.specName === "listChanges"
  );
  assertEquals(reports.length, 2);
  const second = reports[1].data;
  assertEquals(second.firstRun, false);
  assertEquals(second.since, reports[0].data.fetchedAt);
  assertNotEquals(second.sinceGapHours, null);
  const modified = second.modified as Array<Record<string, unknown>>;
  assertEquals(modified.length, 1);
  assertEquals(
    (modified[0].changes as Array<Record<string, unknown>>)[0].to,
    "Confirmed",
  );
});

Deno.test("list_changes: two scopes on one list keep separate baselines", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  const run = (scopeValue: string) =>
    withMockedFetch(
      withChangeRoutes(() =>
        json({ value: [spItem("1", { Title: "First", Status: scopeValue })] })
      ),
      () =>
        model.methods.list_changes.execute(
          { ...CHANGE_ARGS, scopeColumn: "Status", scopeValue },
          context as ExecCtx,
        ),
    );

  await run("Open");
  await run("Closed");

  const baselines = getWrittenResources()
    .filter((r) => r.specName === "listBaseline")
    .map((r) => r.name);
  assertEquals(baselines, [
    "baseline-project-tracker-status-open",
    "baseline-project-tracker-status-closed",
  ]);

  // The second scope has its own population, so it must not read the first
  // one's baseline and report the whole list as both added and removed.
  const reports = getWrittenResources().filter((r) =>
    r.specName === "listChanges"
  );
  assertEquals(reports[1].data.firstRun, true);
  assertEquals((reports[1].data.removed as unknown[]).length, 0);
});

Deno.test("list_changes: a probe that fails holds the baseline back", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
    storedResources: {
      [BASELINE_ALL]: storedBaseline([
        { itemId: "1", fields: { Title: "First", Status: "Open" } },
        { itemId: "9", fields: { Title: "Ninth", Status: "Open" } },
      ]),
    },
  });

  await withMockedFetch(
    withChangeRoutes(
      () =>
        json({
          value: [
            spItem(
              "1",
              { Title: "First", Status: "Open" },
              "2026-09-19T12:00:00Z",
              "2.0",
            ),
          ],
        }),
      (itemId) =>
        itemId === "9"
          ? graphError("serviceNotAvailable", "backend is unwell", 500)
          : undefined,
    ),
    () =>
      model.methods.list_changes.execute(
        { ...CHANGE_ARGS },
        context as ExecCtx,
      ),
  );

  const data = written(getWrittenResources(), "listChanges").data;
  const removed = data.removed as Array<Record<string, unknown>>;
  // A failed probe is doubt, not a deletion.
  assertEquals(removed[0].fate, "unknown");
  assertEquals((data.totals as Record<string, number>).probeFailed, 1);
  assertEquals(data.baselineUpdated, false);
  assertStringIncludes(String(data.baselineHeldReason), "probe(s) failed");
  // Advancing here would drop row 9 from the baseline on the strength of a
  // 500, and it would never be mentioned again.
  assertEquals(getWrittenResources().length, 1);
});

Deno.test("list_changes: probes skipped past the budget do not stall the baseline", async () => {
  // 40 rows leave the scope at once on a 200-row baseline: under the alarm
  // share, so these are real removals, but past the 30-probe budget.
  const rows = Array.from({ length: 200 }, (_, i) => ({
    itemId: String(i + 1),
    fields: { Title: `Row ${i + 1}`, Status: "Open" },
  }));

  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
    storedResources: { [BASELINE_ALL]: storedBaseline(rows) },
  });

  await withMockedFetch(
    withChangeRoutes(
      () =>
        json({
          value: rows.slice(40).map((row) =>
            spItem(
              row.itemId,
              { Title: row.fields.Title, Status: "Open" },
              "2026-09-19T12:00:00Z",
              "2.0",
            )
          ),
        }),
      () => graphError("itemNotFound", "Item does not exist", 404),
    ),
    () =>
      model.methods.list_changes.execute(
        { ...CHANGE_ARGS },
        context as ExecCtx,
      ),
  );

  const data = written(getWrittenResources(), "listChanges").data;
  assertEquals(data.suspectRemoval, false);
  const removed = data.removed as Array<Record<string, unknown>>;
  assertEquals(removed.length, 40);
  assertEquals(removed.filter((r) => r.fate === "deleted").length, 30);
  assertEquals(removed.filter((r) => r.fate === "unprobed").length, 10);
  assertEquals((data.totals as Record<string, number>).probeFailed, 0);
  // An unprobed row must not hold the baseline: past the budget it would hold
  // it on every run and the baseline would never advance again.
  assertEquals(data.baselineUpdated, true);
});

Deno.test("list_changes: what it writes matches the resource schemas exactly", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
    storedResources: {
      [BASELINE_ALL]: storedBaseline([
        { itemId: "1", fields: { Title: "First", Status: "Open" } },
        { itemId: "2", fields: { Title: "Second", Status: "Open" } },
        {
          itemId: "8",
          fields: { Title: "Eighth", Notes: "see /l/***/thread" },
        },
        { itemId: "9", fields: { Title: "Ninth", Status: "Open" } },
      ]),
    },
  });

  await withMockedFetch(
    withChangeRoutes(
      () =>
        json({
          value: [
            // modified
            spItem("1", { Title: "First", Status: "Confirmed" }),
            // touched
            spItem("2", { Title: "Second", Status: "Open" }),
            // redaction-blind
            spItem(
              "8",
              { Title: "Eighth", Notes: "see /l/19:abc/thread" },
              "2026-09-19T12:00:00Z",
              "2.0",
            ),
            // added
            spItem("3", { Title: "Third", Status: "New" }),
          ],
        }),
      (itemId) =>
        itemId === "9"
          ? graphError("itemNotFound", "Item does not exist", 404)
          : undefined,
    ),
    () =>
      model.methods.list_changes.execute(
        { ...CHANGE_ARGS },
        context as ExecCtx,
      ),
  );

  // Every branch of the report is populated, so this covers the whole schema
  // rather than the happy path.
  const changes = written(getWrittenResources(), "listChanges").data;
  assertEquals((changes.added as unknown[]).length, 1);
  assertEquals((changes.modified as unknown[]).length, 1);
  assertEquals((changes.touched as unknown[]).length, 1);
  assertEquals((changes.removed as unknown[]).length, 1);
  assertEquals(changes.redactionBlindFields, ["Notes"]);

  // parse() proves every schema field is present and well-typed; comparing the
  // key sets proves nothing extra was written, since parse() silently strips
  // what the schema does not declare.
  for (
    const [spec, data] of [
      ["listChanges", changes],
      ["listBaseline", written(getWrittenResources(), "listBaseline").data],
    ] as const
  ) {
    const schema = model.resources[spec].schema;
    const parsed = schema.parse(data) as Record<string, unknown>;
    assertEquals(
      Object.keys(data).sort(),
      Object.keys(parsed).sort(),
      `${spec} wrote fields its schema does not declare`,
    );
    assertEquals(
      Object.keys(data.totals ?? {}).sort(),
      Object.keys((parsed.totals ?? {}) as Record<string, unknown>).sort(),
      `${spec} totals wrote fields its schema does not declare`,
    );
  }
});

Deno.test("list_changes: contextFields ride along on a row whose other columns moved", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
    storedResources: {
      [BASELINE_ALL]: storedBaseline([
        {
          itemId: "1",
          fields: { Title: "First", Status: "Open", Notes: "before" },
        },
      ]),
    },
  });

  await withMockedFetch(
    withChangeRoutes(() =>
      json({
        value: [
          spItem("1", { Title: "First", Status: "Open", Notes: "after" }),
        ],
      })
    ),
    () =>
      model.methods.list_changes.execute(
        { ...CHANGE_ARGS, contextFields: ["Status", "Title"] },
        context as ExecCtx,
      ),
  );

  const data = written(getWrittenResources(), "listChanges").data;
  const modified = data.modified as Array<Record<string, unknown>>;
  // Only Notes moved, so without context the report could not say the row is
  // still Open.
  assertEquals(
    (modified[0].changes as Array<Record<string, unknown>>).map((c) => c.field),
    ["Notes"],
  );
  assertEquals(modified[0].context, { Status: "Open", Title: "First" });
  // The baseline keeps none of it — context is read fresh every run.
  const rows = written(getWrittenResources(), "listBaseline").data
    .rows as Array<Record<string, unknown>>;
  assertEquals("context" in rows[0], false);
});

Deno.test("list_changes: summarizeColumn counts the watched population, busiest first", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  await withMockedFetch(
    withChangeRoutes(() =>
      json({
        value: [
          spItem("1", { Title: "a", Status: "Open" }),
          spItem("2", { Title: "b", Status: "Confirmed" }),
          spItem("3", { Title: "c", Status: "Open" }),
          spItem("4", { Title: "d" }),
        ],
      })
    ),
    () =>
      model.methods.list_changes.execute(
        { ...CHANGE_ARGS, summarizeColumn: "Status" },
        context as ExecCtx,
      ),
  );

  const data = written(getWrittenResources(), "listChanges").data;
  assertEquals(data.summarizeColumn, "Status");
  assertEquals(data.byValue, [
    { value: "Open", count: 2 },
    { value: "(empty)", count: 1 },
    { value: "Confirmed", count: 1 },
  ]);
});

Deno.test("list_changes: a display name as summarizeColumn is refused before any network call", async () => {
  const { context } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  const { calls } = await withMockedFetch(
    () => graphError("unexpectedCall", "should not be reached", 500),
    async () => {
      const err = await assertRejects(
        () =>
          model.methods.list_changes.execute(
            { ...CHANGE_ARGS, summarizeColumn: "Request Status" },
            context as ExecCtx,
          ),
        Error,
      );
      // Counting every row under "(empty)" is a wrong answer that looks right.
      assertStringIncludes(err.message, "internal column name");
    },
  );

  assertEquals(calls.length, 0);
});

Deno.test("list_changes: the summary counts only rows inside the scope", async () => {
  const { context, getWrittenResources } = createModelTestContext({
    globalArgs: GLOBAL_ARGS,
    definition: DEFINITION,
  });

  await withMockedFetch(
    withChangeRoutes(() =>
      json({
        value: [
          spItem("1", { Title: "a", Status: "Open", Owner: "Ana" }),
          spItem("2", { Title: "b", Status: "Open", Owner: "Ana" }),
          // Graph ignored the unindexed filter and returned this one too.
          spItem("3", { Title: "c", Status: "Closed", Owner: "Bo" }),
        ],
      })
    ),
    () =>
      model.methods.list_changes.execute(
        {
          ...CHANGE_ARGS,
          scopeColumn: "Status",
          scopeValue: "Open",
          summarizeColumn: "Owner",
        },
        context as ExecCtx,
      ),
  );

  const data = written(getWrittenResources(), "listChanges").data;
  assertEquals(data.filterHonored, false);
  // Bo's row is outside the scope, so it is not part of this population's mix.
  assertEquals(data.byValue, [{ value: "Ana", count: 2 }]);
});
