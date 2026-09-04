// SharePoint Lists Model Tests
// SPDX-License-Identifier: Apache-2.0

import {
  assertEquals,
  assertExists,
  assertRejects,
  assertStringIncludes,
} from "jsr:@std/assert@1.0.19";
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

Deno.test("sharepoint-lists model: has correct type", () => {
  assertEquals(model.type, "@aspec451/microsoft/sharepoint-lists");
});

Deno.test("sharepoint-lists model: has valid version format", () => {
  const versionPattern = /^\d{4}\.\d{2}\.\d{2}\.\d+$/;
  assertEquals(versionPattern.test(model.version), true);
});

Deno.test("sharepoint-lists model: has globalArguments with credentials and site", () => {
  assertExists(model.globalArguments);
  const shape = model.globalArguments.shape;
  assertExists(shape.tenantId);
  assertExists(shape.clientId);
  assertExists(shape.refreshToken);
  assertExists(shape.siteHostPath);
  assertExists(shape.scopes);
});

Deno.test("sharepoint-lists model: refreshToken is marked sensitive", () => {
  const meta = model.globalArguments.shape.refreshToken.meta();
  assertEquals(meta?.sensitive, true);
});

Deno.test("sharepoint-lists model: has required resources", () => {
  assertExists(model.resources);
  assertExists(model.resources.lists);
  assertExists(model.resources.listSchema);
  assertExists(model.resources.listItems);
  assertExists(model.resources.itemWrite);
  assertExists(model.resources.bootstrap);
});

Deno.test("sharepoint-lists model: has required methods", () => {
  assertExists(model.methods);
  assertExists(model.methods.bootstrap);
  assertExists(model.methods.list_lists);
  assertExists(model.methods.get_list);
  assertExists(model.methods.list_items);
  assertExists(model.methods.create_item);
  assertExists(model.methods.update_item);
  assertExists(model.methods.delete_item);
});

Deno.test("sharepoint-lists model: list_items has filter and paging arguments", () => {
  const shape = model.methods.list_items.arguments.shape;
  assertExists(shape.list);
  assertExists(shape.filter);
  assertExists(shape.orderBy);
  assertExists(shape.select);
  assertExists(shape.top);
  assertExists(shape.maxPages);
});

Deno.test("sharepoint-lists model: write methods require a fields object", () => {
  assertExists(model.methods.create_item.arguments.shape.fields);
  assertExists(model.methods.update_item.arguments.shape.fields);
  assertExists(model.methods.update_item.arguments.shape.itemId);
});

Deno.test("sharepoint-lists model: delete_item has a confirm guard argument", () => {
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
        { list: "Project Tracker", includeHidden: false },
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
            { list: "Nope", includeHidden: false },
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
