/**
 * SharePoint Lists Model — read and write SharePoint list items via Graph API.
 *
 * A sibling to the Teams model in this family: same public-client app
 * registration, same device code flow, same shared `_lib` helpers. Where the
 * existing SharePoint extensions in the registry cover document libraries
 * (`/drive/...`) and are read-only, this covers the list surface
 * (`/sites/{site}/lists/...`) and can write.
 *
 * @module
 */

import { z } from "npm:zod@4";
import {
  DEFAULT_SHAREPOINT_SCOPES,
  initiateDeviceCode,
  pollDeviceCode,
  refreshAccessToken,
} from "./_lib/auth.ts";
import { graphRequest, graphRequestPaginated } from "./_lib/graph.ts";

const EXTENSION_NAME = "@aspec451/microsoft/sharepoint-lists";

/**
 * Graph rejects `$filter` and `$orderby` on list-item columns that SharePoint
 * has not indexed unless the caller opts in to the failure risk with this
 * header. Without it a perfectly valid filter on an unindexed column comes
 * back as a flat 400, which reads like a syntax error rather than a missing
 * index — so it is sent whenever a filter or sort is supplied.
 */
const HONOR_NON_INDEXED = "HonorNonIndexedQueriesWarningMayFailRandomly";

// =============================================================================
// Schemas
// =============================================================================

const GlobalArgsSchema = z.object({
  tenantId: z.string().min(1).describe(
    "Entra tenant GUID (e.g. e9b2b7ba-b238-42a9-b271-2adfc82da650)",
  ),
  clientId: z.string().min(1).describe(
    "Azure public client app registration ID",
  ),
  refreshToken: z.string().min(1).meta({ sensitive: true }).describe(
    "OAuth2 refresh token obtained via bootstrap. Re-run bootstrap if expired.",
  ),
  siteHostPath: z.string().min(1).describe(
    "Graph site locator in host:path form, e.g. contoso.sharepoint.com:/sites/Clients",
  ),
  scopes: z.string().default(DEFAULT_SHAREPOINT_SCOPES).describe(
    "Space-separated delegated scopes. Default includes Sites.ReadWrite.All; " +
      "use 'offline_access User.Read Sites.Read.All' for a read-only deployment. " +
      "Changing this requires re-running bootstrap — a refresh token is bound " +
      "to the scopes it was issued with.",
  ),
});

type GlobalArgs = z.infer<typeof GlobalArgsSchema>;

const MetaFields = {
  fetchedAt: z.string().describe("ISO 8601 timestamp when the call ran"),
  durationMs: z.number().optional().describe(
    "Method execution duration in milliseconds",
  ),
  collectedBy: z.string().optional().describe(
    "Extension that collected this data",
  ),
};

const SiteFields = {
  siteId: z.string().describe("Resolved Graph site ID"),
  siteHostPath: z.string().describe("Site locator the model is configured for"),
};

const ListSummarySchema = z.object({
  id: z.string().describe("List GUID"),
  name: z.string().describe("List URL name (use this or the GUID as `list`)"),
  displayName: z.string().describe("List display name as shown in SharePoint"),
  description: z.string().nullable().optional().describe("List description"),
  template: z.string().nullable().optional().describe(
    "List template, e.g. genericList, documentLibrary, events",
  ),
  hidden: z.boolean().optional().describe(
    "Whether SharePoint hides this list from the site contents page",
  ),
  webUrl: z.string().nullable().optional().describe("Browser URL for the list"),
  lastModifiedDateTime: z.string().nullable().optional().describe(
    "Timestamp the list was last modified",
  ),
});

const ListsSchema = z.object({
  ...SiteFields,
  siteName: z.string().nullable().optional().describe("Site display name"),
  lists: z.array(ListSummarySchema).describe("Lists on the site"),
  totalFetched: z.number().describe("Number of lists returned"),
  includedHidden: z.boolean().describe(
    "Whether hidden/system lists were included",
  ),
  truncated: z.boolean().describe("Whether the page cap cut the walk short"),
  ...MetaFields,
});

const ColumnSchema = z.object({
  name: z.string().describe(
    "Internal column name — this is the key to use in create_item/update_item fields",
  ),
  displayName: z.string().nullable().optional().describe(
    "Column name as shown in SharePoint",
  ),
  type: z.string().describe(
    "Derived column type, e.g. text, number, choice, dateTime, lookup, person",
  ),
  required: z.boolean().optional().describe("Whether the column is required"),
  readOnly: z.boolean().optional().describe(
    "Whether the column rejects writes — omit these from create_item/update_item",
  ),
  hidden: z.boolean().optional().describe("Whether the column is hidden"),
  choices: z.array(z.string()).optional().describe(
    "Allowed values, for choice columns",
  ),
});

const ListSchemaResourceSchema = z.object({
  ...SiteFields,
  list: ListSummarySchema.describe("The resolved list"),
  columns: z.array(ColumnSchema).describe(
    "Column definitions, keyed by internal name",
  ),
  writableColumns: z.array(z.string()).describe(
    "Internal names of columns that accept writes",
  ),
  ...MetaFields,
});

const ListItemSchema = z.object({
  id: z.string().describe("List item ID"),
  fields: z.record(z.string(), z.unknown()).describe(
    "Column values keyed by internal column name",
  ),
  createdDateTime: z.string().nullable().optional().describe(
    "Timestamp the item was created",
  ),
  lastModifiedDateTime: z.string().nullable().optional().describe(
    "Timestamp the item was last modified",
  ),
  webUrl: z.string().nullable().optional().describe("Browser URL for the item"),
  createdBy: z.string().nullable().optional().describe(
    "Display name of the creating user",
  ),
  lastModifiedBy: z.string().nullable().optional().describe(
    "Display name of the last user to modify the item",
  ),
});

const ListItemsSchema = z.object({
  ...SiteFields,
  listId: z.string().describe("Resolved list GUID"),
  listName: z.string().describe("Display name of the list"),
  items: z.array(ListItemSchema).describe("List items returned"),
  totalFetched: z.number().describe("Number of items returned"),
  truncated: z.boolean().describe(
    "Whether the fetch hit maxPages before exhausting the list",
  ),
  filter: z.string().nullable().optional().describe("OData filter applied"),
  orderBy: z.string().nullable().optional().describe("OData sort applied"),
  ...MetaFields,
});

const ItemWriteSchema = z.object({
  ...SiteFields,
  operation: z.enum(["create", "update", "delete"]).describe(
    "Which write was performed",
  ),
  listId: z.string().describe("Resolved list GUID"),
  listName: z.string().describe("Display name of the list"),
  itemId: z.string().describe("List item ID affected"),
  fields: z.record(z.string(), z.unknown()).describe(
    "Resulting field values for create/update; the deleted item's values for delete",
  ),
  webUrl: z.string().nullable().optional().describe("Browser URL for the item"),
  ...MetaFields,
});

const BootstrapResultSchema = z.object({
  status: z.string().describe("Authentication result status"),
  message: z.string().describe("Human-readable guidance for the caller"),
  scopes: z.string().optional().describe("Scopes the token was issued for"),
  refreshToken: z.string().meta({ sensitive: true }).optional().describe(
    "OAuth2 refresh token to store in the vault",
  ),
  ...MetaFields,
  fetchedAt: z.string().optional().describe(
    "ISO 8601 timestamp when the call ran",
  ),
});

// =============================================================================
// Graph response shapes
// =============================================================================

interface GraphSite {
  id: string;
  displayName?: string | null;
  name?: string | null;
  webUrl?: string | null;
}

interface GraphList {
  id: string;
  name?: string | null;
  displayName?: string | null;
  description?: string | null;
  webUrl?: string | null;
  lastModifiedDateTime?: string | null;
  list?: { template?: string | null; hidden?: boolean } | null;
}

interface GraphColumn {
  name?: string | null;
  displayName?: string | null;
  description?: string | null;
  required?: boolean;
  readOnly?: boolean;
  hidden?: boolean;
  columnGroup?: string | null;
  choice?: { choices?: string[] } | null;
  [facet: string]: unknown;
}

interface GraphIdentitySet {
  user?: { displayName?: string | null } | null;
}

interface GraphListItem {
  id: string;
  createdDateTime?: string | null;
  lastModifiedDateTime?: string | null;
  webUrl?: string | null;
  createdBy?: GraphIdentitySet | null;
  lastModifiedBy?: GraphIdentitySet | null;
  fields?: Record<string, unknown> | null;
}

// =============================================================================
// Helpers
// =============================================================================

/** The Graph column facets that identify a column's type. */
const COLUMN_TYPE_FACETS = [
  "text",
  "number",
  "boolean",
  "dateTime",
  "choice",
  "currency",
  "lookup",
  "personOrGroup",
  "hyperlinkOrPicture",
  "calculated",
  "geolocation",
  "term",
  "thumbnail",
  "contentApprovalStatus",
] as const;

async function getAccessToken(globalArgs: GlobalArgs): Promise<string> {
  const tokens = await refreshAccessToken({
    tenantId: globalArgs.tenantId,
    clientId: globalArgs.clientId,
    refreshToken: globalArgs.refreshToken,
    scopes: globalArgs.scopes,
  });
  return tokens.access_token;
}

/**
 * Run a Graph API call, and on failure rethrow with the operation being
 * attempted and its identifying arguments prepended, preserving the original
 * error (GraphApiError's status/code or a network error) as `cause`.
 */
async function withGraphContext<T>(
  op: string,
  fn: () => Promise<T>,
): Promise<T> {
  try {
    return await fn();
  } catch (e) {
    throw new Error(
      `${op}: ${e instanceof Error ? e.message : String(e)}`,
      { cause: e },
    );
  }
}

/** Turn an identifier into a short, file-name-safe data instance name. */
function slug(value: string, fallback = "item"): string {
  return value
    .replace(/[^a-zA-Z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 48)
    .toLowerCase() || fallback;
}

function isGuid(value: string): boolean {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i
    .test(value);
}

/** Resolve the configured host:path locator to a Graph site ID. */
async function resolveSite(
  accessToken: string,
  siteHostPath: string,
): Promise<GraphSite> {
  return await withGraphContext(
    `Failed to resolve site "${siteHostPath}"`,
    () =>
      graphRequest<GraphSite>(
        accessToken,
        "GET",
        `/sites/${siteHostPath}?$select=id,name,displayName,webUrl`,
      ),
  );
}

function mapList(raw: GraphList): z.infer<typeof ListSummarySchema> {
  return {
    id: raw.id,
    name: raw.name ?? "",
    displayName: raw.displayName ?? raw.name ?? "",
    description: raw.description ?? null,
    template: raw.list?.template ?? null,
    hidden: raw.list?.hidden ?? false,
    webUrl: raw.webUrl ?? null,
    lastModifiedDateTime: raw.lastModifiedDateTime ?? null,
  };
}

const LIST_SELECT =
  "id,name,displayName,description,webUrl,lastModifiedDateTime,list";

/**
 * Resolve a user-supplied list reference to the real list.
 *
 * Three things can be passed and all three are common: the list GUID, the list
 * URL name (Graph's `name`, the segment in the browser URL), and the display
 * name shown in the SharePoint UI. Graph's `/lists/{id-or-name}` accepts only
 * the first two, and the display name diverges from the URL name whenever a
 * list has been renamed — which is most of them. So a miss on the direct
 * lookup falls back to enumerating the site's lists and matching displayName
 * case-insensitively, rather than reporting a list that plainly exists as
 * missing.
 */
async function resolveList(
  accessToken: string,
  siteId: string,
  list: string,
): Promise<GraphList> {
  if (isGuid(list)) {
    return await withGraphContext(
      `Failed to fetch list "${list}"`,
      () =>
        graphRequest<GraphList>(
          accessToken,
          "GET",
          `/sites/${siteId}/lists/${list}?$select=${LIST_SELECT}`,
        ),
    );
  }

  try {
    return await graphRequest<GraphList>(
      accessToken,
      "GET",
      `/sites/${siteId}/lists/${
        encodeURIComponent(list)
      }?$select=${LIST_SELECT}`,
    );
  } catch {
    // Fall through to a display-name search.
  }

  const all = await withGraphContext(
    `Failed to enumerate lists while resolving "${list}"`,
    () =>
      graphRequestPaginated<GraphList>(
        accessToken,
        `/sites/${siteId}/lists`,
        { "$select": LIST_SELECT },
      ),
  );

  const target = list.toLowerCase();
  const match = all.items.find((l) =>
    (l.displayName ?? "").toLowerCase() === target ||
    (l.name ?? "").toLowerCase() === target
  );

  if (!match) {
    const available = all.items
      .map((l) => l.displayName ?? l.name ?? l.id)
      .slice(0, 25)
      .join(", ");
    throw new Error(
      `No list on this site matches "${list}" by GUID, URL name, or display ` +
        `name. Available lists: ${available || "(none)"}`,
    );
  }

  return match;
}

function mapColumn(raw: GraphColumn): z.infer<typeof ColumnSchema> {
  const facet = COLUMN_TYPE_FACETS.find((f) => raw[f] !== undefined);
  return {
    name: raw.name ?? "",
    displayName: raw.displayName ?? null,
    type: facet ?? "unknown",
    required: raw.required ?? false,
    readOnly: raw.readOnly ?? false,
    hidden: raw.hidden ?? false,
    ...(raw.choice?.choices ? { choices: raw.choice.choices } : {}),
  };
}

function mapItem(raw: GraphListItem): z.infer<typeof ListItemSchema> {
  return {
    id: raw.id,
    fields: raw.fields ?? {},
    createdDateTime: raw.createdDateTime ?? null,
    lastModifiedDateTime: raw.lastModifiedDateTime ?? null,
    webUrl: raw.webUrl ?? null,
    createdBy: raw.createdBy?.user?.displayName ?? null,
    lastModifiedBy: raw.lastModifiedBy?.user?.displayName ?? null,
  };
}

/** The swamp method context these methods use. */
interface Context {
  globalArgs: GlobalArgs;
  writeResource: (
    specName: string,
    name: string,
    data: Record<string, unknown>,
  ) => Promise<{ name: string }>;
  logger: {
    info: (msg: string, props?: Record<string, unknown>) => void;
    warn: (msg: string, props?: Record<string, unknown>) => void;
  };
}

// =============================================================================
// Model Definition
// =============================================================================

/** SharePoint list read/write model via Graph API. */
export const model = {
  type: "@aspec451/microsoft/sharepoint-lists",
  version: "2026.09.03.1",
  globalArguments: GlobalArgsSchema,

  resources: {
    lists: {
      description: "Lists on the configured SharePoint site",
      schema: ListsSchema,
      lifetime: "15m" as const,
      garbageCollection: 5,
    },
    listSchema: {
      description:
        "One list's column definitions — the internal names writes must use",
      schema: ListSchemaResourceSchema,
      lifetime: "15m" as const,
      garbageCollection: 10,
    },
    listItems: {
      description: "Items from a SharePoint list",
      schema: ListItemsSchema,
      lifetime: "15m" as const,
      garbageCollection: 20,
    },
    itemWrite: {
      description:
        "Result of a create, update, or delete against a single list item",
      schema: ItemWriteSchema,
      lifetime: "infinite" as const,
      garbageCollection: 50,
    },
    bootstrap: {
      description: "Device code flow authentication result",
      schema: BootstrapResultSchema,
      lifetime: "infinite" as const,
      garbageCollection: 3,
    },
  },

  methods: {
    bootstrap: {
      description:
        "Authenticate via device code flow. Displays a user code and " +
        "verification URL, then polls until authentication completes. " +
        "Outputs the refresh token to store in the vault.",
      arguments: z.object({}),
      execute: async (_args: Record<string, never>, context: Context) => {
        const startMs = Date.now();
        const { tenantId, clientId, scopes } = context.globalArgs;

        const deviceCode = await initiateDeviceCode(
          tenantId,
          clientId,
          scopes,
        );

        context.logger.info(
          "Device code flow initiated. Go to {uri} and enter code: {code}",
          { uri: deviceCode.verification_uri, code: deviceCode.user_code },
        );

        const tokens = await pollDeviceCode(
          tenantId,
          clientId,
          deviceCode.device_code,
          deviceCode.interval,
          deviceCode.expires_in * 1000,
        );

        context.logger.info("Authentication successful");

        const handle = await context.writeResource("bootstrap", "main", {
          status: "authenticated",
          message:
            "Store the refreshToken in your vault. It expires after 90 days of inactivity.",
          scopes: tokens.scope ?? scopes,
          refreshToken: tokens.refresh_token,
          durationMs: Date.now() - startMs,
          collectedBy: EXTENSION_NAME,
          fetchedAt: new Date().toISOString(),
        });

        return { dataHandles: [handle] };
      },
    },

    list_lists: {
      description:
        "Enumerate the lists on the configured site. Start here — the output " +
        "gives the GUID, URL name, and display name that every other method " +
        "accepts as its `list` argument.",
      arguments: z.object({
        includeHidden: z.boolean().default(false).describe(
          "Include hidden/system lists (Form Templates, Style Library, and the like)",
        ),
        maxPages: z.number().int().min(1).max(50).default(20).describe(
          "Cap on @odata.nextLink pages followed",
        ),
      }),
      execute: async (
        args: { includeHidden: boolean; maxPages: number },
        context: Context,
      ) => {
        const startMs = Date.now();
        const accessToken = await getAccessToken(context.globalArgs);
        const site = await resolveSite(
          accessToken,
          context.globalArgs.siteHostPath,
        );

        const result = await withGraphContext(
          `Failed to list lists on site "${site.id}"`,
          () =>
            graphRequestPaginated<GraphList>(
              accessToken,
              `/sites/${site.id}/lists`,
              { "$select": LIST_SELECT },
              undefined,
              fetch,
              args.maxPages,
            ),
        );

        const lists = result.items
          .map(mapList)
          .filter((l) => args.includeHidden || !l.hidden);

        const handle = await context.writeResource("lists", "main", {
          siteId: site.id,
          siteHostPath: context.globalArgs.siteHostPath,
          siteName: site.displayName ?? site.name ?? null,
          lists,
          totalFetched: lists.length,
          includedHidden: args.includeHidden,
          truncated: result.truncated,
          fetchedAt: new Date().toISOString(),
          durationMs: Date.now() - startMs,
          collectedBy: EXTENSION_NAME,
        });

        context.logger.info("Found {count} lists on {site}", {
          count: lists.length,
          site: site.displayName ?? context.globalArgs.siteHostPath,
        });
        return { dataHandles: [handle] };
      },
    },

    get_list: {
      description:
        "Fetch one list's column definitions. Run this before writing: " +
        "create_item and update_item key their fields by INTERNAL column name " +
        "(e.g. Title, Status, AssignedTo), which is often not the display name " +
        "shown in SharePoint, and read-only columns must be omitted entirely.",
      arguments: z.object({
        list: z.string().min(1, "list must not be empty").describe(
          "List GUID, URL name, or display name",
        ),
        includeHidden: z.boolean().default(false).describe(
          "Include hidden columns in the output",
        ),
      }),
      execute: async (
        args: { list: string; includeHidden: boolean },
        context: Context,
      ) => {
        const startMs = Date.now();
        const accessToken = await getAccessToken(context.globalArgs);
        const site = await resolveSite(
          accessToken,
          context.globalArgs.siteHostPath,
        );
        const list = await resolveList(accessToken, site.id, args.list);

        const result = await withGraphContext(
          `Failed to fetch columns for list "${args.list}"`,
          () =>
            graphRequestPaginated<GraphColumn>(
              accessToken,
              `/sites/${site.id}/lists/${list.id}/columns`,
            ),
        );

        const columns = result.items
          .map(mapColumn)
          .filter((c) => args.includeHidden || !c.hidden);

        const handle = await context.writeResource(
          "listSchema",
          slug(list.displayName ?? list.name ?? list.id, "list"),
          {
            siteId: site.id,
            siteHostPath: context.globalArgs.siteHostPath,
            list: mapList(list),
            columns,
            writableColumns: columns
              .filter((c) => !c.readOnly)
              .map((c) => c.name),
            fetchedAt: new Date().toISOString(),
            durationMs: Date.now() - startMs,
            collectedBy: EXTENSION_NAME,
          },
        );

        context.logger.info(
          "List {name}: {total} columns, {writable} writable",
          {
            name: list.displayName ?? args.list,
            total: columns.length,
            writable: columns.filter((c) => !c.readOnly).length,
          },
        );
        return { dataHandles: [handle] };
      },
    },

    list_items: {
      description:
        "Fetch items from a list, with column values expanded. Supports OData " +
        "$filter and $orderby against internal column names, e.g. " +
        "filter=\"fields/Status eq 'Open'\".",
      arguments: z.object({
        list: z.string().min(1, "list must not be empty").describe(
          "List GUID, URL name, or display name",
        ),
        filter: z.string().optional().describe(
          "OData $filter over fields, e.g. fields/Status eq 'Open'. " +
            "Filtering an unindexed column is best-effort — see truncated/warnings.",
        ),
        orderBy: z.string().optional().describe(
          "OData $orderby, e.g. fields/Created desc",
        ),
        select: z.string().optional().describe(
          "Comma-separated internal column names to return; omit for all columns",
        ),
        top: z.number().int().min(1).max(5000).default(200).describe(
          "Page size requested from Graph",
        ),
        maxPages: z.number().int().min(1).max(100).default(20).describe(
          "Cap on @odata.nextLink pages followed",
        ),
      }),
      execute: async (
        args: {
          list: string;
          filter?: string;
          orderBy?: string;
          select?: string;
          top: number;
          maxPages: number;
        },
        context: Context,
      ) => {
        const startMs = Date.now();
        const accessToken = await getAccessToken(context.globalArgs);
        const site = await resolveSite(
          accessToken,
          context.globalArgs.siteHostPath,
        );
        const list = await resolveList(accessToken, site.id, args.list);

        const params: Record<string, string> = {
          "$expand": args.select ? `fields($select=${args.select})` : "fields",
          "$top": String(args.top),
        };
        if (args.filter) params["$filter"] = args.filter;
        if (args.orderBy) params["$orderby"] = args.orderBy;

        const headers = (args.filter || args.orderBy)
          ? { "Prefer": HONOR_NON_INDEXED }
          : undefined;

        const result = await withGraphContext(
          `Failed to list items in "${args.list}"`,
          () =>
            graphRequestPaginated<GraphListItem>(
              accessToken,
              `/sites/${site.id}/lists/${list.id}/items`,
              params,
              headers,
              fetch,
              args.maxPages,
            ),
        );

        const items = result.items.map(mapItem);

        if (result.truncated) {
          context.logger.warn(
            "Hit the maxPages cap ({cap}) with more items available — " +
              "{count} is a floor, not the whole list. Raise maxPages or narrow with filter.",
            { cap: args.maxPages, count: items.length },
          );
        }

        const handle = await context.writeResource(
          "listItems",
          slug(list.displayName ?? list.name ?? list.id, "list"),
          {
            siteId: site.id,
            siteHostPath: context.globalArgs.siteHostPath,
            listId: list.id,
            listName: list.displayName ?? list.name ?? list.id,
            items,
            totalFetched: items.length,
            truncated: result.truncated,
            filter: args.filter ?? null,
            orderBy: args.orderBy ?? null,
            fetchedAt: new Date().toISOString(),
            durationMs: Date.now() - startMs,
            collectedBy: EXTENSION_NAME,
          },
        );

        context.logger.info(
          "Fetched {count} item(s) from {name}{caveat}",
          {
            count: items.length,
            name: list.displayName ?? args.list,
            caveat: result.truncated ? " — TRUNCATED at the maxPages cap" : "",
          },
        );
        return { dataHandles: [handle] };
      },
    },

    create_item: {
      description:
        "Create one list item. Fields are keyed by INTERNAL column name — run " +
        "get_list first to see them and to check which columns are writable. " +
        "Runs as the signed-in user, so list validation and item-level " +
        "permissions apply.",
      arguments: z.object({
        list: z.string().min(1, "list must not be empty").describe(
          "List GUID, URL name, or display name",
        ),
        fields: z.record(z.string(), z.unknown()).describe(
          'Column values keyed by internal name, e.g. {"Title":"New row","Status":"Open"}',
        ),
      }),
      execute: async (
        args: { list: string; fields: Record<string, unknown> },
        context: Context,
      ) => {
        const startMs = Date.now();

        if (Object.keys(args.fields).length === 0) {
          throw new Error(
            "fields is empty — a create with no column values would produce a " +
              "blank item. Pass at least one internal column name (get_list " +
              "lists them under writableColumns).",
          );
        }

        const accessToken = await getAccessToken(context.globalArgs);
        const site = await resolveSite(
          accessToken,
          context.globalArgs.siteHostPath,
        );
        const list = await resolveList(accessToken, site.id, args.list);

        const created = await withGraphContext(
          `Failed to create item in "${args.list}"`,
          () =>
            graphRequest<GraphListItem>(
              accessToken,
              "POST",
              `/sites/${site.id}/lists/${list.id}/items`,
              { fields: args.fields },
            ),
        );

        const handle = await context.writeResource(
          "itemWrite",
          slug(`create-${list.displayName ?? list.id}-${created.id}`, "create"),
          {
            siteId: site.id,
            siteHostPath: context.globalArgs.siteHostPath,
            operation: "create",
            listId: list.id,
            listName: list.displayName ?? list.name ?? list.id,
            itemId: created.id,
            fields: created.fields ?? args.fields,
            webUrl: created.webUrl ?? null,
            fetchedAt: new Date().toISOString(),
            durationMs: Date.now() - startMs,
            collectedBy: EXTENSION_NAME,
          },
        );

        context.logger.info("Created item {id} in {name}", {
          id: created.id,
          name: list.displayName ?? args.list,
        });
        return { dataHandles: [handle] };
      },
    },

    update_item: {
      description:
        "Update one list item's fields (PATCH — only the columns you pass are " +
        "changed). Fields are keyed by INTERNAL column name. Verify the item " +
        "ID with list_items before running this.",
      arguments: z.object({
        list: z.string().min(1, "list must not be empty").describe(
          "List GUID, URL name, or display name",
        ),
        itemId: z.string().min(1, "itemId must not be empty").describe(
          "List item ID (from list_items output)",
        ),
        fields: z.record(z.string(), z.unknown()).describe(
          "Column values to set, keyed by internal name. Unlisted columns are left alone.",
        ),
      }),
      execute: async (
        args: {
          list: string;
          itemId: string;
          fields: Record<string, unknown>;
        },
        context: Context,
      ) => {
        const startMs = Date.now();

        if (Object.keys(args.fields).length === 0) {
          throw new Error(
            "fields is empty — an update with no column values is a no-op. " +
              "Pass at least one internal column name to change.",
          );
        }

        const accessToken = await getAccessToken(context.globalArgs);
        const site = await resolveSite(
          accessToken,
          context.globalArgs.siteHostPath,
        );
        const list = await resolveList(accessToken, site.id, args.list);

        const updated = await withGraphContext(
          `Failed to update item "${args.itemId}" in "${args.list}"`,
          () =>
            graphRequest<Record<string, unknown>>(
              accessToken,
              "PATCH",
              `/sites/${site.id}/lists/${list.id}/items/${
                encodeURIComponent(args.itemId)
              }/fields`,
              args.fields,
            ),
        );

        const handle = await context.writeResource(
          "itemWrite",
          slug(
            `update-${list.displayName ?? list.id}-${args.itemId}`,
            "update",
          ),
          {
            siteId: site.id,
            siteHostPath: context.globalArgs.siteHostPath,
            operation: "update",
            listId: list.id,
            listName: list.displayName ?? list.name ?? list.id,
            itemId: args.itemId,
            fields: updated ?? args.fields,
            webUrl: null,
            fetchedAt: new Date().toISOString(),
            durationMs: Date.now() - startMs,
            collectedBy: EXTENSION_NAME,
          },
        );

        context.logger.info(
          "Updated item {id} in {name} ({count} column(s))",
          {
            id: args.itemId,
            name: list.displayName ?? args.list,
            count: Object.keys(args.fields).length,
          },
        );
        return { dataHandles: [handle] };
      },
    },

    delete_item: {
      description:
        "Delete one list item. Requires confirm=true. The item's field values " +
        "are read and persisted before the delete, so the itemWrite resource " +
        "is a record of what was removed.",
      arguments: z.object({
        list: z.string().min(1, "list must not be empty").describe(
          "List GUID, URL name, or display name",
        ),
        itemId: z.string().min(1, "itemId must not be empty").describe(
          "List item ID (from list_items output)",
        ),
        confirm: z.boolean().default(false).describe(
          "Must be true. Guards against a delete run with the wrong itemId.",
        ),
      }),
      execute: async (
        args: { list: string; itemId: string; confirm: boolean },
        context: Context,
      ) => {
        const startMs = Date.now();

        if (!args.confirm) {
          throw new Error(
            `Refusing to delete item "${args.itemId}" without confirm=true. ` +
              "Verify the item with list_items first, then re-run with " +
              "--input confirm=true.",
          );
        }

        const accessToken = await getAccessToken(context.globalArgs);
        const site = await resolveSite(
          accessToken,
          context.globalArgs.siteHostPath,
        );
        const list = await resolveList(accessToken, site.id, args.list);

        // Read before deleting. A delete that returns 204 tells you nothing
        // about what was in the row, and SharePoint's recycle bin is not
        // reachable from Graph — so the pre-image is the only record.
        const existing = await withGraphContext(
          `Failed to read item "${args.itemId}" in "${args.list}" before delete`,
          () =>
            graphRequest<GraphListItem>(
              accessToken,
              "GET",
              `/sites/${site.id}/lists/${list.id}/items/${
                encodeURIComponent(args.itemId)
              }?$expand=fields`,
            ),
        );

        await withGraphContext(
          `Failed to delete item "${args.itemId}" in "${args.list}"`,
          () =>
            graphRequest<Record<string, never>>(
              accessToken,
              "DELETE",
              `/sites/${site.id}/lists/${list.id}/items/${
                encodeURIComponent(args.itemId)
              }`,
            ),
        );

        const handle = await context.writeResource(
          "itemWrite",
          slug(
            `delete-${list.displayName ?? list.id}-${args.itemId}`,
            "delete",
          ),
          {
            siteId: site.id,
            siteHostPath: context.globalArgs.siteHostPath,
            operation: "delete",
            listId: list.id,
            listName: list.displayName ?? list.name ?? list.id,
            itemId: args.itemId,
            fields: existing.fields ?? {},
            webUrl: existing.webUrl ?? null,
            fetchedAt: new Date().toISOString(),
            durationMs: Date.now() - startMs,
            collectedBy: EXTENSION_NAME,
          },
        );

        context.logger.info("Deleted item {id} from {name}", {
          id: args.itemId,
          name: list.displayName ?? args.list,
        });
        return { dataHandles: [handle] };
      },
    },
  },
};
