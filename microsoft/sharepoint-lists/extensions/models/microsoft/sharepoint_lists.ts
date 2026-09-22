/**
 * SharePoint Lists Model — read and write SharePoint list items via Graph API.
 *
 * Authenticates as a public client through the device code flow, with the
 * OAuth2 and Graph plumbing in `_lib` (adapted from @webframp/microsoft/teams;
 * attribution in those file headers). Where the existing SharePoint extensions
 * in the registry cover document libraries (`/drive/...`) and are read-only,
 * this covers the list surface (`/sites/{site}/lists/...`) and can write.
 *
 * @module
 */
// deno-lint-ignore-file no-import-prefix
import { z } from "npm:zod@4";
import {
  DEFAULT_SHAREPOINT_SCOPES,
  initiateDeviceCode,
  pollDeviceCode,
  refreshAccessToken,
} from "./_lib/auth.ts";
import {
  GraphApiError,
  graphRequest,
  graphRequestPaginated,
} from "./_lib/graph.ts";

const EXTENSION_NAME = "@aspec451/microsoft/sharepoint-lists";

/**
 * Graph rejects `$filter` and `$orderby` on list-item columns that SharePoint
 * has not indexed unless the caller opts in to the failure risk with this
 * header. Without it a perfectly valid filter on an unindexed column comes
 * back as a flat 400, which reads like a syntax error rather than a missing
 * index — so it is sent whenever a filter or sort is supplied.
 */
const HONOR_NON_INDEXED = "HonorNonIndexedQueriesWarningMayFailRandomly";

/**
 * Columns SharePoint moves on its own, excluded from the `list_changes` diff.
 *
 * Diffing these reports an edit on every row that was saved for any reason and
 * buries the real change. `Modified`, `_UIVersionString` and `@odata.etag` bump
 * on every save — they are carried as corroboration of a bump, never as a field
 * change. `EditorLookupId` follows whoever saved, which `lastModifiedBy`
 * already says by name. The rest is list bookkeeping: `LinkTitle*` mirrors
 * Title, and the `_Compliance*` set is retention plumbing.
 *
 * Everything else is diffed, including columns added to the list after this
 * ran last — the field set comes from the rows themselves, so a new column is
 * picked up without a code change. Add to this per call with `ignoreFields`.
 */
const BOOKKEEPING_FIELDS = new Set([
  "@odata.etag",
  "id",
  "ContentType",
  "Created",
  "Modified",
  "AuthorLookupId",
  "EditorLookupId",
  "_UIVersionString",
  "Edit",
  "LinkTitle",
  "LinkTitleNoMenu",
  "ItemChildCount",
  "FolderChildCount",
  "Attachments",
  "_ComplianceFlags",
  "_ComplianceTag",
  "_ComplianceTagWrittenTime",
  "_ComplianceTagUserId",
]);

/**
 * What swamp leaves behind where it scrubbed a vault secret out of resource
 * content on its way to disk.
 *
 * It applies to the whole value, so a secret that also appears inside list
 * content — a tenant GUID pasted into a rich-text column as part of a
 * SharePoint or Teams deep link is the common case — comes back from the
 * baseline with `***` where a fresh fetch has the real characters. Comparing
 * those two naively reports that column as changed on every run forever, so a
 * redacted span is treated as unknowable rather than different, and reported
 * in `redactionBlindFields` so the blind spot is visible.
 */
const REDACTED = "***";

// A run that lost its filter, or hit a permissions change mid-flight, returns
// far fewer rows than the baseline holds — and every missing row looks like a
// deletion. Past this share of the baseline the absences are treated as a bad
// fetch rather than as news: they are still reported, but the baseline is NOT
// replaced and `suspectRemoval` is set, so a caller can fail the run instead of
// announcing that the department was disbanded. The floor keeps ordinary churn
// on a small list from tripping it — three rows deleted from a ten-row list is
// 30% and entirely plausible.
const REMOVAL_ALARM_SHARE = 0.25;
const REMOVAL_ALARM_FLOOR = 8;

// Each absent row costs one Graph GET to tell a deletion from a row that was
// edited out of scope. Removals are normally 0, and this cap only binds on a
// run already flagged suspect, where the classification does not matter because
// the baseline is held back regardless.
const MAX_REMOVAL_PROBES = 30;

/**
 * An internal column name, as Graph reports it and as `$filter` accepts it.
 *
 * `scopeColumn` is interpolated into an OData filter, where a quote or a space
 * is either a syntax error or an injection; internal names never contain either
 * (a space is escaped as `_x0020_`), so rejecting anything else costs nothing.
 */
const INTERNAL_COLUMN_NAME = /^[A-Za-z_][A-Za-z0-9_]*$/;

// =============================================================================
// Schemas
// =============================================================================

const GlobalArgsSchema = z.object({
  tenantId: z.string().min(1).describe(
    "Entra tenant GUID (e.g. aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee)",
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
  truncated: z.boolean().describe(
    "Whether the fetch hit maxPages before exhausting the column list",
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
  alreadyAbsent: z.boolean().optional().describe(
    "For delete: the item was already gone, so nothing was removed",
  ),
  deleteConfirmed: z.boolean().optional().describe(
    "For delete: true once Graph confirmed removal. A record with false is a " +
      "pre-image written before the DELETE was issued — the item may still exist",
  ),
  ...MetaFields,
});

const FieldChangeSchema = z.object({
  field: z.string().describe(
    "Internal column name, e.g. Expected_x0020_Reply_x0020_By",
  ),
  label: z.string().describe(
    "The column's display name, read from the list's own column definitions " +
      "per run so a rename follows. Internal names are _x0020_-escaped and " +
      "several read nothing like the form does, so a report built on them " +
      "makes the reader translate. Falls back to the internal name for a " +
      "column that has since been removed from the list.",
  ),
  from: z.string().nullable().describe(
    "Previous value; null when the column was empty or absent",
  ),
  to: z.string().nullable().describe(
    "Current value; null when the column was cleared",
  ),
  kind: z.enum(["set", "cleared", "changed"]).describe(
    "set: was empty, now has a value. cleared: had a value, now empty. " +
      "changed: both sides non-empty.",
  ),
});

const ChangedRowSchema = z.object({
  itemId: z.string().describe("List item ID"),
  title: z.string().nullable().describe(
    "The row's Title column — the one column every SharePoint list has",
  ),
  scopeValue: z.string().nullable().describe(
    "The row's current value in scopeColumn, or null when the run is unscoped",
  ),
  lastModified: z.string().describe("lastModifiedDateTime, ISO 8601"),
  lastModifiedBy: z.string().nullable().describe(
    "Display name of whoever saved the row last",
  ),
  version: z.string().nullable().describe(
    "_UIVersionString — cheap corroboration that SharePoint bumped the row",
  ),
  itemUrl: z.string().nullable().describe(
    "DispForm link. Built from the list's webUrl because Graph's own webUrl " +
      "for a list item is not a link a browser can open.",
  ),
  context: z.record(z.string(), z.string()).describe(
    "Current values of the columns named in `contextFields`, carried on every " +
      "reported row. A modified row only lists the columns that moved, so " +
      "without this a reader cannot see what state the row is in now without " +
      "going back to SharePoint. Empty when contextFields was not passed.",
  ),
  changes: z.array(FieldChangeSchema).describe(
    "Field-level differences against the baseline, in display-name order. On " +
      "an added row every non-empty tracked column is reported as `set`, " +
      "since all of it is new. Empty on a touched row — that is the point of " +
      "that category.",
  ),
});

const RemovedRowSchema = z.object({
  itemId: z.string().describe("List item ID as the baseline held it"),
  title: z.string().nullable().describe("Title as the baseline last saw it"),
  lastSeenAt: z.string().describe(
    "When the baseline that still held this row was captured",
  ),
  lastModified: z.string().describe(
    "The row's lastModifiedDateTime as the baseline saw it",
  ),
  itemUrl: z.string().nullable().describe("DispForm link"),
  fate: z.enum([
    "deleted",
    "rescoped",
    "still-present",
    "unknown",
    "unprobed",
  ]).describe(
    "deleted: the item ID no longer resolves in the list. rescoped: the row " +
      "is still there but its scopeColumn now says something else, so it left " +
      "this report's scope rather than the list — `nowInScope` says where it " +
      "went. still-present: the row is still there and still in scope, so " +
      "nothing happened to it and the fetch that missed it came back short — " +
      "not a departure, and it holds the baseline back. unknown: the probe " +
      "ran and failed, so whether this row is gone is genuinely not known — " +
      "it holds the baseline back too, since advancing would drop a row that " +
      "may never have left. unprobed: the probe was deliberately skipped, " +
      "either because the run was already flagged suspect or because the " +
      "removals ran past the probe budget.",
  ),
  nowInScope: z.string().nullable().describe(
    "The row's current scopeColumn value, for a rescoped row",
  ),
  lastKnownFields: z.record(z.string(), z.string()).describe(
    "The row's tracked column values as the baseline held them — the only " +
      "record of a deleted row that is left",
  ),
});

const ListChangesSchema = z.object({
  ...SiteFields,
  listId: z.string().describe("Resolved list GUID"),
  listName: z.string().describe("Display name of the list"),
  scopeColumn: z.string().nullable().describe(
    "Internal column name the run was scoped to, or null for the whole list",
  ),
  scopeValue: z.string().nullable().describe(
    "Value of scopeColumn the run was scoped to",
  ),
  filter: z.string().nullable().describe(
    "The OData $filter actually sent, scope and caller filter combined",
  ),
  baselineInstance: z.string().describe(
    "Data instance the baseline is kept under. Derived from the list and the " +
      "scope, so each scope diffs against its own population.",
  ),
  firstRun: z.boolean().describe(
    "True when there was no usable baseline to compare against — none " +
      "stored, or one covering a different list, a different scope, or a " +
      "different set of tracked columns. Nothing is reported as changed, " +
      "since a first run would otherwise announce every row as new, and the " +
      "baseline is established for next time.",
  ),
  since: z.string().nullable().describe(
    "When the baseline being compared against was captured. This is the true " +
      "left edge of the report: the last SUCCESSFUL run, not the last " +
      "scheduled one, so after a failed run the window is wider and nothing " +
      "falls through it. Null on a first run.",
  ),
  sinceGapHours: z.number().nullable().describe(
    "Hours between the baseline capture and this fetch. Well over the " +
      "expected interval means runs have been missing — worth knowing before " +
      "reading a quiet report as a quiet day.",
  ),
  added: z.array(ChangedRowSchema).describe(
    "Rows in scope now that the baseline did not hold",
  ),
  modified: z.array(ChangedRowSchema).describe(
    "Rows whose tracked columns differ from the baseline, newest edit first",
  ),
  touched: z.array(ChangedRowSchema).describe(
    "Rows SharePoint bumped — a newer lastModifiedDateTime or version — " +
      "where no tracked column actually differs. Reported rather than " +
      "dropped, because the honest answer is 'somebody saved this row and " +
      "nothing visible moved' (an attachment, a permission, a re-save), not " +
      "'nothing happened'.",
  ),
  removed: z.array(RemovedRowSchema).describe(
    "Rows the baseline held that this fetch did not return, each classified " +
      "as deleted, rescoped, or still present",
  ),
  redactionBlindFields: z.array(z.string()).describe(
    "Columns whose baseline value carries a `***` where swamp scrubbed a " +
      "vault secret out of it, so the redacted span could not be compared. " +
      "The literal text around it still matched, and an edit outside the " +
      "redacted span would still have been caught — this is the list of " +
      "places where a change could hide.",
  ),
  suspectRemoval: z.boolean().describe(
    "True when so much of the baseline went missing at once that a bad fetch " +
      "is likelier than real deletions. The removals are still listed, but " +
      "the baseline is NOT replaced, so the next run compares against the " +
      "same known-good population instead of inheriting the damage.",
  ),
  baselineUpdated: z.boolean().describe(
    "Whether this run replaced the baseline. False on a truncated fetch, a " +
      "suspect removal, or updateBaseline=false — in each case the next run " +
      "re-reports this window rather than starting from a population that is " +
      "missing rows.",
  ),
  baselineHeldReason: z.string().nullable().describe(
    "Why the baseline was held back, if it was",
  ),
  totals: z.object({
    scanned: z.number().describe(
      "Rows Graph returned, before the client-side scope re-check",
    ),
    inScope: z.number().describe(
      "Rows in scope now — the population being watched",
    ),
    baselineRows: z.number().describe("Rows the baseline held"),
    added: z.number(),
    modified: z.number(),
    touched: z.number(),
    removed: z.number(),
    fieldChanges: z.number().describe(
      "Individual column changes across the modified rows",
    ),
    stillPresent: z.number().describe(
      "Baseline rows this fetch did not return that are still in scope — " +
        "evidence of a short fetch, not of anything leaving",
    ),
    probeFailed: z.number().describe(
      "Absent rows whose classification probe failed. Counted in `removed` " +
        "and in `changed`, because a row missing from the fetch is worth " +
        "reading either way, but they hold the baseline back so the next run " +
        "says so definitively.",
    ),
    changed: z.number().describe(
      "added + modified + genuinely-removed — the count that answers 'is " +
        "there anything to read'. Two things are deliberately excluded: " +
        "`touched`, because a bumped row with no visible field change is not " +
        "an update, and still-present removals, because those are a short " +
        "fetch rather than anything leaving.",
    ),
  }),
  summarizeColumn: z.string().nullable().describe(
    "Internal column name the population was counted by, or null",
  ),
  byValue: z.array(z.object({
    value: z.string().describe(
      "The column value, or '(empty)' where it had none",
    ),
    count: z.number(),
  })).describe(
    "How the population divides by summarizeColumn right now, busiest first " +
      "— context for the changes rather than part of them. Empty when " +
      "summarizeColumn was not passed.",
  ),
  filterHonored: z.boolean().describe(
    "False when Graph returned rows outside the requested scope, meaning it " +
      "ignored the filter on an unindexed column and the scope was applied " +
      "client-side instead. Counts are still correct; the fetch was larger.",
  ),
  truncated: z.boolean().describe(
    "True when the page cap cut the walk short — the population is a floor, " +
      "so absences are not conclusions and the baseline is held back",
  ),
  ...MetaFields,
});

/**
 * The population and column values as of the last clean run — the only memory
 * `list_changes` has, and not something to read as a report.
 */
const ChangeBaselineSchema = z.object({
  listId: z.string().describe(
    "The list this baseline covers. Checked before diffing: the instance name " +
      "is derived from the list's display name, which a rename can move onto " +
      "another list's population.",
  ),
  listName: z.string().describe("Display name at capture time"),
  scopeColumn: z.string().nullable().describe(
    "The scope this baseline covers. Also checked before diffing — comparing " +
      "one scope's population against another's would report every row as " +
      "both added and removed.",
  ),
  scopeValue: z.string().nullable(),
  fieldPolicy: z.string().describe(
    "Canonical form of the trackFields and ignoreFields the capture was taken " +
      "under. Both decide which columns a stored row even carries, so a " +
      "baseline captured under a different policy would report every column " +
      "the new policy drops as cleared. A mismatch is treated as no baseline.",
  ),
  capturedAt: z.string().describe("ISO 8601 timestamp of the capture"),
  rowCount: z.number().describe("Rows the capture held"),
  rows: z.array(z.object({
    itemId: z.string(),
    title: z.string().nullable(),
    lastModified: z.string(),
    lastModifiedBy: z.string().nullable(),
    version: z.string().nullable(),
    itemUrl: z.string().nullable(),
    fields: z.record(z.string(), z.string()).describe(
      "Tracked columns as strings, bookkeeping columns dropped. Values are " +
        "kept in full rather than hashed: a hash can prove a change happened " +
        "but cannot say what it was, and 'Status: In Progress → Confirmed' is " +
        "the whole product here.",
    ),
  })).describe("The rows as they stood, in no particular order"),
  ...MetaFields,
});

const BootstrapResultSchema = z.object({
  status: z.string().describe("Authentication result status"),
  message: z.string().describe("Human-readable guidance for the caller"),
  scopes: z.string().optional().describe(
    "Space-separated delegated scopes the sign-in actually granted",
  ),
  refreshToken: z.string().meta({ sensitive: true }).describe(
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

/**
 * Short stable digest of a value, base36. FNV-1a — not a security hash, just
 * enough to keep two names that agree on their leading characters apart.
 */
function shortHash(value: string): string {
  let h = 2166136261;
  for (let i = 0; i < value.length; i++) {
    h ^= value.charCodeAt(i);
    h = Math.imul(h, 16777619);
  }
  return (h >>> 0).toString(36);
}

/**
 * Turn an identifier into a short, file-name-safe data instance name.
 *
 * Truncation is where instance names turn dangerous: two values agreeing on
 * their first `max` characters would slug identically, and writes that should
 * be separate records would land on one instance and overwrite each other. So
 * when truncation actually bites, a digest of the full value replaces the tail.
 * Values that fit are returned untouched and stay readable.
 */
function slug(value: string, fallback = "item", max = 48): string {
  const cleaned = value
    .replace(/[^a-zA-Z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .toLowerCase();

  if (!cleaned) return fallback;
  if (cleaned.length <= max) return cleaned;

  const digest = shortHash(value);
  const head = cleaned.slice(0, Math.max(1, max - digest.length - 1));
  return `${head}-${digest}`;
}

/** Name the per-list instance the read methods write to. */
function listInstance(list: GraphList): string {
  return slug(list.displayName ?? list.name ?? list.id, "list");
}

/**
 * Name the itemWrite instance for one operation against one item.
 *
 * The item ID is the part that has to survive, and slug's budget applies to
 * whatever it is handed — so interpolating a variable-width list name ahead of
 * the ID can push the ID off the end, and then every item in a long-named list
 * writes to one instance and clobbers the last one's record. Slug the parts
 * separately, under their own budgets, so the ID is always present.
 */
function itemWriteInstance(
  operation: "create" | "update" | "delete",
  list: GraphList,
  itemId: string,
): string {
  const listPart = slug(list.displayName ?? list.name ?? list.id, "list", 28);
  return `${operation}-${listPart}-${slug(itemId, "item", 12)}`;
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
  } catch (e) {
    // Only a miss earns the fallback. Swallowing everything would let a 401,
    // 403, or 429 trigger a site-wide enumeration and then report a list that
    // plainly exists as missing, misattributing an auth or throttling failure.
    // Graph answers an unknown list segment with 404, and a segment it will
    // not even accept as an identifier with 400 — both mean "not this list".
    const missed = e instanceof GraphApiError &&
      (e.statusCode === 404 || e.statusCode === 400);
    if (!missed) {
      throw new Error(
        `Failed to fetch list "${list}": ${
          e instanceof Error ? e.message : String(e)
        }`,
        { cause: e },
      );
    }
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

// ---------------------------------------------------------------------------
// Change detection helpers
// ---------------------------------------------------------------------------

type FieldChange = z.infer<typeof FieldChangeSchema>;
type ChangedRow = z.infer<typeof ChangedRowSchema>;
type RemovedRow = z.infer<typeof RemovedRowSchema>;
type Baseline = z.infer<typeof ChangeBaselineSchema>;
type BaselineRow = Baseline["rows"][number];

/**
 * A row as this run fetched it. The baseline keeps a subset: `scopeValue` and
 * `context` are read fresh every run, so storing them would only let them go
 * stale.
 */
type FetchedRow = BaselineRow & {
  scopeValue: string | null;
  context: Record<string, string>;
};

/** A trimmed string, or null for anything empty or not a string. */
function nonEmpty(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  return trimmed === "" ? null : trimmed;
}

/**
 * Render one column value as the string the diff compares.
 *
 * Everything becomes a string because the baseline round-trips through JSON: a
 * number or boolean survives that unchanged, but a Date does not, and a diff
 * that compares a Date against its own serialization reports every row as
 * modified forever. Lookup and person columns arrive as objects whose label is
 * the value a reader means, so that is what is kept.
 */
function fieldToString(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "string") return value.trim();
  if (typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  if (Array.isArray(value)) {
    return value.map(fieldToString).filter((v) => v !== "").join("; ");
  }
  if (typeof value === "object") {
    const o = value as Record<string, unknown>;
    const label = o["LookupValue"] ?? o["displayName"] ?? o["Label"] ??
      o["Email"];
    if (label !== undefined) return fieldToString(label);
    return JSON.stringify(o);
  }
  return String(value);
}

/**
 * Reduce a row's expanded fields to the columns worth diffing, as strings.
 *
 * An absent column and an empty one are the same thing here — SharePoint omits
 * a field it has never been given a value for, so keeping empties would make
 * "column filled in for the first time" indistinguishable from a row that
 * simply grew a column — so empties are dropped on both sides and a key
 * present on one side only is a set or a clear.
 */
function trackedFields(
  fields: Record<string, unknown>,
  ignore: Set<string>,
  track: Set<string> | null,
): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [key, value] of Object.entries(fields)) {
    if (track ? !track.has(key) : ignore.has(key)) continue;
    const rendered = fieldToString(value);
    if (rendered !== "") out[key] = rendered;
  }
  return out;
}

/**
 * Flatten one Graph list item into the row shape the diff and the baseline
 * share. `lastModifiedBy` arrives nested, the same way `mapItem` unwraps it.
 */
function changeRow(
  raw: GraphListItem,
  listWebUrl: string | null,
  scopeColumn: string | null,
  ignore: Set<string>,
  track: Set<string> | null,
  contextFields: string[],
): FetchedRow {
  const fields = raw.fields ?? {};
  const itemId = String(raw.id ?? fields["id"] ?? "");
  const context: Record<string, string> = {};
  for (const name of contextFields) {
    const rendered = fieldToString(fields[name]);
    if (rendered !== "") context[name] = rendered;
  }
  return {
    context,
    itemId,
    title: nonEmpty(fields["Title"]),
    scopeValue: scopeColumn ? nonEmpty(fields[scopeColumn]) : null,
    lastModified: String(
      raw.lastModifiedDateTime ?? fields["Modified"] ?? "",
    ),
    lastModifiedBy: raw.lastModifiedBy?.user?.displayName ?? null,
    version: nonEmpty(fields["_UIVersionString"]),
    itemUrl: listWebUrl && itemId
      ? `${listWebUrl}/DispForm.aspx?ID=${itemId}`
      : null,
    fields: trackedFields(fields, ignore, track),
  };
}

/** Drop the carrier fields a baseline row does not keep. */
function baselineRow(row: FetchedRow): BaselineRow {
  return {
    itemId: row.itemId,
    title: row.title,
    lastModified: row.lastModified,
    lastModifiedBy: row.lastModifiedBy,
    version: row.version,
    itemUrl: row.itemUrl,
    fields: row.fields,
  };
}

/** Present one fetched row as a report row carrying the given changes. */
function reportRow(
  row: FetchedRow,
  changes: FieldChange[],
): ChangedRow {
  return {
    itemId: row.itemId,
    title: row.title,
    scopeValue: row.scopeValue,
    lastModified: row.lastModified,
    lastModifiedBy: row.lastModifiedBy,
    version: row.version,
    itemUrl: row.itemUrl,
    context: row.context,
    changes,
  };
}

/**
 * True when `stored` is `fresh` with one or more spans replaced by `***`.
 *
 * The literals around each redaction are anchored — a stored "a***b" requires
 * `fresh` to start with "a" and end with "b" — so this only ever forgives the
 * redacted span itself. Everything outside it still has to match exactly, so a
 * real edit elsewhere in the same column is still caught; the span itself is a
 * blind spot and is reported as one rather than passed off as "no change".
 */
function redactionBlind(stored: string, fresh: string): boolean {
  if (!stored.includes(REDACTED)) return false;
  const parts = stored.split(REDACTED);
  let pos = 0;
  for (let i = 0; i < parts.length; i++) {
    const part = parts[i];
    if (part === "") continue;
    const at = fresh.indexOf(part, pos);
    if (at === -1) return false;
    // Text before the first redaction has to sit at the very start, or the two
    // values differ in a part that was never redacted.
    if (i === 0 && at !== 0) return false;
    pos = at + part.length;
  }
  // Likewise, text after the last redaction has to run to the end.
  return parts[parts.length - 1] === "" || pos === fresh.length;
}

/** Order a row's changes the way the form reads: by display name. */
function byLabel(a: FieldChange, b: FieldChange): number {
  return a.label.localeCompare(b.label);
}

/**
 * Compare two field maps. Both sides have had empties dropped, so a key on one
 * side only is a column being filled in or cleared, not a schema difference.
 */
function diffFields(
  before: Record<string, string>,
  after: Record<string, string>,
  labels: Map<string, string>,
  blind: string[],
): FieldChange[] {
  const out: FieldChange[] = [];
  for (const key of new Set([...Object.keys(before), ...Object.keys(after)])) {
    const from = before[key] ?? "";
    const to = after[key] ?? "";
    if (from === to) continue;
    if (redactionBlind(from, to)) {
      blind.push(key);
      continue;
    }
    out.push({
      field: key,
      label: labels.get(key) ?? key,
      from: from === "" ? null : from,
      to: to === "" ? null : to,
      kind: from === "" ? "set" : to === "" ? "cleared" : "changed",
    });
  }
  return out.sort(byLabel);
}

/**
 * Every tracked column of an added row, as a `set`. An added row has no
 * before-state to diff against, and reporting it as a bare ID and title would
 * make the reader fetch the row to learn anything about it.
 */
function setChanges(
  fields: Record<string, string>,
  labels: Map<string, string>,
): FieldChange[] {
  return Object.entries(fields)
    .map(([field, to]) => ({
      field,
      label: labels.get(field) ?? field,
      from: null,
      to,
      kind: "set" as const,
    }))
    .sort(byLabel);
}

/**
 * Internal → display names for a list's columns.
 *
 * Read per run rather than kept in the baseline, so a renamed column follows
 * immediately. Hidden and read-only columns are included: a calculated column
 * that someone starts using should still be labelled.
 */
async function columnLabels(
  accessToken: string,
  siteId: string,
  listId: string,
  maxPages: number,
): Promise<Map<string, string>> {
  const labels = new Map<string, string>();
  try {
    const result = await graphRequestPaginated<GraphColumn>(
      accessToken,
      `/sites/${siteId}/lists/${listId}/columns`,
      { "$select": "name,displayName" },
      undefined,
      fetch,
      maxPages,
    );
    for (const column of result.items) {
      const name = nonEmpty(column.name);
      const displayName = nonEmpty(column.displayName);
      if (name && displayName) labels.set(name, displayName);
    }
  } catch {
    // A label is a nicety and the internal name is a usable fallback. Losing
    // the map must not cost the whole diff.
  }
  return labels;
}

/**
 * Tell a deleted row from one that was edited out of scope, and both from one
 * the fetch simply failed to return.
 *
 * All three vanish from a scoped fetch and none of them is the same news: a
 * deletion is a row withdrawn, a rescope is a row that became someone else's,
 * and a row still sitting in scope is a short page walk — which is news about
 * the fetch, not about the row, and is the one fate that holds the baseline
 * back.
 */
async function probeRemovedItem(
  accessToken: string,
  siteId: string,
  listId: string,
  itemId: string,
  scopeColumn: string | null,
  scopeValue: string | null,
): Promise<{ fate: RemovedRow["fate"]; nowInScope: string | null }> {
  try {
    const item = await graphRequest<GraphListItem>(
      accessToken,
      "GET",
      `/sites/${siteId}/lists/${listId}/items/${itemId}?$expand=fields`,
    );
    if (!scopeColumn) {
      // Unscoped: the row is still in the list, so it did not go anywhere.
      return { fate: "still-present", nowInScope: null };
    }
    const current = nonEmpty((item.fields ?? {})[scopeColumn]);
    if (current === scopeValue) {
      return { fate: "still-present", nowInScope: current };
    }
    return { fate: "rescoped", nowInScope: current };
  } catch (e) {
    // A 404 is the answer, not a failure. Anything else — a timeout, a
    // throttle that outlasted its retries — leaves the fate unknown rather
    // than asserting a deletion on the strength of it.
    const missing = e instanceof GraphApiError &&
      (e.statusCode === 404 || /itemNotFound/i.test(e.graphCode));
    return missing
      ? { fate: "deleted", nowInScope: null }
      : { fate: "unknown", nowInScope: null };
  }
}

/**
 * Canonical key for the column policy a run diffs under.
 *
 * Only the caller's own arguments go in. The built-in bookkeeping set is left
 * out deliberately: folding it in would invalidate every stored baseline the
 * day a release adds a column to it, which costs a re-baseline everywhere to
 * save one spurious `cleared` on one column once.
 */
function fieldPolicyKey(
  trackFields: string[] | undefined,
  ignoreFields: string[],
): string {
  const track = trackFields && trackFields.length > 0
    ? [...trackFields].sort().join(",")
    : "*";
  const ignore = [...ignoreFields].sort().join(",");
  return `track=${track};ignore=${ignore}`;
}

/**
 * Name the changes and baseline instances for one list and scope.
 *
 * Both parts get their own slug budget, for the reason `itemWriteInstance`
 * documents: slugging the concatenation lets a long list name push the scope
 * off the end, and then every scope on that list would diff against one shared
 * baseline and report the whole population as added and removed by turns.
 */
function changeInstances(
  list: GraphList,
  scopeColumn: string | null,
  scopeValue: string | null,
): { changes: string; baseline: string } {
  const listPart = slug(list.displayName ?? list.name ?? list.id, "list", 28);
  const scopePart = scopeColumn
    ? slug(`${scopeColumn}-${scopeValue ?? ""}`, "scoped", 24)
    : "all-rows";
  const changes = `${listPart}-${scopePart}`;
  return { changes, baseline: `baseline-${changes}` };
}

/**
 * The subset of the check context these checks use.
 *
 * Separate from `Context` because checks deliberately cannot write resources —
 * they inspect and return a verdict, so `writeResource` is absent there.
 */
interface CheckContext {
  globalArgs: GlobalArgs;
  logger: {
    info: (msg: string, props?: Record<string, unknown>) => void;
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
  /**
   * The newest version of one data instance, or null when nothing has been
   * written under that name. `list_changes` is the only method that reads:
   * everything else here answers from Graph, but a diff needs a previous
   * state, and the previous run's own output is the only thing that has one.
   */
  readResource: (
    name: string,
    version?: number,
  ) => Promise<Record<string, unknown> | null>;
  logger: {
    info: (msg: string, props?: Record<string, unknown>) => void;
    warn: (msg: string, props?: Record<string, unknown>) => void;
    error: (msg: string, props?: Record<string, unknown>) => void;
  };
}

// =============================================================================
// Model Definition
// =============================================================================

/** SharePoint list read/write model via Graph API. */
export const model = {
  type: "@aspec451/microsoft/sharepoint-lists",
  version: "2026.09.22.1",
  globalArguments: GlobalArgsSchema,

  upgrades: [
    {
      toVersion: "2026.09.22.1",
      description:
        "Add the list_changes method and its listChanges/listBaseline " +
        "resources. No globalArguments change — an existing instance needs " +
        "nothing but the version bump, and the first list_changes run " +
        "establishes its own baseline.",
      upgradeAttributes: (old: Record<string, unknown>) => old,
    },
  ],

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
    listChanges: {
      description:
        "What changed on one list since the previous list_changes run — rows " +
        "added, rows whose columns moved, rows bumped with nothing visible, " +
        "and rows that are gone",
      schema: ListChangesSchema,
      lifetime: "infinite" as const,
      garbageCollection: 30,
    },
    listBaseline: {
      description:
        "The population and column values as of the last clean list_changes " +
        "run, kept so the next run has something to diff against. Machine " +
        "state, not a report — read listChanges instead.",
      schema: ChangeBaselineSchema,
      lifetime: "infinite" as const,
      // Only the newest version is ever read, but a few are kept so a bad
      // baseline can be rolled back to by hand
      // (`swamp data get <model> <instance> --version N`).
      garbageCollection: 10,
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

        context.logger.info(
          "Starting device code flow for tenant {tenant} with scopes: {scopes}",
          { tenant: tenantId, scopes },
        );

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

        // `offline_access` is what earns a refresh token. Without it Entra
        // answers with an access token alone, and the whole point of bootstrap
        // is the durable half — so fail here rather than write a record that
        // says "authenticated" and tells the user to vault a token that was
        // never issued, leaving the miss to surface at the first real call.
        if (!tokens.refresh_token) {
          throw new Error(
            "Authentication succeeded but Entra issued no refresh token, so " +
              "there is nothing to store and every later method would fail at " +
              "token refresh. The `offline_access` scope is what grants one — " +
              `add it to \`scopes\` (currently "${scopes}") and re-run ` +
              "bootstrap.",
          );
        }

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

        context.logger.info("Enumerating lists on site {site}", {
          site: context.globalArgs.siteHostPath,
        });

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
        maxPages: z.number().int().min(1).max(50).default(20).describe(
          "Cap on @odata.nextLink pages followed",
        ),
      }),
      execute: async (
        args: { list: string; includeHidden: boolean; maxPages: number },
        context: Context,
      ) => {
        const startMs = Date.now();

        context.logger.info(
          "Fetching column definitions for list {list} on site {site}",
          { list: args.list, site: context.globalArgs.siteHostPath },
        );

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
              undefined,
              undefined,
              fetch,
              args.maxPages,
            ),
        );

        if (result.truncated) {
          context.logger.warn(
            "Hit the maxPages cap ({cap}) before exhausting the column list — " +
              "{count} column(s) is a floor, and writableColumns may be " +
              "incomplete. Raise maxPages.",
            { cap: args.maxPages, count: result.items.length },
          );
        }

        const columns = result.items
          .map(mapColumn)
          .filter((c) => args.includeHidden || !c.hidden);

        const handle = await context.writeResource(
          "listSchema",
          listInstance(list),
          {
            siteId: site.id,
            siteHostPath: context.globalArgs.siteHostPath,
            list: mapList(list),
            columns,
            writableColumns: columns
              .filter((c) => !c.readOnly)
              .map((c) => c.name),
            truncated: result.truncated,
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

        context.logger.info(
          "Fetching items from list {list} on site {site} " +
            "(filter={filter}, orderBy={orderBy}, top={top}, maxPages={maxPages})",
          {
            list: args.list,
            site: context.globalArgs.siteHostPath,
            filter: args.filter ?? "none",
            orderBy: args.orderBy ?? "none",
            top: args.top,
            maxPages: args.maxPages,
          },
        );

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
          listInstance(list),
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

    list_changes: {
      description:
        "What changed on a list since this method last ran. Diffs the list " +
        "against a baseline the method keeps itself, reporting added rows, " +
        "field-level column changes with their before and after values, rows " +
        "SharePoint bumped where nothing visible moved, and rows that are " +
        "gone — each classified as deleted, edited out of scope, or missing " +
        "from a short fetch. Answers what `list_items` cannot: Graph returns " +
        "each row's lastModifiedDateTime, but a bump is not a diff and a " +
        "deleted row has no timestamp left to filter on. Holds the baseline " +
        "back on a truncated or implausible fetch, so the next run re-reports " +
        "the window rather than inheriting a gap. Read-only.",
      arguments: z.object({
        list: z.string().min(1, "list must not be empty").describe(
          "List GUID, URL name, or display name",
        ),
        scopeColumn: z.string().optional().describe(
          "Internal column name to scope the watch to, e.g. Status or " +
            "Assigned_x0020_To. Sent to Graph as a $filter and re-checked " +
            "client-side, and it is what lets a missing row be told apart " +
            "from a row edited out of scope. Goes with scopeValue.",
        ),
        scopeValue: z.string().optional().describe(
          "Value scopeColumn must equal, matched exactly. Omit both to watch " +
            "every row, which means a baseline the size of the whole list.",
        ),
        filter: z.string().optional().describe(
          "Extra OData $filter, ANDed with the scope, e.g. " +
            "fields/Archived ne 1. Narrows the fetch only — rows it excludes " +
            "are indistinguishable from rows that are gone, so prefer " +
            "scopeColumn for anything that changes over a row's life.",
        ),
        trackFields: z.array(z.string()).optional().describe(
          "Internal column names to diff, to the exclusion of all others. " +
            "Omit to diff every column except SharePoint's own bookkeeping.",
        ),
        ignoreFields: z.array(z.string()).default([]).describe(
          "Internal column names to add to the ignored set — for a column " +
            "that a workflow or a calculation rewrites on its own schedule " +
            "and that would otherwise report an edit on every run",
        ),
        contextFields: z.array(z.string()).default([]).describe(
          "Internal column names whose current values ride along on every " +
            "reported row, in `context`. A modified row lists only the " +
            "columns that moved, so name the two or three that say what a row " +
            "IS — a status, an owner, a reference — and the report reads " +
            "without a second lookup.",
        ),
        summarizeColumn: z.string().optional().describe(
          "Internal column name to count the current population by, reported " +
            "as `byValue` busiest first. Context for the changes — the mix a " +
            "status column is sitting in, not part of the diff.",
        ),
        updateBaseline: z.boolean().default(true).describe(
          "Whether a clean run replaces the baseline. False shows the same " +
            "window again without consuming it, which is how to re-read a " +
            "report or try a different trackFields. A scheduled run must " +
            "leave this true or every run re-reports the same changes.",
        ),
        top: z.number().int().min(1).max(5000).default(1000).describe(
          "Page size requested from Graph. Larger than list_items' default " +
            "because the whole population is walked on every run, so fewer " +
            "round trips means less throttling.",
        ),
        maxPages: z.number().int().min(1).max(100).default(20).describe(
          "Cap on @odata.nextLink pages followed. A walk that hits this has " +
            "not seen the whole population, so the baseline is held back.",
        ),
      }),
      execute: async (
        args: {
          list: string;
          scopeColumn?: string;
          scopeValue?: string;
          filter?: string;
          trackFields?: string[];
          ignoreFields: string[];
          contextFields: string[];
          summarizeColumn?: string;
          updateBaseline: boolean;
          top: number;
          maxPages: number;
        },
        context: Context,
      ) => {
        const startMs = Date.now();

        // Both halves of the scope pair or neither. A column with no value has
        // nothing to match against and a value with no column has nothing to
        // match on, and the quiet reading of either — watch the whole list —
        // silently builds a baseline the caller believed was scoped.
        if (
          (args.scopeColumn === undefined) !== (args.scopeValue === undefined)
        ) {
          throw new Error(
            "scopeColumn and scopeValue go together: pass both to watch one " +
              "slice of the list, or neither to watch every row.",
          );
        }
        if (
          args.scopeColumn !== undefined &&
          !INTERNAL_COLUMN_NAME.test(args.scopeColumn)
        ) {
          throw new Error(
            `scopeColumn "${args.scopeColumn}" is not an internal column ` +
              "name. Use the internal name from `get_list` — spaces and " +
              "punctuation are escaped there (Assigned_x0020_To), so the " +
              "display name will not do.",
          );
        }

        // Not interpolated into a filter, but a display name here silently
        // counts every row under "(empty)" — a wrong answer that looks like a
        // real one, which is worth refusing for the same reason.
        if (
          args.summarizeColumn !== undefined &&
          !INTERNAL_COLUMN_NAME.test(args.summarizeColumn)
        ) {
          throw new Error(
            `summarizeColumn "${args.summarizeColumn}" is not an internal ` +
              "column name. Use the internal name from `get_list`.",
          );
        }

        const scopeColumn = args.scopeColumn ?? null;
        const scopeValue = args.scopeValue ?? null;
        const track = args.trackFields && args.trackFields.length > 0
          ? new Set(args.trackFields)
          : null;
        const ignore = new Set([...BOOKKEEPING_FIELDS, ...args.ignoreFields]);
        const fieldPolicy = fieldPolicyKey(args.trackFields, args.ignoreFields);

        context.logger.info(
          "Diffing list {list} on site {site} against its baseline " +
            "(scope={scope}, top={top}, maxPages={maxPages})",
          {
            list: args.list,
            site: context.globalArgs.siteHostPath,
            scope: scopeColumn ? `${scopeColumn}='${scopeValue}'` : "every row",
            top: args.top,
            maxPages: args.maxPages,
          },
        );

        const accessToken = await getAccessToken(context.globalArgs);
        const site = await resolveSite(
          accessToken,
          context.globalArgs.siteHostPath,
        );
        const list = await resolveList(accessToken, site.id, args.list);
        const listWebUrl = list.webUrl ?? null;
        const labels = await columnLabels(
          accessToken,
          site.id,
          list.id,
          args.maxPages,
        );

        // Parenthesised when both halves are present: OData's `and` binds
        // tighter than `or`, so ANDing a caller filter that contains a
        // top-level `or` onto the scope would otherwise silently regroup it
        // into `scope and x or y` and widen the population.
        const scopeFilter = scopeColumn
          ? `fields/${scopeColumn} eq '${
            (scopeValue ?? "").replace(/'/g, "''")
          }'`
          : undefined;
        const filterParts = [scopeFilter, args.filter].filter(
          (part): part is string => part !== undefined && part !== "",
        );
        const filter = filterParts.length === 0
          ? undefined
          : filterParts.length === 1
          ? filterParts[0]
          : filterParts.map((part) => `(${part})`).join(" and ");

        const params: Record<string, string> = {
          "$expand": "fields",
          "$top": String(args.top),
        };
        if (filter) params["$filter"] = filter;

        const result = await withGraphContext(
          `Failed to read items from "${args.list}" for a change diff`,
          () =>
            graphRequestPaginated<GraphListItem>(
              accessToken,
              `/sites/${site.id}/lists/${list.id}/items`,
              params,
              filter ? { "Prefer": HONOR_NON_INDEXED } : undefined,
              fetch,
              args.maxPages,
            ),
        );

        // Graph is free to ignore a filter on an unindexed column — that is
        // what the Prefer header consents to. Re-checking the scope row by row
        // keeps the population correct either way; an ignored filter costs a
        // bigger fetch and nothing else.
        const current = new Map<string, FetchedRow>();
        const counts = new Map<string, number>();
        for (const raw of result.items) {
          const row = changeRow(
            raw,
            listWebUrl,
            scopeColumn,
            ignore,
            track,
            args.contextFields,
          );
          if (scopeColumn && row.scopeValue !== scopeValue) continue;
          current.set(row.itemId, row);
          if (args.summarizeColumn) {
            // Counted after the scope re-check, so the mix describes the
            // population being watched rather than whatever Graph returned.
            const value =
              fieldToString((raw.fields ?? {})[args.summarizeColumn]) ||
              "(empty)";
            counts.set(value, (counts.get(value) ?? 0) + 1);
          }
        }
        const filterHonored = !filter || result.items.length === current.size;

        const instances = changeInstances(list, scopeColumn, scopeValue);
        const stored = await context.readResource(instances.baseline);
        const previous = stored as Baseline | null;

        const fetchedAt = new Date().toISOString();
        const added: ChangedRow[] = [];
        const modified: ChangedRow[] = [];
        const touched: ChangedRow[] = [];
        const removed: RemovedRow[] = [];
        const blindFields: string[] = [];
        let fieldChanges = 0;
        let firstRun = false;
        let since: string | null = null;
        let suspectRemoval = false;

        // A baseline for a different list or a different scope cannot be
        // diffed against this population — every row would read as both added
        // and removed. The instance name already separates them, but a list
        // rename moves the name onto a different population, so the stored
        // baseline says what it covers and that is what gets checked. A
        // mismatch is treated exactly like no baseline at all: report nothing
        // as changed, re-establish, and say so.
        const usable = previous !== null && Array.isArray(previous.rows) &&
          String(previous.listId) === list.id &&
          (previous.scopeColumn ?? null) === scopeColumn &&
          (previous.scopeValue ?? null) === scopeValue &&
          String(previous.fieldPolicy ?? "") === fieldPolicy;

        if (!usable) {
          firstRun = true;
          if (previous !== null) {
            context.logger.warn(
              "Baseline under {instance} covers list {hadList} scoped to " +
                "{hadScope} under columns {hadPolicy}, not {wantList} scoped " +
                "to {wantScope} under {wantPolicy} — re-establishing rather " +
                "than diffing across populations",
              {
                instance: instances.baseline,
                hadList: previous.listId,
                hadScope: previous.scopeColumn
                  ? `${previous.scopeColumn}='${previous.scopeValue}'`
                  : "every row",
                hadPolicy: previous.fieldPolicy ?? "(none recorded)",
                wantList: list.id,
                wantScope: scopeColumn
                  ? `${scopeColumn}='${scopeValue}'`
                  : "every row",
                wantPolicy: fieldPolicy,
              },
            );
          }
        } else {
          since = String(previous.capturedAt);
          const baseRows = new Map(
            previous.rows.map((row) => [String(row.itemId), row]),
          );

          for (const [itemId, row] of current) {
            const before = baseRows.get(itemId);
            if (!before) {
              added.push(reportRow(row, setChanges(row.fields, labels)));
              continue;
            }
            const changes = diffFields(
              before.fields ?? {},
              row.fields,
              labels,
              blindFields,
            );
            if (changes.length > 0) {
              fieldChanges += changes.length;
              modified.push(reportRow(row, changes));
            } else if (
              String(before.lastModified) !== row.lastModified ||
              (before.version ?? null) !== row.version
            ) {
              touched.push(reportRow(row, []));
            }
          }

          const missing = previous.rows.filter(
            (row) => !current.has(String(row.itemId)),
          );

          // Mass absence is a fetch problem, not news. Flagged before the rows
          // are classified, so a run that lost its filter does not spend three
          // hundred Graph calls proving that rows it could not see are still
          // there.
          suspectRemoval = missing.length >= REMOVAL_ALARM_FLOOR &&
            missing.length >= previous.rows.length * REMOVAL_ALARM_SHARE;

          for (const row of missing) {
            const entry: RemovedRow = {
              itemId: String(row.itemId),
              title: row.title ?? null,
              lastSeenAt: String(previous.capturedAt),
              lastModified: String(row.lastModified),
              itemUrl: row.itemUrl ??
                (listWebUrl
                  ? `${listWebUrl}/DispForm.aspx?ID=${String(row.itemId)}`
                  : null),
              // Overwritten by the probe below unless it is skipped, which
              // is a different thing from a probe that ran and failed.
              fate: "unprobed",
              nowInScope: null,
              lastKnownFields: row.fields ?? {},
            };
            if (!suspectRemoval && removed.length < MAX_REMOVAL_PROBES) {
              const probed = await probeRemovedItem(
                accessToken,
                site.id,
                list.id,
                String(row.itemId),
                scopeColumn,
                scopeValue,
              );
              entry.fate = probed.fate;
              entry.nowInScope = probed.nowInScope;
            }
            removed.push(entry);
          }
        }

        const newestFirst = (a: ChangedRow, b: ChangedRow) =>
          b.lastModified.localeCompare(a.lastModified) ||
          (a.title ?? "").localeCompare(b.title ?? "");
        added.sort(newestFirst);
        modified.sort(newestFirst);
        touched.sort(newestFirst);

        // A row the fetch missed while it was still sitting in scope is not
        // news about that row — it is news about the fetch.
        const stillPresent =
          removed.filter((row) => row.fate === "still-present").length;
        const probeFailed =
          removed.filter((row) => row.fate === "unknown").length;
        const changedTotal = added.length + modified.length +
          (removed.length - stillPresent);

        // The baseline is the next run's only memory, so it is replaced only
        // when this fetch can be trusted. A truncated walk did not see the
        // whole population; a suspect removal means rows vanished for a reason
        // this cannot explain; a still-present row proves the walk came back
        // short. In each case keeping the old baseline means the next run
        // re-reports this window instead of treating the gap as delivered.
        let baselineHeldReason: string | null = null;
        if (result.truncated) {
          baselineHeldReason =
            `the fetch hit its ${args.maxPages}-page cap, so the population ` +
            `is incomplete`;
        } else if (suspectRemoval) {
          baselineHeldReason =
            `${removed.length} of ${
              previous?.rows.length ?? 0
            } baseline rows went missing at once, which is likelier to be a ` +
            `failed filter or a permissions change than ${removed.length} ` +
            `deletions`;
        } else if (stillPresent > 0) {
          baselineHeldReason = `${stillPresent} baseline row(s) missing from ` +
            `this fetch are still in scope, so the page walk came back short`;
        } else if (probeFailed > 0) {
          // A probe that failed — a throttle that outlasted its retries, a
          // timeout — leaves it unknown whether those rows left at all.
          // Advancing the baseline would drop them on the strength of a failed
          // call, and they would never be mentioned again; holding it costs
          // one repeated window. Rows the probe never ran for are excluded
          // deliberately: past the probe budget they would hold the baseline
          // on every run and it would never advance again.
          baselineHeldReason = `${probeFailed} removal probe(s) failed, so ` +
            `whether those rows are gone is not known`;
        } else if (!args.updateBaseline) {
          baselineHeldReason =
            "updateBaseline=false was passed, so this window stays unconsumed";
        }
        const baselineUpdated = baselineHeldReason === null;

        const changesHandle = await context.writeResource(
          "listChanges",
          instances.changes,
          {
            siteId: site.id,
            siteHostPath: context.globalArgs.siteHostPath,
            listId: list.id,
            listName: list.displayName ?? list.name ?? list.id,
            scopeColumn,
            scopeValue,
            filter: filter ?? null,
            baselineInstance: instances.baseline,
            firstRun,
            since,
            sinceGapHours: since
              ? Math.round(
                ((Date.parse(fetchedAt) - Date.parse(since)) / 3_600_000) * 10,
              ) / 10
              : null,
            added,
            modified,
            touched,
            removed,
            redactionBlindFields: [...new Set(blindFields)].sort(),
            suspectRemoval,
            baselineUpdated,
            baselineHeldReason,
            totals: {
              scanned: result.items.length,
              inScope: current.size,
              baselineRows: usable ? previous.rows.length : 0,
              added: added.length,
              modified: modified.length,
              touched: touched.length,
              removed: removed.length,
              fieldChanges,
              stillPresent,
              probeFailed,
              changed: changedTotal,
            },
            summarizeColumn: args.summarizeColumn ?? null,
            byValue: [...counts.entries()]
              .map(([value, count]) => ({ value, count }))
              .sort((a, b) =>
                b.count - a.count || a.value.localeCompare(b.value)
              ),
            filterHonored,
            truncated: result.truncated,
            fetchedAt,
            durationMs: Date.now() - startMs,
            collectedBy: EXTENSION_NAME,
          },
        );

        const handles = [changesHandle];
        if (baselineUpdated) {
          handles.push(
            await context.writeResource("listBaseline", instances.baseline, {
              listId: list.id,
              listName: list.displayName ?? list.name ?? list.id,
              scopeColumn,
              scopeValue,
              fieldPolicy,
              capturedAt: fetchedAt,
              rowCount: current.size,
              rows: [...current.values()].map(baselineRow),
              fetchedAt,
              durationMs: Date.now() - startMs,
              collectedBy: EXTENSION_NAME,
            }),
          );
        }

        if (firstRun) {
          context.logger.info(
            "Baseline established for {list} ({scope}): {rows} row(s). " +
              "Nothing is reported as changed on a first run — the next run " +
              "has something to diff against.",
            {
              list: list.displayName ?? args.list,
              scope: scopeColumn
                ? `${scopeColumn}='${scopeValue}'`
                : "every row",
              rows: current.size,
            },
          );
        } else {
          context.logger.info(
            "{changed} change(s) since {since}: {added} added, {modified} " +
              "modified, {removed} removed ({touched} bumped with nothing " +
              "visible)",
            {
              changed: changedTotal,
              since,
              added: added.length,
              modified: modified.length,
              removed: removed.length,
              touched: touched.length,
            },
          );
        }
        if (suspectRemoval) {
          context.logger.error(
            "{removed} of {base} baseline rows are missing from this fetch — " +
              "baseline held back. Check the scope filter and the token's " +
              "permissions before reading these as deletions.",
            { removed: removed.length, base: previous?.rows.length ?? 0 },
          );
        } else if (!baselineUpdated) {
          context.logger.warn(
            "Baseline not replaced: {reason}. The next run re-reports this " +
              "window.",
            { reason: baselineHeldReason },
          );
        }
        if (!filterHonored) {
          context.logger.warn(
            "Graph ignored the filter on the unindexed column and returned " +
              "{scanned} row(s); the scope was applied client-side instead",
            { scanned: result.items.length },
          );
        }

        return { dataHandles: handles };
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

        context.logger.info(
          "Creating item in list {list} on site {site} with {count} field(s): {fields}",
          {
            list: args.list,
            site: context.globalArgs.siteHostPath,
            count: Object.keys(args.fields).length,
            fields: Object.keys(args.fields).join(", "),
          },
        );

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
          itemWriteInstance("create", list, created.id),
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

        context.logger.info(
          "Updating item {id} in list {list} on site {site}, " +
            "{count} field(s): {fields}",
          {
            id: args.itemId,
            list: args.list,
            site: context.globalArgs.siteHostPath,
            count: Object.keys(args.fields).length,
            fields: Object.keys(args.fields).join(", "),
          },
        );

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

        // Graph normally answers this PATCH with the updated field map, but it
        // can answer 204 with no body — graphRequest yields {} then, which is
        // not nullish, so `??` would persist an empty record and lose the only
        // account of what changed. Fall back on emptiness, not nullishness.
        const resultFields = Object.keys(updated).length > 0
          ? updated
          : args.fields;

        const handle = await context.writeResource(
          "itemWrite",
          itemWriteInstance("update", list, args.itemId),
          {
            siteId: site.id,
            siteHostPath: context.globalArgs.siteHostPath,
            operation: "update",
            listId: list.id,
            listName: list.displayName ?? list.name ?? list.id,
            itemId: args.itemId,
            fields: resultFields,
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
        "is a record of what was removed. Idempotent: an item that is already " +
        "gone succeeds with alreadyAbsent=true rather than failing.",
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

        context.logger.info(
          "Deleting item {id} from list {list} on site {site}",
          {
            id: args.itemId,
            list: args.list,
            site: context.globalArgs.siteHostPath,
          },
        );

        const accessToken = await getAccessToken(context.globalArgs);
        const site = await resolveSite(
          accessToken,
          context.globalArgs.siteHostPath,
        );
        const list = await resolveList(accessToken, site.id, args.list);

        const itemPath = `/sites/${site.id}/lists/${list.id}/items/${
          encodeURIComponent(args.itemId)
        }`;
        const instanceName = itemWriteInstance("delete", list, args.itemId);
        const record = {
          siteId: site.id,
          siteHostPath: context.globalArgs.siteHostPath,
          operation: "delete",
          listId: list.id,
          listName: list.displayName ?? list.name ?? list.id,
          itemId: args.itemId,
          collectedBy: EXTENSION_NAME,
        };

        // Read before deleting. A delete that returns 204 tells you nothing
        // about what was in the row, and SharePoint's recycle bin is not
        // reachable from Graph — so the pre-image is the only record.
        let existing: GraphListItem;
        try {
          existing = await graphRequest<GraphListItem>(
            accessToken,
            "GET",
            `${itemPath}?$expand=fields`,
          );
        } catch (e) {
          if (!(e instanceof GraphApiError) || e.statusCode !== 404) {
            throw new Error(
              `Failed to read item "${args.itemId}" in "${args.list}" before ` +
                `delete: ${e instanceof Error ? e.message : String(e)}`,
              { cause: e },
            );
          }

          // Already gone. The requested end state holds, so report success
          // rather than failing a delete with nothing left to do.
          const absent = await context.writeResource(
            "itemWrite",
            instanceName,
            {
              ...record,
              fields: {},
              webUrl: null,
              alreadyAbsent: true,
              deleteConfirmed: true,
              fetchedAt: new Date().toISOString(),
              durationMs: Date.now() - startMs,
            },
          );

          context.logger.info(
            "Item {id} is already absent from {name} — nothing to delete",
            { id: args.itemId, name: list.displayName ?? args.list },
          );
          return { dataHandles: [absent] };
        }

        // Persist the pre-image BEFORE issuing the DELETE. If the process dies
        // between the two, what the row held still survives; deleteConfirmed
        // is what separates that record from a completed one.
        await context.writeResource("itemWrite", instanceName, {
          ...record,
          fields: existing.fields ?? {},
          webUrl: existing.webUrl ?? null,
          alreadyAbsent: false,
          deleteConfirmed: false,
          fetchedAt: new Date().toISOString(),
          durationMs: Date.now() - startMs,
        });

        await withGraphContext(
          `Failed to delete item "${args.itemId}" in "${args.list}"`,
          () =>
            graphRequest<Record<string, never>>(
              accessToken,
              "DELETE",
              itemPath,
            ),
        );

        const handle = await context.writeResource("itemWrite", instanceName, {
          ...record,
          fields: existing.fields ?? {},
          webUrl: existing.webUrl ?? null,
          alreadyAbsent: false,
          deleteConfirmed: true,
          fetchedAt: new Date().toISOString(),
          durationMs: Date.now() - startMs,
        });

        context.logger.info("Deleted item {id} from {name}", {
          id: args.itemId,
          name: list.displayName ?? args.list,
        });
        return { dataHandles: [handle] };
      },
    },
  },

  checks: {
    "site-locator-format": {
      description:
        "siteHostPath is a Graph site locator (host:/path), not a browser URL",
      labels: ["policy"],
      execute: (
        context: CheckContext,
      ): { pass: boolean; errors?: string[] } => {
        const raw = context.globalArgs.siteHostPath;
        const errors: string[] = [];

        if (raw !== raw.trim()) {
          errors.push(
            "siteHostPath has leading or trailing whitespace, which Graph " +
              "carries into the path and fails to resolve.",
          );
        }

        const value = raw.trim();

        // The two ways this argument is habitually got wrong: pasting the
        // browser URL, and dropping the colon that separates host from path.
        // Both resolve as a site Graph has never heard of, so the first real
        // failure is a 404 on the site rather than a word about the locator.
        if (/^https?:\/\//i.test(value)) {
          errors.push(
            `siteHostPath "${raw}" is a browser URL. Graph wants a host:path ` +
              "locator — drop the scheme and put a colon before the path, " +
              "e.g. contoso.sharepoint.com:/sites/Clients.",
          );
        } else if (value.includes("/") && !value.includes(":")) {
          errors.push(
            `siteHostPath "${raw}" has a path but no ":" separating it from ` +
              "the host, so Graph reads the whole string as a hostname. " +
              "Write it as contoso.sharepoint.com:/sites/Clients.",
          );
        }

        return errors.length > 0 ? { pass: false, errors } : { pass: true };
      },
    },

    "write-scope-granted": {
      description:
        "The configured scopes include a SharePoint write scope, so a write can succeed",
      labels: ["policy"],
      appliesTo: ["create_item", "update_item", "delete_item"],
      execute: (
        context: CheckContext,
      ): { pass: boolean; errors?: string[] } => {
        const scopes = context.globalArgs.scopes;

        // A refresh token is bound to the scopes it was issued with, so a
        // read-only deployment reaching a write method cannot succeed — Graph
        // answers `accessDenied` after the site and list have been resolved.
        // Matched loosely because Entra accepts both the bare scope and the
        // fully-qualified https://graph.microsoft.com/... form.
        if (!/sites\.(readwrite|manage|fullcontrol)/i.test(scopes)) {
          return {
            pass: false,
            errors: [
              `The configured scopes ("${scopes}") carry no SharePoint write ` +
              "permission, so this write would be refused by Graph with " +
              "accessDenied. Add Sites.ReadWrite.All to `scopes` and re-run " +
              "`bootstrap` — changing scopes requires a new refresh token.",
            ],
          };
        }

        return { pass: true };
      },
    },

    "credentials-resolve-site": {
      description:
        "The refresh token still exchanges for an access token and siteHostPath resolves",
      labels: ["live"],
      appliesTo: ["create_item", "update_item", "delete_item"],
      execute: async (
        context: CheckContext,
      ): Promise<{ pass: boolean; errors?: string[] }> => {
        let accessToken: string;
        try {
          accessToken = await getAccessToken(context.globalArgs);
        } catch (e) {
          return {
            pass: false,
            errors: [
              `Could not exchange the stored refresh token for an access ` +
              `token: ${e instanceof Error ? e.message : String(e)}. ` +
              "Re-run the `bootstrap` method and store the new refreshToken " +
              "in your vault.",
            ],
          };
        }

        try {
          const site = await resolveSite(
            accessToken,
            context.globalArgs.siteHostPath,
          );
          context.logger.info(
            "Pre-flight: credentials valid, site {site} resolved to {id}",
            {
              site: context.globalArgs.siteHostPath,
              id: site.id,
            },
          );
        } catch (e) {
          return {
            pass: false,
            errors: [
              `Credentials are valid but siteHostPath ` +
              `"${context.globalArgs.siteHostPath}" did not resolve: ${
                e instanceof Error ? e.message : String(e)
              }. Check the locator, and that the signed-in user can open ` +
              `that site.`,
            ],
          };
        }

        return { pass: true };
      },
    },
  },
};
