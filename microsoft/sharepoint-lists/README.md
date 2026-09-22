# @aspec451/microsoft/sharepoint-lists

Read and write SharePoint **list items** via the Microsoft Graph API, using a
public client app registration with device code flow — no client secret.

The `_lib/auth.ts` and `_lib/graph.ts` helpers are adapted from
[`@webframp/microsoft/teams`][teams], which is Apache-2.0; attribution and the
list of changes are in those file headers. The registry's existing SharePoint
coverage (`@dougschaefer/ms-graph-sharepoint`) is document-library only
(`/drive/...`) and read-only; this covers the list surface
(`/sites/{site}/lists/...`) and can write.

[teams]: https://github.com/webframp/swamp-extensions

## Methods

| Method         | Description                                                          |
| -------------- | -------------------------------------------------------------------- |
| `bootstrap`    | Device code flow auth; outputs refresh token for the vault           |
| `list_lists`   | Enumerate lists on the site — start here to get list identifiers     |
| `get_list`     | One list's column definitions and which columns accept writes        |
| `list_items`   | Fetch items with fields expanded; supports `$filter`/`$orderby`      |
| `list_changes` | What moved since the last run — column-level diff against a baseline |
| `create_item`  | Create one item                                                      |
| `update_item`  | PATCH one item's fields                                              |
| `delete_item`  | Delete one item; requires `confirm=true`                             |

## Prerequisite: app registration

This model needs an Entra **public client** app registration with delegated
scopes `offline_access`, `User.Read`, and `Sites.ReadWrite.All`
(`Sites.Read.All` if you only need reads — set `scopes` accordingly). Keep
`offline_access` in any override: without it Entra issues no refresh token and
`bootstrap` fails rather than reporting a success you cannot store. Device
code flow requires "Allow public client flows" enabled on the registration.

Delegated, not application: every call runs as the signed-in user, so the model
can never reach a list that user cannot already open, and list validation and
item-level permissions apply normally to writes.

## Setup

```bash
swamp vault put my-vault sharepointTenantId "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee"
swamp vault put my-vault sharepointClientId "11111111-2222-3333-4444-555555555555"
swamp vault put my-vault sharepointRefreshToken "placeholder"

swamp model create @aspec451/microsoft/sharepoint-lists my-sp-lists
# then set globalArguments in the model definition:
#   tenantId:     ${{ vault.get(my-vault, sharepointTenantId) }}
#   clientId:     ${{ vault.get(my-vault, sharepointClientId) }}
#   refreshToken: ${{ vault.get(my-vault, sharepointRefreshToken) }}
#   siteHostPath: contoso.sharepoint.com:/sites/YourSite

swamp model method run my-sp-lists bootstrap
# follow the device code prompt, then store the token it outputs:
swamp vault put my-vault sharepointRefreshToken "<refreshToken from output>"
```

## Usage

```bash
# What lists are on the site?
swamp model method run my-sp-lists list_lists

# What columns does one have, and which are writable?
swamp model method run my-sp-lists get_list --input list="Project Tracker"

# Read items, filtered and sorted
swamp model method run my-sp-lists list_items \
  --input list="Project Tracker" \
  --input filter="fields/Status eq 'Open'" \
  --input orderBy="fields/Created desc"

# What changed since the last time this ran?
swamp model method run my-sp-lists list_changes \
  --input list="Project Tracker" \
  --input scopeColumn=Status --input scopeValue=Open

# Create
swamp model method run my-sp-lists create_item \
  --input list="Project Tracker" \
  --input fields='{"Title":"New request","Status":"Open"}'

# Update
swamp model method run my-sp-lists update_item \
  --input list="Project Tracker" \
  --input itemId=42 \
  --input fields='{"Status":"Closed"}'

# Delete (confirm required)
swamp model method run my-sp-lists delete_item \
  --input list="Project Tracker" --input itemId=42 --input confirm=true

# Query stored data
swamp data query my-sp-lists 'attributes.totalFetched > 0'
```

## Tracking changes

`list_changes` answers the one question `list_items` cannot: not *which rows
were touched* but *what moved*. Graph returns every row's
`lastModifiedDateTime`, but a bump is not a diff — it cannot say
`Status: In Progress → Confirmed` — and a deleted row has no timestamp left to
filter on and no row left to return.

So the method keeps a baseline of its own. Each run reads the baseline the last
clean run wrote, fetches the population again, and reports the difference
column by column: rows `added`, rows `modified` with each column's `from` and
`to`, rows `touched` (SharePoint bumped them and nothing visible moved — an
attachment, a permission, a re-save), and rows `removed`, each probed and
classified as `deleted`, `rescoped`, or `still-present` — or `unknown` when the
probe itself failed, and `unprobed` when it was deliberately skipped.

```bash
# First run establishes the baseline and reports nothing as changed
swamp model method run my-sp-lists list_changes --input list="Project Tracker"

# Scope it to one slice, so each slice diffs against its own population
swamp model method run my-sp-lists list_changes \
  --input list="Project Tracker" \
  --input scopeColumn=Assigned_x0020_To --input scopeValue="Platform Team"

# Carry a row's current state alongside the diff, and the mix it sits in
swamp model method run my-sp-lists list_changes \
  --input list="Project Tracker" \
  --input contextFields='["Status","Assigned_x0020_To"]' \
  --input summarizeColumn=Status

# Read the same window again without consuming it
swamp model method run my-sp-lists list_changes \
  --input list="Project Tracker" --input updateBaseline=false

# Is there anything to read?
swamp data query my-sp-lists 'attributes.totals.changed > 0'
```

A modified row lists only the columns that moved, which is the point but can
leave a reader without the state the row is in now — so `contextFields` names
the two or three columns that say what a row *is* (a status, an owner, a
reference) and carries their current values on every reported row as `context`.
`summarizeColumn` counts the watched population by one column into `byValue`,
busiest first: context for the changes rather than part of them.

A baseline, not a time window, because a window has to guess the gap between
runs: a fixed lookback either double-reports or drops changes, and a run that
fails silently widens the gap it cannot see. The baseline has no seam — `since`
is the last *successful* run, so whatever a failed run missed is still pending
and turns up on the next one.

That only holds if the baseline is trustworthy, so it is replaced only when the
fetch is. A truncated page walk, a mass absence (`suspectRemoval`), a missing
row the probe finds still sitting in scope, or a probe that failed outright all
mean this run cannot vouch for the population; the changes are still reported,
`baselineUpdated` is `false`, `baselineHeldReason` says why, and the next run
re-reports the window rather than inheriting the gap. Assert on
`suspectRemoval` in a workflow if you want a failed filter to fail the run
instead of announcing that the department was disbanded.

## Pre-flight checks

Three checks run before the write methods. Each is skippable by name or label
(`--skip-check <name>`, `--skip-check-label live`, `--skip-checks`).

| Check                     | Label    | Applies to                                  | Validates                                                                       |
| ------------------------- | -------- | ------------------------------------------- | ------------------------------------------------------------------------------- |
| `site-locator-format`     | `policy` | every method                                | `siteHostPath` is a `host:/path` locator, not a pasted browser URL              |
| `write-scope-granted`     | `policy` | `create_item`, `update_item`, `delete_item` | `scopes` carries a SharePoint write permission                                  |
| `credentials-resolve-site`| `live`   | `create_item`, `update_item`, `delete_item` | the refresh token still exchanges, and `siteHostPath` resolves to a real site   |

`credentials-resolve-site` is the one that costs a round trip — it exists
because a write otherwise resolves the site and the list before Graph refuses
it, and it distinguishes an expired refresh token (re-run `bootstrap`) from a
locator that does not resolve (fix `siteHostPath`). Skip it with
`--skip-check-label live` in an offline or replayed environment.

## Notes and gotchas

**`list` accepts three forms.** GUID, list URL name, or display name. Graph's
own `/lists/{id-or-name}` accepts only the first two, and the display name
diverges from the URL name whenever a list has been renamed — so a miss falls
back to enumerating the site's lists and matching the display name. A failed
resolve lists the available lists in the error. Only a 404 or 400 on the direct
lookup earns that fallback — a 401, 403 or 429 is reported as itself rather
than recast as a missing list.

**Fields are keyed by INTERNAL column name, not display name.** A column shown
as "Assigned To" is usually `AssignedTo`, and a renamed column keeps its
original internal name. Run `get_list` and read `writableColumns` before a
write; read-only columns must be omitted entirely.

**Filtering an unindexed column is best-effort.** Graph refuses `$filter` and
`$orderby` on list columns SharePoint has not indexed unless the caller sends
`Prefer: HonorNonIndexedQueriesWarningMayFailRandomly`. The model sends it
whenever a filter or sort is supplied, which converts a flat HTTP 400 into a
call that usually works but may fail intermittently on a large list. Index the
column in SharePoint if a filter matters operationally.

**Truncation is a floor, not a total.** `list_items`, `get_list` and
`list_changes` all stop paginating at `maxPages` (default 20) and report
`truncated`. When it is true,
`list_items`' `totalFetched` is a lower bound — raise `maxPages` or narrow with
`filter`. For `get_list` it means `writableColumns` may be incomplete, so treat
a write built from it as unverified until the walk completes.

**`list_changes` keeps two data instances per list and scope.** The report
(`listChanges`) and the baseline behind it (`listBaseline`, named
`baseline-<list>-<scope>`), so two scopes on one list never diff against each
other's population. The baseline records the list, the scope, and the tracked
column set it was captured under; change any of them — including `trackFields`
or `ignoreFields`, which decide what a stored row even carries — and the run is
treated as a first run and re-establishes, rather than reporting every column
the new policy drops as cleared.

**Scoping beats filtering for anything that changes.** `scopeColumn` and
`scopeValue` are sent as a `$filter`, re-checked client-side, and are what let
a row edited out of scope be told apart from one that was deleted. Rows
excluded by the freeform `filter` are indistinguishable from rows that are
gone, so keep it to things that do not change over a row's life.

**Bookkeeping columns are never diffed.** `Modified`, `_UIVersionString`,
`@odata.etag`, `EditorLookupId`, the `LinkTitle*` mirrors and the
`_Compliance*` set all move on their own and would report an edit on every row
saved for any reason; the version and timestamp are carried as corroboration of
a bump instead. Every other column is diffed, including ones added to the list
later — the field set comes from the rows themselves. Add your own with
`ignoreFields`, or name an explicit set with `trackFields`.

**A vault secret inside list content shows up as `***`.** swamp scrubs vault
values out of resource content on the way to disk, and a tenant GUID pasted
into a rich-text column as part of a Teams or SharePoint deep link is the
common case. The stored baseline then carries `***` where a fresh fetch has the
real characters, which a naive comparison reports as a change on every run
forever. A redacted span is treated as unknowable rather than different — the
literal text around it still has to match, so an edit elsewhere in the same
column is still caught — and every column it applied to is listed in
`redactionBlindFields`.

**`delete_item` reads before deleting, and is idempotent.** A DELETE returns 204
and Graph cannot reach SharePoint's recycle bin, so the pre-delete field values
persisted in the `itemWrite` resource are the only record of what was removed.
That pre-image is written *before* the DELETE is issued and rewritten after,
so an interruption between the two leaves the record behind rather than
nothing — `deleteConfirmed` separates the two states. An item that is already
gone succeeds with `alreadyAbsent: true` instead of failing, so a retried
delete is safe.

**Throttled and transient responses retry automatically.** Graph throttles the
list surface aggressively, and a long paginated walk is what trips it. A 429,
503 or 504 is retried up to three times, honouring `Retry-After` when present
and backing off exponentially from one second when it is not, capped at 30s per
wait. Retries happen per request, so a throttle partway through a paginated
walk no longer discards the pages already collected. Other statuses — including
404 and 403 — fail immediately.

**`invalid_grant` on any method** means the refresh token expired (90 days of
inactivity) or was revoked. Re-run `bootstrap` and update the vault. Changing
`scopes` also requires a re-bootstrap — a refresh token is bound to the scopes
it was issued with.

## License

MIT — see [LICENSE.md](LICENSE.md). The portions of `_lib/auth.ts` and
`_lib/graph.ts` derived from `@webframp/microsoft/teams` remain under the
Apache License 2.0, as noted in those files.
