# microsoft/sharepoint-lists

Read and write SharePoint **list items** via the Microsoft Graph API, using a
public client app registration with device code flow — no client secret.

Modeled on [`@webframp/microsoft/teams`][teams] and sharing its `_lib/auth.ts`
and `_lib/graph.ts` helpers (adapted, attribution in the file headers). The
registry's existing SharePoint coverage (`@dougschaefer/ms-graph-sharepoint`) is
document-library only (`/drive/...`) and read-only; this covers the list surface
(`/sites/{site}/lists/...`) and can write.

[teams]: https://github.com/webframp/swamp-extensions

## Methods

| Method        | Description                                                      |
| ------------- | ---------------------------------------------------------------- |
| `bootstrap`   | Device code flow auth; outputs refresh token for the vault       |
| `list_lists`  | Enumerate lists on the site — start here to get list identifiers |
| `get_list`    | One list's column definitions and which columns accept writes    |
| `list_items`  | Fetch items with fields expanded; supports `$filter`/`$orderby`  |
| `create_item` | Create one item                                                  |
| `update_item` | PATCH one item's fields                                          |
| `delete_item` | Delete one item; requires `confirm=true`                         |

## Prerequisite: app registration

This model needs an Entra **public client** app registration with delegated
scopes `offline_access`, `User.Read`, and `Sites.ReadWrite.All`
(`Sites.Read.All` if you only need reads — set `scopes` accordingly). Device
code flow requires "Allow public client flows" enabled on the registration.

Delegated, not application: every call runs as the signed-in user, so the model
can never reach a list that user cannot already open, and list validation and
item-level permissions apply normally to writes.

## Setup

```bash
swamp vault put my-vault sharepointTenantId "e9b2b7ba-b238-42a9-b271-2adfc82da650"
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

## Notes and gotchas

**`list` accepts three forms.** GUID, list URL name, or display name. Graph's
own `/lists/{id-or-name}` accepts only the first two, and the display name
diverges from the URL name whenever a list has been renamed — so a miss falls
back to enumerating the site's lists and matching the display name. A failed
resolve lists the available lists in the error.

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

**`list_items` truncation is a floor, not a total.** Pagination stops at
`maxPages` (default 20). When `truncated: true`, `totalFetched` is a lower bound
— raise `maxPages` or narrow with `filter`.

**`delete_item` reads before deleting.** A DELETE returns 204 and Graph cannot
reach SharePoint's recycle bin, so the pre-delete field values persisted in the
`itemWrite` resource are the only record of what was removed.

**`invalid_grant` on any method** means the refresh token expired (90 days of
inactivity) or was revoked. Re-run `bootstrap` and update the vault. Changing
`scopes` also requires a re-bootstrap — a refresh token is bound to the scopes
it was issued with.
