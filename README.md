# Swamp Extensions

Extensions for [swamp](https://github.com/swamp-club/swamp) published under the
`@aspec451` collective.

## Model Extensions

| Extension                                                             | Description                                                                                                                                                             | Dependencies      |
| --------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------- |
| [`@aspec451/microsoft/sharepoint-lists`](microsoft/sharepoint-lists/) | Read and write SharePoint **list items** via the Microsoft Graph list surface (`/sites/{site}/lists/...`), using a public client app registration with device code flow | None (uses fetch) |

## Installation

Extensions are installed automatically when referenced in a swamp repository,
or manually with:

```bash
swamp extension pull @aspec451/microsoft/sharepoint-lists
```

## Repository Layout

Each extension lives in its own directory and is a self-contained swamp
repository:

```
<vendor>/<extension>/
├── manifest.yaml       # registry metadata (name, version, models, labels)
├── .swamp.yaml         # swamp repo marker
├── deno.json           # check / lint / fmt / test tasks
├── README.md           # per-extension docs
├── LICENSE.md          # Apache-2.0
└── extensions/
    ├── models/         # export const model
    └── reports/        # export const report
```

CI derives its job matrix from the filesystem — every directory containing a
`manifest.yaml` is discovered and checked independently, so adding an extension
needs no workflow change.

## Development

```bash
cd microsoft/sharepoint-lists

deno task check      # type-check
deno task lint
deno task fmt:check
deno task test
```

Swamp bundles its own Deno at `~/.swamp/deno/deno` (find it with
`swamp doctor extensions --json` → `denoPath`) if you don't have Deno on
`PATH`.

## Publishing

```bash
cd microsoft/sharepoint-lists

swamp extension fmt manifest.yaml --check
swamp extension version --manifest manifest.yaml --json
swamp extension quality manifest.yaml --json
swamp extension push manifest.yaml --dry-run
swamp extension push manifest.yaml
```

`push` will not proceed cleanly without a content-hash-bound adversarial review
report for the exact code being pushed. The dry run prints the report path and a
fill-in skeleton. See [CLAUDE.md](CLAUDE.md#publishing).

## Support and Warranty

These are free, best-effort extensions provided as-is under the Apache-2.0
license, with no warranty and no SLA. See [SUPPORT.md](SUPPORT.md) for how to
get help and [SECURITY.md](SECURITY.md) for reporting a vulnerability.

## License

[Apache-2.0](LICENSE).

The Graph auth and fetch helpers under
`microsoft/sharepoint-lists/extensions/models/microsoft/_lib/` are adapted from
[`@webframp/microsoft/teams`](https://github.com/webframp/swamp-extensions),
Copyright 2026 Sean Escriva, licensed under the Apache License 2.0. Changes from
the originals are noted in those files' headers.
