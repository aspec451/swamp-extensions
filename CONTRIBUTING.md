# Contributing to swamp-extensions

This is a personal, single-maintainer collection of swamp extensions. Issues
are genuinely welcome; pull requests are accepted only occasionally.

## Reporting bugs and requesting features

Open a [GitHub
issue](https://github.com/aspec451/swamp-extensions/issues). For a bug, include:

- The extension name and version (`swamp extension info <package> --json`)
- The method and inputs you ran
- What you expected and what actually happened, with the error text

For a feature, describe the operation you need and the API it maps to. An issue
that names the endpoint and the shape of the data is most of the work.

## Why pull requests are limited

Every extension here is published under the `@aspec451` collective, so anything
merged is code the maintainer signs and ships to the registry under their name.
Reviewing an external contribution to that standard usually costs more than
writing it, and the adversarial review gate (below) has to be redone regardless.

A well-described issue is more useful than an unsolicited PR.

## For trusted contributors

If a PR has been invited, it must arrive green:

```bash
cd <vendor>/<extension>

deno task check
deno task lint
deno task fmt:check
deno task test

swamp extension fmt manifest.yaml --check
swamp extension quality manifest.yaml --json
swamp extension push manifest.yaml --dry-run
```

Non-trivial code changes also need a fresh adversarial review — the report is
bound to a content hash, so any source edit invalidates the previous one. See
[CLAUDE.md](CLAUDE.md#publishing).

Don't bump the version in `manifest.yaml`; the maintainer does that at publish
time.

## Licensing

By contributing you agree your contribution is licensed under
[Apache-2.0](LICENSE), the license this repository ships under.

If you adapt code from another Apache-2.0 project, keep its attribution and add
a note of your changes in the file header — as
`microsoft/sharepoint-lists/extensions/models/microsoft/_lib/` does for
`@webframp/microsoft/teams`.

## Code of conduct

See [CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md).
