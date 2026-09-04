# Security Policy

## Reporting a vulnerability

**Don't open a public issue for a security problem.**

Use GitHub's [private vulnerability
reporting](https://github.com/aspec451/swamp-extensions/security/advisories/new)
on this repository, which opens a draft advisory visible only to the maintainer.

Please include:

- The affected extension and version
- What an attacker can do with it, and what access they need to start
- Reproduction steps, or the code path if a repro isn't practical

Describe the class of problem rather than attaching a working exploit.

## What to expect

Best-effort, single-maintainer response. Expect an acknowledgement within a
week. Confirmed issues are fixed in a new version published to the registry,
and the advisory is published once a fixed version is available.

## Scope

**In scope** — anything in a published extension that leaks credentials, sends
data to an unintended destination, executes untrusted input, or grants access
beyond the scopes an extension documents.

**Out of scope** — vulnerabilities in swamp itself (report those to
[swamp-club/swamp](https://github.com/swamp-club/swamp)), in Microsoft Graph or
any other upstream API, or in a downstream repository's own configuration.
Over-broad permissions that an extension documents and requires are a design
tradeoff, not a vulnerability — though an issue arguing the design is wrong is
welcome.
