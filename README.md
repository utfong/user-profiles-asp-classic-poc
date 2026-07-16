# ASP Classic Playground

A sandbox for Classic ASP proofs of concept, one subfolder per IIS virtual host. Each host is independent: its own `web.config`, its own site binding, no shared code between them.

## Hosts

| Host | Language | Status |
|---|---|---|
| [`asp-vbscript/`](asp-vbscript/README.md) | VBScript | Done: User Profiles CRUD API over a session-scoped Microsoft Access database, plain HTML/CSS + axios frontend, CSRF-protected writes, and a self-contained test suite. See its [README](asp-vbscript/README.md) for setup, the API reference, and how to run the tests. |
| [`asp-jscript/`](asp-jscript/README.md) | JScript | Done: same User Profiles CRUD API, built as a genuine JScript port (not a syntax translation) rather than sharing code with `asp-vbscript`. Shares its frontend verbatim (backend-agnostic). Needs `jscript-engine-fix.cmd enable` (repo root) run once to work around a real Windows 11 25H2 JScript engine bug; see its [README](asp-jscript/README.md#known-issues-resolved) for details. |

## Setup

Each host needs to be bound as its own IIS site/virtual host with the Classic ASP feature enabled. Host-specific prerequisites (database engine, handler mapping, etc.) are documented in that host's own README. Start there.

`asp-classic.code-workspace` opens both folders together in VS Code.

For AI assistants working in this repo: see [`CLAUDE.md`](CLAUDE.md) for environment gotchas discovered while building the vbscript host that will very likely resurface building others.
