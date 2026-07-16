# CLAUDE.md: asp-classic

This repo is a sandbox of Classic ASP proofs of concept, one subfolder per IIS virtual host:

- **`asp-vbscript/`**: VBScript host. User Profiles CRUD API + session-scoped Access database. See [`asp-vbscript/CLAUDE.md`](asp-vbscript/CLAUDE.md) for conventions specific to that code, and [`asp-vbscript/README.md`](asp-vbscript/README.md) for setup/usage/API reference.
- **`asp-jscript/`**: JScript host. Same CRUD app, ported natively (not translated) to JScript. See [`asp-jscript/CLAUDE.md`](asp-jscript/CLAUDE.md) for JScript-specific conventions and a documented (now resolved) Windows 11 25H2 JScript engine bug, and [`asp-jscript/README.md`](asp-jscript/README.md) for setup/usage/API reference. Needs `jscript-engine-fix.cmd enable` (repo root) run once - see below.

Each folder is its own IIS site/virtual host with its own `web.config`. They don't share code or a document root.

## Environment gotchas that apply to any host here

Discovered the hard way building `asp-vbscript/`, on this machine's Windows 11 + IIS setup. Not specific to VBScript: expect these to resurface building `asp-jscript/` too.

- **The `.asp` handler mapping doesn't allow `PUT`/`DELETE` by default.** IIS's built-in ASP handler is normally scoped to `GET,HEAD,POST,DEBUG`. Any host using real REST verbs needs this widened per-site in IIS Manager: *Handler Mappings* → the ASP mapping → *Edit Feature Permissions*/*Request Restrictions* → *Verbs* tab.
- **IIS already blocks `.asa` at the server level by default** (`requestFiltering`). Adding an explicit `<add fileExtension=".asa" allowed="false" />` in a site's `web.config` collides with that (`0x800700b7` duplicate key) and 500s the *entire site*, not just `.asa` requests. Don't add it.
- **`global.asa` cannot contain `<% %>` script blocks at all**: not directly, not via `#include` of a file that uses them (`ASP 0137 Invalid Global Script`). Only `<script language="VBScript" runat="server">` blocks are allowed there, and only for `Application_OnStart/OnEnd`/`Session_OnStart/OnEnd`. Shared library code needed inside `global.asa` has to be written in that script-block style, or simpler: inlined directly rather than shared via include.
- **Access (Jet/ACE) `TEXT` columns cap at 255 characters.** `TEXT(500)` etc. fails at `CREATE TABLE` time ("Size of field is too long"). Long text needs `MEMO` (ACE also accepts `LONGTEXT`), paired with ADO parameter type `adLongVarWChar` (203), not `adVarWChar` (202) with a size.
- **Provider bitness must match the app pool.** This machine has the Access Database Engine 2016 Redistributable installed 64-bit, matching the (default) 64-bit app pools, using `Microsoft.ACE.OLEDB.12.0`. If a new host's app pool is 32-bit, either the 32-bit ACE provider or the built-in `Microsoft.Jet.OLEDB.4.0` (older, `.mdb` only) is needed instead.

Discovered the hard way building `asp-jscript/` instead - not specific to JScript, could resurface on any future host:

- **This machine's classic ASP feature has no `web.config` schema for `<system.webServer><asp>`** (`%windir%\System32\inetsrv\config\schema\asp_schema.xml` doesn't exist - only `ASPNET_schema.xml`, a different feature, for ASP.NET). Adding an `<asp>` element to a site's `web.config` here (e.g. to try `<cache scriptEngineCacheMax="0" />`) doesn't just fail silently - it breaks that site's entire config parse ("New Application Failed" on every request) and can trip IIS's Rapid-Fail Protection, stopping the app pool. If a future host seems to need an ASP-level `web.config` setting, check this file exists first, or make the change via IIS Manager's GUI instead (which may use a different mechanism than the web.config schema).
- **Windows 11 24H2/25H2 has a known Microsoft regression breaking classic ASP under JScript** (not VBScript): the OS defaults to a newer `jscript9legacy.dll` engine that isn't fully compatible with classic ASP's script-engine-reuse model, causing intermittent `500`s with `ASP 0240` / `C0000005` in `CActiveScriptEngine::ReuseEngine()`. **Resolved on this machine** via `jscript-engine-fix.cmd enable` (repo root, needs an elevated prompt) - sets a machine-wide `JScriptReplacement` registry value and restarts IIS. Two other Microsoft-documented registry fixes were tried first and did *not* work (see [`asp-jscript/CLAUDE.md`](asp-jscript/CLAUDE.md) for the full investigation and why this third one is the one that stuck). Any future JScript host on this machine benefits automatically once the fix is enabled, since it's machine-wide, not per-site - check `jscript-engine-fix.cmd status` first before assuming it's on.
- **`iisreset` can fail mid-restart on this machine** ("Access denied, you must be an administrator of the remote computer to use this command"), even when genuinely elevated, leaving IIS's `W3SVC`/`WAS` services stopped with no automatic recovery (discovered via the registry-key testing above). Prefer `net stop w3svc` / `net start w3svc` over `iisreset` here, and verify `Get-Service W3SVC` actually shows `Running` afterward rather than assuming a restart succeeded.

## Working across hosts

When porting a pattern from `asp-vbscript/` to `asp-jscript/` (or a new host), the JScript syntax will differ but the IIS-level gotchas above are identical. Don't re-derive them from scratch.
