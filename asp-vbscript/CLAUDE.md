# CLAUDE.md: asp-vbscript

Guidance for working in this specific host. See [`../CLAUDE.md`](../CLAUDE.md) first for repo-wide IIS/Access environment gotchas that apply here too. See [`README.md`](README.md) for the user-facing overview, API reference, and setup steps.

## Architecture in one paragraph

Static HTML/CSS/axios frontend (`index.html`, `assets/`) talks to a pure-JSON ASP API (`api/session.asp`, `api/profiles.asp`). No database ships with the app: `api/session.asp` provisions a session-scoped `.accdb` under `App_Data/sessions/{SessionID}.accdb` on demand, and `global.asa`'s `Session_OnEnd` (plus an explicit `DELETE /api/session.asp`) tears it back down. Writes require an active session + CSRF token; reads don't.

## Conventions to preserve

- **`Option Explicit` placement**: only in entry-point pages (`api/*.asp`, `tests/run.asp`), as the very first statement, before any `#include`. Never add it inside `lib/*.asp`: those get textually spliced into whatever includes them, and a second `Option Explicit` after other code in the merged script is a compile error.
- **ADO type codes are hardcoded integers** (`adInteger=3`, `adDate=7`, `adVarWChar=202`, `adLongVarWChar=203`, `adParamInput=1`) rather than pulled from `adovbs.inc`, since there's no such include in this project. Keep using literals with a comment, don't introduce a dependency on that file.
- **`Response.End` aborts the entire request**, not just the current `Sub`/`Function`: that's why `SendError`/`SendData`/`SendJson` in `lib/http.asp` never need an `Exit Sub` after them to prevent fallthrough. Relied on throughout `api/*.asp`.
- **Error handling pattern**: `On Error Resume Next` around a risky call, check `Err.Number`, `Err.Clear`, `On Error Goto 0`, then act. Used for ADO/ADOX calls that can legitimately fail (e.g. `CreateSessionDatabase`, `ExecuteLocked` in `api/profiles.asp`) so a single write failure can't leave `Application.Lock` held forever.
- **Response envelope**: every API response is `{"data": ...}` or `{"error": {"code": "...", "message": "..."}}`, via `SendData`/`SendError` in `lib/http.asp`. Keep new endpoints consistent with this.
- **JSON is hand-rolled** (`lib/json.asp`) and **`JsonParse` only supports flat objects**: it raises on nested `{`/`[` in values, by design (matches the shape of the actual payloads this app sends: `firstName`/`lastName`/`email`/`bio`). Don't reach for `JsonParse` on the API's own envelope-wrapped responses (`{"data": {...}}`/`{"data": [...]}`): that's why `tests/api.tests.asp` uses `InStr`/`Extract*` string helpers instead.

## Security requirements for any new endpoint

- Gate every write (`POST`/`PUT`/`DELETE`) behind `RequireSession` + `RequireCsrf` (see `lib/db.asp`, `lib/csrf.asp`). Reads can skip both.
- Use parameterized `ADODB.Command` objects: never concatenate user input into SQL text.
- Validate required fields and max lengths server-side before hitting the DB (see `ValidateProfileFields` in `api/profiles.asp`) - don't rely on the DB column size or the frontend's `required`/`maxlength` attributes alone. Return a `400 validation_error` with a combined message, not a raw ADO error.
- Wrap writes in `Application.Lock`/`Unlock` (use the existing `ExecuteLocked` helper in `api/profiles.asp` rather than inlining a new Lock/Unlock pair) so a failed write can't hang the app pool.
- HTML-encode anything rendered into the DOM from user-supplied data: `assets/app.js` builds rows with `textContent`/DOM APIs rather than string-concatenated `innerHTML` for exactly this reason; keep doing that.

## Frontend logic split: app.core.js vs app.js

`assets/app.core.js` holds the frontend's pure(ish) logic (`extractErrorMessage`, `textCell`, `buildProfilePayload`, `createConfirmController`), exposed as `window.App`. `assets/app.js` does DOM lookups, axios calls, and event wiring, and calls into `App.*` rather than reimplementing that logic locally. This split exists so `assets/tests.html` can load `app.core.js` alone without needing the DOM elements/axios that `app.js` assumes exist. When adding new frontend logic that doesn't need `api`/live DOM event wiring to do its job, put it in `app.core.js`, not `app.js`, and add it to the `App` export.

## Testing

`tests/run.asp` is the whole story for the backend; see the README's Testing section. When adding a new `lib/*.asp` function, add a corresponding `Sub RunXTests()` in a matching `tests/*.tests.asp` file, `#include` it from `tests/run.asp`, and call it there. Prefer in-process assertions (the test page already has `Session`/`Server`/`Application` available) over a live HTTP round trip in `tests/api.tests.asp`. Reserve the HTTP round trip for things that genuinely depend on real request headers or verb routing.

For the frontend, `assets/tests.html` (+ `assets/testing.js`, `assets/tests.core.js`) is the equivalent for `app.core.js`, in the same hand-rolled assertion-framework style as the backend, with no npm/build step. It drives a real off-screen `<dialog>` in the test page itself for `createConfirmController` tests rather than mocking the DOM.

`a11y.html` (+ `assets/a11y.run.js`) runs axe-core against the live `index.html` markup, fetched at request time rather than duplicated in `a11y.html` itself - if `index.html` changes, this checker automatically reflects it, nothing to keep in sync manually. This is the one place this project pulls in an external test dependency (see the README's Testing section for why). When `index.html` gains new interactive elements, there's nothing to update here - just rerun the checker.

When hand-constructing JSON string literals inside a test (for tricky cases like embedded quotes/backslashes), build them with `Chr(34)`/`Chr(92)` rather than hand-doubled quote characters. See the comment at the top of `tests/json.tests.asp`. A doubled `""` is VBScript's own quoting convention, not JSON's `\"` escaping, and conflating the two produced a real (if short-lived) false test failure during development.
