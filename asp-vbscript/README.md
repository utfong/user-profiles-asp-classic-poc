# User Profiles CRUD POC: ASP Classic (VBScript) + Access

A proof of concept demonstrating a "modern" CRUD API pattern (real HTTP verbs, JSON, CSRF-protected writes) built entirely on Classic ASP/VBScript, with a plain HTML/CSS + [axios](https://axios-http.com/) frontend.

**The twist:** the app ships with no database file. A *Create Session* button provisions a fresh, session-scoped Microsoft Access (`.accdb`) file; it's deleted automatically when the session ends (explicitly, or on idle timeout). Nothing persists between sessions by design, so it's safe to demo repeatedly.

## Prerequisites

- IIS with the Classic ASP feature enabled, and this folder set up as its own virtual host / site.
- **Microsoft Access Database Engine 2016 Redistributable** ([ACE OLEDB 12.0](https://www.microsoft.com/en-us/download/details.aspx?id=54920)), installed with a bitness that **matches the app pool** (this project assumes 64-bit/64-bit; see [CLAUDE.md](CLAUDE.md) if yours differs).
- The site's `.asp` handler mapping must allow `PUT` and `DELETE`, not just the IIS default `GET,HEAD,POST,DEBUG`. In IIS Manager: *Handler Mappings* → the ASP mapping (usually `ASPClassic`) → *Edit Feature Permissions*/*Request Restrictions* → *Verbs* tab → allow all verbs (or add `PUT,DELETE`).
- The app pool identity needs write access to this folder (it creates `App_Data\sessions\` on first use).

## Project structure

```
asp-vbscript/
  global.asa            # Session_OnEnd -> deletes the session's .accdb (idle-timeout cleanup)
  web.config             # default document + blocks direct HTTP access to .accdb/.mdb
  index.html              # frontend: session banner, profile table, create/edit dialog, delete confirmation dialog
  a11y.html                # accessibility checker - loads index.html + axe-core, see Testing section
  assets/
    app.css
    app.core.js              # pure(ish) logic: error extraction, payload building, confirm dialog controller
    app.js                    # DOM wiring: axios calls, CSRF header wiring, rendering (uses app.core.js)
    testing.js                 # browser assertion framework, JS equivalent of lib/testing.asp
    tests.core.js                # test cases for app.core.js
    tests.html                    # test runner - open this file to run the JS suite
    a11y.run.js                     # axe-core runner + report rendering for a11y.html
  api/
    session.asp            # session lifecycle: GET status / POST create / DELETE end
    profiles.asp            # profiles CRUD: GET / POST / PUT / DELETE
  lib/
    json.asp                 # hand-rolled JSON encode/parse (classic ASP has no native JSON)
    http.asp                  # request/response helpers (raw body reading, JSON responses)
    csrf.asp                   # CSRF token minting + validation
    db.asp                      # per-session .accdb lifecycle + ADO connection helpers
    testing.asp                  # assertion framework used by tests/
  tests/
    run.asp                       # test runner - hit this URL to run everything
    *.tests.asp                    # test suites (json, csrf, db, api)
  App_Data/sessions/                # created at runtime, holds {SessionID}.accdb per active session
```

## Running it

1. Browse to the site (e.g. `http://asp-vbscript.local`).
2. The profile list loads immediately: it's empty, since no session exists yet.
3. Click **Create Session** to provision a session-scoped `.accdb`.
4. Create/edit/delete profiles as usual.
5. Click **End Session** (or leave it idle for ~5 minutes) to tear the database file down.

## API reference

All responses are JSON: `{"data": ...}` on success, `{"error": {"code": "...", "message": "..."}}` on failure.

| Method | Path | Auth | Description |
|---|---|---|---|
| `GET` | `/api/session.asp` | none | Current session status (`active`, `csrfToken`, `timeoutMinutes`) |
| `POST` | `/api/session.asp` | none | Create a session + `.accdb` (idempotent if already active) |
| `DELETE` | `/api/session.asp` | session + CSRF | End the session, delete the `.accdb` immediately |
| `GET` | `/api/profiles.asp` | none | List all profiles (empty array if no session) |
| `GET` | `/api/profiles.asp?id=N` | none | Retrieve one profile (`404` if no session or not found) |
| `POST` | `/api/profiles.asp` | session + CSRF | Create a profile (body: `firstName`, `lastName`, `email`, `bio`) |
| `PUT` | `/api/profiles.asp?id=N` | session + CSRF | Update a profile (same body shape) |
| `DELETE` | `/api/profiles.asp?id=N` | session + CSRF | Delete a profile |

CSRF token is sent via the `X-CSRF-Token` header, obtained from `POST /api/session.asp`'s response.

### Validation

Enforced server-side in `api/profiles.asp` (`ValidateProfileFields`) on both create and update, in addition to the DB column sizes in `lib/db.asp`. A `400 validation_error` is returned with a combined message when one or more rules fail.

| Field | Required | Max length |
|---|---|---|
| `firstName` | Yes | 100 |
| `lastName` | Yes | 100 |
| `email` | Yes | 255 |
| `bio` | No | none (MEMO column) |

## Security model

- **CSRF**: synchronizer token minted into `Session("CsrfToken")` on session creation, required via `X-CSRF-Token` on every write. Missing/mismatched → `403`.
- **Reads are session-optional, writes are not**: `GET` works with or without an active session (there's just nothing to read without one); `POST`/`PUT`/`DELETE` require both an active session and a valid CSRF token.
- **SQL injection**: every query uses parameterized `ADODB.Command` objects instead of string-concatenated SQL.
- **No direct file access**: `.accdb`/`.mdb` extensions are blocked at the IIS level (`web.config` request filtering), so the database can't be downloaded directly even during an active session.
- **Session isolation**: each session gets its own file under `App_Data/sessions/{SessionID}.accdb`, so there's no shared state between visitors.
- **Write serialization**: writes run inside `Application.Lock`/`Unlock` (see `ExecuteLocked` in `api/profiles.asp`) as insurance against Access/Jet file-locking flakiness under concurrent requests; guaranteed to unlock even if the write fails.

## Testing

Visit `/tests/run.asp` (browser or any HTTP client) to run the full suite. It's localhost-only, since it creates/destroys real data and makes live calls against the API. Covers:

- `lib/json.asp` encode/parse logic (in-process, no HTTP)
- `lib/csrf.asp` token generation (in-process)
- `lib/db.asp` against a throwaway `.accdb`: schema, parameterized queries, cleanup (in-process)
- Full session + CRUD + CSRF-rejection flow over real HTTP (via `MSXML2.ServerXMLHTTP`, self-calling the live API)

The response includes `X-Test-Status: pass|fail` and pass/fail count headers for scripted checks, e.g.:

```powershell
Invoke-WebRequest -Uri "http://asp-vbscript.local/tests/run.asp" -UseBasicParsing
```

### Frontend (JS) tests

`assets/tests.html` covers the pure(ish) logic in `assets/app.core.js` (error message extraction, profile payload trimming, and the confirm-dialog controller) using a small hand-rolled assertion framework (`assets/testing.js`), in the same style as `lib/testing.asp` and with no npm/build step. Open the file directly in a browser, or serve it like any other static asset:

```
http://asp-vbscript.local/assets/tests.html
```

Pass/fail is rendered on the page and also set as `data-test-status` on `<body>` for scripted checks.

### Accessibility check

`a11y.html` scans the live `index.html` with [axe-core](https://github.com/dequelabs/axe-core) (the one exception to this project's "no external test dependencies" rule, since hand-rolling contrast-ratio math and WCAG rule checks isn't realistic). It fetches `index.html`'s current markup at request time (rather than duplicating it, which would drift out of sync), so it always reflects whatever's actually in `index.html`. Must be served over HTTP (the `fetch()` call needs it). Opening the file directly (`file://`) won't work:

```
http://asp-vbscript.local/a11y.html
```

Results render on the page and are also set as `data-a11y-status` on `<body>` (`pass`/`fail`) for scripted checks.

## Known limitations (it's a POC)

- No authentication/login: anyone who can reach the site can create a session and write data to it.
- `email` validation is presence and length only, with no format check (e.g. requiring an `@`). The frontend's `<input type="email">` gives light browser-side format hinting, but the API doesn't enforce it.
- CSRF token is generated with VBScript's `Rnd()`, not a CSPRNG (see `lib/csrf.asp`).
- No URL Rewrite: resource IDs are passed as `?id=` query strings rather than path segments.
- `Application.Lock` serializes writes **across the whole app**, not per-session. That's fine for a low-traffic demo, but unnecessarily strict for many concurrent unrelated sessions at real scale.
- Accessibility has been checked with `a11y.html` (axe-core) plus manual review, but not with a real screen reader (NVDA/JAWS/VoiceOver) or a manual keyboard-only walkthrough. Native `<dialog>` should handle focus trapping and return focus to the trigger element on close, but that assumption hasn't been verified by hand.
