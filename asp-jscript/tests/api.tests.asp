<%
// Live HTTP round-trip tests: the API is self-called over real HTTP, because
// verb routing and header-based CSRF checking can only be observed through an
// actual request, not in-process. Unlike asp-vbscript's version of this file
// (which uses MSXML2.ServerXMLHTTP.6.0, same as here), response bodies here
// are parsed with jsonParse() directly (including the {"data": ...}/
// {"error": {...}} envelope) instead of InStr/Extract* string-scraping - our
// parser supports nesting, so there's no need to work around a flat-objects-
// only limitation.
//
// Uses WinHttp.WinHttpRequest.5.1 instead of MSXML2.ServerXMLHTTP.6.0: the
// latter's Open() method isn't resolvable through JScript's IDispatch calls
// on this machine (confirmed via a throwaway probe - "Open" in obj is false,
// and calling it throws "Object doesn't support property or method 'Open'"),
// even though the identical component works fine from VBScript. WinHttpRequest
// exposes the same Open/setRequestHeader/send/status/responseText/
// getResponseHeader surface and works correctly from JScript here.

function getSelfBaseUrl() {
	var scheme = (String(Request.ServerVariables("HTTPS")) === "on") ? "https" : "http";
	return scheme + "://" + String(Request.ServerVariables("HTTP_HOST"));
}

function extractCookiePair(setCookieHeader) {
	if (setCookieHeader === "") return "";
	var firstPart = setCookieHeader.split(";")[0];
	return firstPart.replace(/^\s+|\s+$/g, "");
}

// Classic ASP JScript has no native sleep - busy-wait against Date.getTime()
// is the standard workaround (same technique as tests/db.tests.asp).
function briefDelay(ms) {
	var start = (new Date()).getTime();
	while ((new Date()).getTime() - start < ms) { /* busy wait */ }
}

// A real error response from this app is always JSON starting with "{"
// (see lib/http.asp's sendError envelope). A 500 with anything else - an
// HTML error page - is this host's known intermittent script-engine crash
// (see CLAUDE.md), not a genuine application error.
function looksLikeEngineCrash(http) {
	if (http.status !== 500) return false;
	var body = http.responseText;
	return body.length === 0 || body.charAt(0) !== "{";
}

// Returns null (and records a failure) instead of throwing, so one flaky
// request can't take down the whole test run. Retries automatically on the
// known intermittent engine crash (see looksLikeEngineCrash) - it's
// probabilistic per-request, so a few attempts make any single call very
// likely to eventually succeed, which in turn makes the whole chained test
// flow far more likely to complete cleanly in one run.
function httpCall(method, url, cookieHeader, csrfToken, jsonBody) {
	var maxAttempts = 4;
	for (var attempt = 1; attempt <= maxAttempts; attempt++) {
		var http = null;
		try {
			http = Server.CreateObject("WinHttp.WinHttpRequest.5.1");
			http.Open(method, url, false);
			if (jsonBody !== "") http.setRequestHeader("Content-Type", "application/json");
			if (cookieHeader !== "") http.setRequestHeader("Cookie", cookieHeader);
			if (csrfToken !== "") http.setRequestHeader("X-CSRF-Token", csrfToken);
			if (jsonBody !== "") {
				http.send(jsonBody);
			} else {
				http.send();
			}
		} catch (e) {
			recordResult(false, "HTTP " + method + " " + url + " failed to complete: " + e.message);
			return null;
		}
		if (!looksLikeEngineCrash(http)) return http;
		if (attempt === maxAttempts) {
			recordResult(false, "HTTP " + method + " " + url + " kept hitting the known engine crash after " + maxAttempts + " attempts");
			return http; // let the caller's own assertions report the specific mismatch
		}
		briefDelay(300);
	}
}

// httpCall() only returns null on a network-level failure. A server-side
// crash (this host's known intermittent script-engine crash, see CLAUDE.md)
// still produces a real HTTP response - status 500, HTML error page body -
// so jsonParse(http.responseText) would throw on the "<" of that HTML,
// crashing the whole test run instead of just failing one assertion. Route
// every response body through this instead of calling jsonParse directly.
function safeJsonParse(http) {
	try {
		return jsonParse(http.responseText);
	} catch (e) {
		recordResult(false, "Response body was not valid JSON (status " + http.status + "): " + e.message);
		return null;
	}
}

function runApiTests() {
	var baseUrl = getSelfBaseUrl();

	beginTest("api", "GET /api/profiles.asp works with no session at all");
	var http = httpCall("GET", baseUrl + "/api/profiles.asp", "", "", "");
	if (http === null) return;
	assertEqual(200, http.status, "status");
	var listBody = safeJsonParse(http);
	if (listBody !== null) {
		assertEqual(0, listBody.data.length, "empty list when no session");
	}

	beginTest("api", "POST /api/profiles.asp without a session is rejected");
	http = httpCall("POST", baseUrl + "/api/profiles.asp", "", "", jsonEncode({ firstName: "X", lastName: "Y" }));
	if (http === null) return;
	assertEqual(400, http.status, "status");
	var noSessionBody = safeJsonParse(http);
	if (noSessionBody !== null) {
		assertEqual("no_session", noSessionBody.error.code, "error code");
	}

	beginTest("api", "POST /api/session.asp creates a session");
	http = httpCall("POST", baseUrl + "/api/session.asp", "", "", "");
	if (http === null) return;
	assertEqual(201, http.status, "status");
	var cookieHeader = extractCookiePair(http.getResponseHeader("Set-Cookie"));
	var sessionBody = safeJsonParse(http);
	// Everything below needs a real session + CSRF token - if this step
	// itself hit the crash (or returned something unparseable), there's
	// nothing meaningful left to test.
	if (sessionBody === null) return;
	var csrf = sessionBody.data.csrfToken;
	assertTrue(cookieHeader !== "", "session cookie should be set");
	assertEqual(32, csrf.length, "csrf token length");

	beginTest("api", "POST /api/profiles.asp without email is rejected");
	http = httpCall("POST", baseUrl + "/api/profiles.asp", cookieHeader, csrf, jsonEncode({ firstName: "Test", lastName: "User" }));
	if (http !== null) {
		assertEqual(400, http.status, "status");
		var validationBody = safeJsonParse(http);
		if (validationBody !== null) {
			assertEqual("validation_error", validationBody.error.code, "error code");
			assertTrue(validationBody.error.message.indexOf("email is required") !== -1, "message mentions missing email");
		}
	}

	beginTest("api", "POST /api/profiles.asp with an oversized firstName is rejected");
	// 101 chars - one over api/profiles.asp's FIRSTNAME_MAX_LEN.
	var longName = "";
	for (var i = 0; i < 101; i++) longName += "A";
	http = httpCall("POST", baseUrl + "/api/profiles.asp", cookieHeader, csrf,
		jsonEncode({ firstName: longName, lastName: "User", email: "t@example.com" }));
	if (http !== null) {
		assertEqual(400, http.status, "status");
		var longNameBody = safeJsonParse(http);
		if (longNameBody !== null) {
			assertTrue(longNameBody.error.message.indexOf("firstName must be 100 characters or fewer") !== -1, "message mentions max length");
		}
	}

	beginTest("api", "POST /api/profiles.asp creates a profile");
	http = httpCall("POST", baseUrl + "/api/profiles.asp", cookieHeader, csrf,
		jsonEncode({ firstName: "Test", lastName: "User", email: "t@example.com", bio: "hi" }));
	if (http === null) {
		httpCall("DELETE", baseUrl + "/api/session.asp", cookieHeader, csrf, "");
		return;
	}
	assertEqual(201, http.status, "status");
	var createdBody = safeJsonParse(http);
	// The rest of the CRUD flow needs a real id - bail out (after cleaning up
	// the session) if creation didn't give us one.
	if (createdBody === null) {
		httpCall("DELETE", baseUrl + "/api/session.asp", cookieHeader, csrf, "");
		return;
	}
	var newId = createdBody.data.id;
	assertTrue(newId !== undefined && newId !== null, "should return a new id");

	beginTest("api", "PUT without a CSRF header is rejected and does not modify data");
	http = httpCall("PUT", baseUrl + "/api/profiles.asp?id=" + newId, cookieHeader, "",
		jsonEncode({ firstName: "TAMPERED", lastName: "User", email: "t@example.com", bio: "hi" }));
	if (http !== null) {
		assertEqual(403, http.status, "status");
	}

	beginTest("api", "GET after the failed tamper still shows the original data");
	http = httpCall("GET", baseUrl + "/api/profiles.asp?id=" + newId, cookieHeader, "", "");
	if (http !== null) {
		var afterTamperBody = safeJsonParse(http);
		if (afterTamperBody !== null) {
			assertEqual("Test", afterTamperBody.data.firstName, "original value should remain");
		}
	}

	beginTest("api", "PUT with a valid CSRF header succeeds");
	http = httpCall("PUT", baseUrl + "/api/profiles.asp?id=" + newId, cookieHeader, csrf,
		jsonEncode({ firstName: "Updated", lastName: "User", email: "t@example.com", bio: "hi" }));
	if (http !== null) {
		assertEqual(200, http.status, "status");
		var updatedBody = safeJsonParse(http);
		if (updatedBody !== null) {
			assertEqual("Updated", updatedBody.data.firstName, "update applied");
		}
	}

	beginTest("api", "DELETE /api/profiles.asp removes the profile");
	http = httpCall("DELETE", baseUrl + "/api/profiles.asp?id=" + newId, cookieHeader, csrf, "");
	if (http !== null) {
		assertEqual(204, http.status, "status");
	}

	beginTest("api", "GET after delete returns 404");
	http = httpCall("GET", baseUrl + "/api/profiles.asp?id=" + newId, cookieHeader, "", "");
	if (http !== null) {
		assertEqual(404, http.status, "status");
	}

	beginTest("api", "DELETE /api/session.asp ends the session");
	http = httpCall("DELETE", baseUrl + "/api/session.asp", cookieHeader, csrf, "");
	if (http !== null) {
		assertEqual(200, http.status, "status");
		var endBody = safeJsonParse(http);
		if (endBody !== null) {
			assertEqual(false, endBody.data.active, "session reports inactive");
		}
	}

	beginTest("api", "session's accdb file is removed from disk after ending");
	var sessionIdFromCookie = cookieHeader.substring(cookieHeader.indexOf("=") + 1);
	var accdbPath = getSessionsFolder() + "\\" + sessionIdFromCookie + ".accdb";
	var fso2 = Server.CreateObject("Scripting.FileSystemObject");
	assertFalse(fso2.FileExists(accdbPath), "session accdb should be deleted from disk");
}
%>
