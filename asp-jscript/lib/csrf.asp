<%
// CSRF synchronizer-token helpers.
// Note: token is generated with Math.random(), not a CSPRNG - same caveat as
// asp-vbscript's Rnd()-based version. Classic ASP/JScript has no built-in
// crypto RNG without an extra COM component. Fine for a POC; a real
// deployment should source this from something like CAPICOM or a native
// helper DLL instead.

function newCsrfToken() {
	var token = "";
	for (var i = 0; i < 32; i++) {
		token += Math.floor(Math.random() * 16).toString(16);
	}
	return token;
}

function isCsrfValid() {
	// Request.ServerVariables returns a collection object even for a header
	// the client didn't send - use requestParam (lib/http.asp), not a plain
	// ===undefined check, or a missing header reads as the literal text
	// "undefined" instead of "".
	var headerToken = requestParam(Request.ServerVariables("HTTP_X_CSRF_TOKEN"));
	var sessionToken = Session("CsrfToken");
	sessionToken = (sessionToken === null || sessionToken === undefined) ? "" : String(sessionToken);
	return sessionToken !== "" && headerToken !== "" && headerToken === sessionToken;
}

// Guard for write endpoints: ends the response with 403 if the token is
// missing or doesn't match. Requires lib/http.asp (sendError) to be included.
function requireCsrf() {
	if (!isCsrfValid()) {
		sendError(403, "invalid_csrf", "Missing or invalid CSRF token.");
	}
}
%>
