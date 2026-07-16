<%
// Request/response helpers for the JSON API.
// Requires lib/json.asp to already be included (jsonEncode).

// Request.QueryString(key)/Request.ServerVariables(key) return a COM
// collection object even when key doesn't exist in the request - stringifying
// that empty collection produces the literal 9-character text "undefined",
// not an empty string or JScript's own undefined primitive (confirmed
// empirically - see CLAUDE.md). Checking .Count is the reliable way to tell
// "absent" from "present", regardless of language-level typeof/===undefined
// checks, which don't distinguish this case since the collection object
// itself is neither null nor undefined.
function requestParam(collectionItem) {
	if (collectionItem.Count === 0) return "";
	return String(collectionItem);
}

function getRequestMethod() {
	return requestParam(Request.ServerVariables("REQUEST_METHOD")).toUpperCase();
}

// Classic ASP has no native JSON body parsing - Request.Form only understands
// multipart/form-data and x-www-form-urlencoded. Read the raw bytes ourselves
// and decode as UTF-8 text so jsonParse can work on it. Same ADODB.Stream
// binary-to-text technique as asp-vbscript's lib/http.asp - BinaryRead's
// returned byte array is passed straight through to stream.Write untouched,
// so there's nothing JScript-specific to adapt here.
function readRawBody() {
	var totalBytes = Request.TotalBytes;
	if (totalBytes <= 0) return "";
	var bytes = Request.BinaryRead(totalBytes);
	var stream = Server.CreateObject("ADODB.Stream");
	stream.Type = 1; // adTypeBinary
	stream.Open();
	stream.Write(bytes);
	stream.Position = 0;
	stream.Type = 2; // adTypeText
	stream.Charset = "utf-8";
	var raw = stream.ReadText();
	stream.Close();
	return raw;
}

function statusText(code) {
	switch (code) {
		case 200: return "200 OK";
		case 201: return "201 Created";
		case 204: return "204 No Content";
		case 400: return "400 Bad Request";
		case 403: return "403 Forbidden";
		case 404: return "404 Not Found";
		case 405: return "405 Method Not Allowed";
		case 409: return "409 Conflict";
		case 500: return "500 Internal Server Error";
		default: return String(code) + " Error";
	}
}

// Writes { "data": <payload> } and ends the response.
function sendData(statusCode, payload) {
	sendJson(statusCode, { data: payload });
}

// Writes { "error": { "code": ..., "message": ... } } and ends the response.
function sendError(statusCode, errCode, message) {
	sendJson(statusCode, { error: { code: errCode, message: message } });
}

function sendJson(statusCode, payload) {
	Response.Status = statusText(statusCode);
	Response.ContentType = "application/json";
	Response.CharSet = "UTF-8";
	if (statusCode !== 204) {
		Response.Write(jsonEncode(payload));
	}
	// Confirmed (throwaway probe, Phase 0): Response.End() aborts the request
	// past any surrounding try/catch, uncaught - same as VBScript's
	// Response.End - so call sites never need an explicit return after
	// sendData/sendError/sendJson to prevent fallthrough.
	Response.End();
}
%>
