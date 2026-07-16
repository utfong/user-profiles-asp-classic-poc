<%@ Language="JScript" %>
<%
// Test runner: hit this page directly (browser or HTTP client) to run the
// whole suite. Localhost-only, since these tests create/destroy real .accdb
// files and make live calls against the API.
%>
<!--#include virtual="/lib/testing.asp" -->
<!--#include virtual="/lib/json.asp" -->
<!--#include virtual="/lib/csrf.asp" -->
<!--#include virtual="/lib/db.asp" -->
<!--#include virtual="/tests/json.tests.asp" -->
<!--#include virtual="/tests/csrf.tests.asp" -->
<!--#include virtual="/tests/db.tests.asp" -->
<!--#include virtual="/tests/api.tests.asp" -->
<%
var remoteAddr = String(Request.ServerVariables("REMOTE_ADDR"));
if (remoteAddr !== "127.0.0.1" && remoteAddr !== "::1") {
	Response.Status = "403 Forbidden";
	Response.ContentType = "text/plain";
	Response.Write("Test runner is only reachable from localhost.");
	Response.End();
}

resetTests();

// Tests borrow the live Session object (it's the only way to get one outside
// a real ASP request) - save/restore it so a real visitor's active session
// isn't disturbed by running the suite.
var savedDbPath = Session("DbPath");
var savedCsrfToken = Session("CsrfToken");

runJsonTests();
runCsrfTests();
runDbTests();
runApiTests();

Session("DbPath") = savedDbPath;
Session("CsrfToken") = savedCsrfToken;

renderTestReport();
%>
