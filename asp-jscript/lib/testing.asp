<%
// Minimal xUnit-style assertion/reporting framework for the tests/ suite.
// Not reentrant - designed for a single test-runner page (tests/run.asp)
// that calls each runXTests() function once per request. Mirrors
// asp-vbscript's lib/testing.asp; testResults is a plain array of object
// literals pushed with .push(...) rather than ReDim Preserve against a
// Dictionary-per-row - genuinely less code here, not just different syntax.

var testResults = [];
var currentSuiteName = "";
var currentTestName = "";

function resetTests() {
	testResults = [];
	currentSuiteName = "";
	currentTestName = "";
}

function beginTest(suiteName, testName) {
	currentSuiteName = suiteName;
	currentTestName = testName;
}

function recordResult(passed, detail) {
	testResults.push({ suite: currentSuiteName, test: currentTestName, passed: passed, detail: detail });
}

function assertEqual(expected, actual, message) {
	if (String(expected) === String(actual)) {
		recordResult(true, message);
	} else {
		recordResult(false, message + " -- expected [" + expected + "] but got [" + actual + "]");
	}
}

function assertTrue(condition, message) {
	if (condition) {
		recordResult(true, message);
	} else {
		recordResult(false, message + " -- expected true but got false");
	}
}

function assertFalse(condition, message) {
	assertTrue(!condition, message);
}

function renderTestReport() {
	var passCount = 0;
	var failCount = 0;
	for (var i = 0; i < testResults.length; i++) {
		if (testResults[i].passed) {
			passCount++;
		} else {
			failCount++;
		}
	}

	var testStatus = failCount === 0 ? "pass" : "fail";

	// Headers first, before any Write, regardless of buffering settings.
	Response.ContentType = "text/html";
	Response.CharSet = "UTF-8";
	Response.AddHeader("X-Test-Status", testStatus);
	Response.AddHeader("X-Test-Pass-Count", String(passCount));
	Response.AddHeader("X-Test-Fail-Count", String(failCount));

	Response.Write('<!DOCTYPE html><html><head><meta charset="utf-8"><title>ASP JScript Test Results</title>');
	Response.Write("<style>body{font-family:Consolas,Menlo,monospace;background:#0b0f19;color:#e4e7ec;padding:24px}");
	Response.Write(".pass{color:#4ade80}.fail{color:#f87171}h1{font-size:1.15rem}");
	Response.Write(".summary{font-weight:bold;margin-top:16px;font-size:1.1rem;padding:10px 14px;border-radius:6px;display:inline-block}");
	Response.Write(".summary.pass{background:#0f2a1c}.summary.fail{background:#3a1414}</style></head><body>");
	Response.Write("<h1>ASP JScript Test Results</h1><pre>");

	for (var j = 0; j < testResults.length; j++) {
		var r = testResults[j];
		if (r.passed) {
			Response.Write('<span class="pass">PASS</span>  ');
		} else {
			Response.Write('<span class="fail">FAIL</span>  ');
		}
		Response.Write("[" + Server.HTMLEncode(r.suite) + "] " + Server.HTMLEncode(r.test) +
			" -- " + Server.HTMLEncode(r.detail) + "\r\n");
	}

	Response.Write("</pre>");
	Response.Write('<div class="summary ' + testStatus + '">' + passCount + " passed, " + failCount +
		" failed (" + testResults.length + " total)</div>");
	Response.Write("</body></html>");
}
%>
