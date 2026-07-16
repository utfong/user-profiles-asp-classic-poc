// Minimal xUnit-style assertion/reporting framework for assets/tests.html.
// Browser equivalent of lib/testing.asp - same BeginTest/Assert*/report
// shape, adapted to render into the DOM instead of Response.Write.
(function (global) {
	"use strict";

	var results = [];
	var currentSuite = "";
	var currentTest = "";

	function beginTest(suite, test) {
		currentSuite = suite;
		currentTest = test;
	}

	function recordResult(passed, detail) {
		results.push({ suite: currentSuite, test: currentTest, passed: passed, detail: detail });
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

	function renderReport(containerEl) {
		var passCount = 0;
		var failCount = 0;
		results.forEach(function (r) {
			if (r.passed) {
				passCount = passCount + 1;
			} else {
				failCount = failCount + 1;
			}
		});

		containerEl.textContent = "";

		var pre = document.createElement("pre");
		results.forEach(function (r) {
			var line = document.createElement("div");
			line.className = r.passed ? "pass" : "fail";
			line.textContent = (r.passed ? "PASS  " : "FAIL  ") +
				"[" + r.suite + "] " + r.test + " -- " + r.detail;
			pre.appendChild(line);
		});
		containerEl.appendChild(pre);

		var summary = document.createElement("div");
		summary.className = "summary " + (failCount === 0 ? "pass" : "fail");
		summary.textContent = passCount + " passed, " + failCount +
			" failed (" + results.length + " total)";
		containerEl.appendChild(summary);

		document.title = (failCount === 0 ? "PASS" : "FAIL") + " - JS Test Results";
		document.body.setAttribute("data-test-status", failCount === 0 ? "pass" : "fail");
	}

	global.Testing = {
		beginTest: beginTest,
		assertEqual: assertEqual,
		assertTrue: assertTrue,
		assertFalse: assertFalse,
		renderReport: renderReport
	};
})(window);
