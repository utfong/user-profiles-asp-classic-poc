// Runner for a11y.html - scans the (by this point fully wired-up) page
// with axe-core and renders the result into #a11yReport. Loaded as the
// very last script, after index.html's own axios/app.core.js/app.js.
(function () {
	"use strict";

	function formatViolation(v) {
		var nodeList = v.nodes.map(function (n) { return "  " + n.target.join(" "); }).join("\n");
		return "[" + v.impact + "] " + v.id + " - " + v.help + "\n" + nodeList;
	}

	function render(results) {
		var el = document.getElementById("a11yReport");
		var violations = results.violations;

		if (violations.length === 0) {
			el.textContent = "0 accessibility violations found (" + results.passes.length + " checks passed).";
			el.className = "pass";
		} else {
			el.textContent = violations.length + " accessibility violation(s):\n\n" +
				violations.map(formatViolation).join("\n\n");
			el.className = "fail";
		}

		document.title = (violations.length === 0 ? "PASS" : "FAIL") + " - Accessibility Check";
		document.body.setAttribute("data-a11y-status", violations.length === 0 ? "pass" : "fail");
	}

	// Give the app a moment to finish its initial render (session status +
	// profile list fetch) before scanning, same reasoning as the delay in
	// assets/tests.core.js's confirm-controller tests.
	setTimeout(function () {
		// Exclude the report panel itself - it's this checker's own UI, not
		// part of the app under test, and would otherwise always show up as
		// a false "content not in a landmark" violation.
		axe.run({ exclude: [["#a11yReport"]] }, {}, function (err, results) {
			if (err) {
				document.getElementById("a11yReport").textContent = "axe.run error: " + err.message;
				document.getElementById("a11yReport").className = "fail";
				return;
			}
			render(results);
		});
	}, 1500);
})();
