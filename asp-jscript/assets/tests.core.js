// Test cases for assets/app.core.js. Loaded by assets/tests.html after
// testing.js and app.core.js. Runs on load and renders into #report.
(function () {
	"use strict";

	function runExtractErrorMessageTests() {
		Testing.beginTest("core", "extractErrorMessage returns the API error message when present");
		var apiErr = { response: { data: { error: { message: "boom" } } } };
		Testing.assertEqual("boom", App.extractErrorMessage(apiErr), "message");

		Testing.beginTest("core", "extractErrorMessage falls back to err.message");
		Testing.assertEqual("network fail", App.extractErrorMessage({ message: "network fail" }), "message");

		Testing.beginTest("core", "extractErrorMessage falls back to a generic message");
		Testing.assertEqual("Something went wrong.", App.extractErrorMessage({}), "message");
	}

	function runTextCellTests() {
		Testing.beginTest("core", "textCell renders the given value as a <td>");
		var td = App.textCell("Ada");
		Testing.assertEqual("TD", td.tagName, "tag name");
		Testing.assertEqual("Ada", td.textContent, "text content");

		Testing.beginTest("core", "textCell falls back to an empty string for null/undefined");
		Testing.assertEqual("", App.textCell(null).textContent, "text content for null");
		Testing.assertEqual("", App.textCell(undefined).textContent, "text content for undefined");

		Testing.beginTest("core", "textCell applies the extra class when given");
		Testing.assertEqual("bio", App.textCell("hi", "bio").className, "class name");
	}

	function runBuildProfilePayloadTests() {
		Testing.beginTest("core", "buildProfilePayload trims whitespace on all fields");
		var payload = App.buildProfilePayload({
			firstName: "  Ada  ",
			lastName: " Lovelace ",
			email: " ada@example.com ",
			bio: " hi "
		});
		Testing.assertEqual("Ada", payload.firstName, "firstName");
		Testing.assertEqual("Lovelace", payload.lastName, "lastName");
		Testing.assertEqual("ada@example.com", payload.email, "email");
		Testing.assertEqual("hi", payload.bio, "bio");

		Testing.beginTest("core", "buildProfilePayload treats missing fields as empty strings");
		var empty = App.buildProfilePayload({});
		Testing.assertEqual("", empty.firstName, "firstName");
		Testing.assertEqual("", empty.lastName, "lastName");
		Testing.assertEqual("", empty.email, "email");
		Testing.assertEqual("", empty.bio, "bio");
	}

	// Drives the actual <dialog> defined in tests.html, so this exercises
	// real showModal()/close() browser behavior rather than a mock.
	function runConfirmControllerTests() {
		var dialog = document.getElementById("testConfirmDialog");
		var messageEl = document.getElementById("testConfirmMessage");
		var cancelBtn = document.getElementById("testConfirmCancel");
		var deleteBtn = document.getElementById("testConfirmDelete");
		var controller = App.createConfirmController(dialog, messageEl, cancelBtn, deleteBtn);

		return Promise.resolve()
			.then(function () {
				var pending = controller.confirm("Delete this?");

				Testing.beginTest("core", "confirm() opens the dialog with the given message");
				Testing.assertTrue(dialog.open, "dialog.open");
				Testing.assertEqual("Delete this?", messageEl.textContent, "message text");

				deleteBtn.click();
				return pending;
			})
			.then(function (result) {
				Testing.beginTest("core", "confirm() resolves true and closes the dialog when Delete is clicked");
				Testing.assertTrue(result, "result");
				Testing.assertFalse(dialog.open, "dialog.open after close");

				var pending = controller.confirm("Delete this?");
				cancelBtn.click();
				return pending;
			})
			.then(function (result) {
				Testing.beginTest("core", "confirm() resolves false when Cancel is clicked");
				Testing.assertFalse(result, "result");

				var pending = controller.confirm("Delete this?");
				dialog.dispatchEvent(new Event("cancel"));
				dialog.close();
				return pending;
			})
			.then(function (result) {
				Testing.beginTest("core", "confirm() resolves false on the dialog's native cancel event (Esc key)");
				Testing.assertFalse(result, "result");
			});
	}

	runExtractErrorMessageTests();
	runTextCellTests();
	runBuildProfilePayloadTests();
	runConfirmControllerTests().then(function () {
		Testing.renderReport(document.getElementById("report"));
	});
})();
