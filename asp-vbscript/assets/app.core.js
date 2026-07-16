// Pure(ish) logic pulled out of app.js so it's loadable from tests.html
// without pulling in the DOM-wiring/axios-calling half of the app.
// Keep this free of module-scoped DOM lookups (document.getElementById etc.)
// - everything here should work given only its own arguments.
(function (global) {
	"use strict";

	function extractErrorMessage(err) {
		if (err && err.response && err.response.data && err.response.data.error) {
			return err.response.data.error.message;
		}
		return (err && err.message) || "Something went wrong.";
	}

	function textCell(value, extraClass) {
		var td = document.createElement("td");
		if (extraClass) td.className = extraClass;
		td.textContent = value || "";
		return td;
	}

	function buildProfilePayload(raw) {
		raw = raw || {};
		return {
			firstName: (raw.firstName || "").trim(),
			lastName: (raw.lastName || "").trim(),
			email: (raw.email || "").trim(),
			bio: (raw.bio || "").trim()
		};
	}

	// Wraps a <dialog> + Cancel/Delete button pair into a Promise-based
	// confirm(), standing in for window.confirm. Event listeners are bound
	// once at construction; confirm() can be called repeatedly afterward.
	function createConfirmController(dialog, messageEl, cancelBtn, deleteBtn) {
		var pendingResolve = null;

		function resolvePending(result) {
			if (pendingResolve) {
				pendingResolve(result);
				pendingResolve = null;
			}
		}

		cancelBtn.addEventListener("click", function () {
			dialog.close();
			resolvePending(false);
		});
		deleteBtn.addEventListener("click", function () {
			dialog.close();
			resolvePending(true);
		});
		// Esc key fires "cancel" (and then closes the dialog itself) - treat
		// it the same as clicking Cancel.
		dialog.addEventListener("cancel", function () {
			resolvePending(false);
		});

		return {
			confirm: function (message) {
				messageEl.textContent = message;
				dialog.showModal();
				return new Promise(function (resolve) {
					pendingResolve = resolve;
				});
			}
		};
	}

	global.App = {
		extractErrorMessage: extractErrorMessage,
		textCell: textCell,
		buildProfilePayload: buildProfilePayload,
		createConfirmController: createConfirmController
	};
})(window);
