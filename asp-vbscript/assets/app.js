(function () {
	"use strict";

	var api = axios.create({ baseURL: "api/" });

	var sessionBanner = document.getElementById("sessionBanner");
	var sessionBannerText = document.getElementById("sessionBannerText");
	var btnCreateSession = document.getElementById("btnCreateSession");
	var btnEndSession = document.getElementById("btnEndSession");
	var errorBanner = document.getElementById("errorBanner");
	var profilesSection = document.getElementById("profilesSection");
	var profilesBody = document.getElementById("profilesBody");
	var btnNewProfile = document.getElementById("btnNewProfile");

	var profileDialog = document.getElementById("profileDialog");
	var profileForm = document.getElementById("profileForm");
	var profileDialogTitle = document.getElementById("profileDialogTitle");
	var btnCancelProfile = document.getElementById("btnCancelProfile");
	var fieldId = document.getElementById("profileId");
	var fieldFirstName = document.getElementById("firstName");
	var fieldLastName = document.getElementById("lastName");
	var fieldEmail = document.getElementById("email");
	var fieldBio = document.getElementById("bio");

	var confirmDialog = document.getElementById("confirmDialog");
	var confirmDialogMessage = document.getElementById("confirmDialogMessage");
	var btnConfirmCancel = document.getElementById("btnConfirmCancel");
	var btnConfirmDelete = document.getElementById("btnConfirmDelete");

	// Shared logic lives in app.core.js (window.App) so it's loadable from
	// assets/tests.html without the DOM-wiring/axios-calling code below.
	var extractErrorMessage = App.extractErrorMessage;
	var textCell = App.textCell;
	var confirmController = App.createConfirmController(confirmDialog, confirmDialogMessage, btnConfirmCancel, btnConfirmDelete);

	function showError(message) {
		errorBanner.textContent = message;
		errorBanner.classList.remove("hidden");
	}

	function clearError() {
		errorBanner.textContent = "";
		errorBanner.classList.add("hidden");
	}

	function setCsrfToken(token) {
		if (token) {
			api.defaults.headers.common["X-CSRF-Token"] = token;
		} else {
			delete api.defaults.headers.common["X-CSRF-Token"];
		}
	}

	// Reads are always allowed (there's just nothing to read without a
	// session, since profiles only exist inside a session's own Access
	// file). Only creating/editing/deleting requires an active session.
	function setSessionUi(active, timeoutMinutes) {
		if (active) {
			sessionBanner.classList.remove("banner-inactive");
			sessionBanner.classList.add("banner-active");
			sessionBannerText.textContent =
				"Session active. Data is stored in a temporary Access file " +
				"(idle timeout: " + timeoutMinutes + " min).";
			btnCreateSession.classList.add("hidden");
			btnEndSession.classList.remove("hidden");
			btnNewProfile.disabled = false;
		} else {
			sessionBanner.classList.add("banner-inactive");
			sessionBanner.classList.remove("banner-active");
			sessionBannerText.textContent =
				"Click \"Create Session\" to start adding, editing, or deleting profiles. " +
				"Session data is temporary and clears automatically when the session ends.";
			btnCreateSession.classList.remove("hidden");
			btnEndSession.classList.add("hidden");
			btnNewProfile.disabled = true;
		}
	}

	function renderProfiles(profiles) {
		profilesBody.replaceChildren();
		if (!profiles.length) {
			var tr = document.createElement("tr");
			var td = document.createElement("td");
			td.colSpan = 5;
			td.className = "empty";
			td.textContent = "No profiles yet.";
			tr.appendChild(td);
			profilesBody.appendChild(tr);
			return;
		}
		profiles.forEach(function (p) {
			var row = document.createElement("tr");
			row.appendChild(textCell(p.firstName));
			row.appendChild(textCell(p.lastName));
			row.appendChild(textCell(p.email));
			row.appendChild(textCell(p.bio, "bio"));

			var actionsCell = document.createElement("td");
			actionsCell.className = "row-actions";

			// "Edit"/"Delete" alone read as identical, context-free buttons
			// to a screen reader user browsing a list of buttons - name
			// each one after the profile it acts on.
			var fullName = (p.firstName + " " + p.lastName).trim() || "this profile";

			var editBtn = document.createElement("button");
			editBtn.type = "button";
			editBtn.className = "secondary";
			editBtn.textContent = "Edit";
			editBtn.setAttribute("aria-label", "Edit profile for " + fullName);
			editBtn.addEventListener("click", function () {
				openDialog(p);
			});

			var deleteBtn = document.createElement("button");
			deleteBtn.type = "button";
			deleteBtn.className = "danger";
			deleteBtn.textContent = "Delete";
			deleteBtn.setAttribute("aria-label", "Delete profile for " + fullName);
			deleteBtn.addEventListener("click", function () {
				handleDelete(p.id, fullName);
			});

			actionsCell.appendChild(editBtn);
			actionsCell.appendChild(deleteBtn);
			row.appendChild(actionsCell);

			profilesBody.appendChild(row);
		});
	}

	function openDialog(profile) {
		profileForm.reset();
		if (profile) {
			profileDialogTitle.textContent = "Edit Profile";
			fieldId.value = profile.id;
			fieldFirstName.value = profile.firstName;
			fieldLastName.value = profile.lastName;
			fieldEmail.value = profile.email;
			fieldBio.value = profile.bio;
		} else {
			profileDialogTitle.textContent = "New Profile";
			fieldId.value = "";
		}
		clearError();
		profileDialog.showModal();
	}

	async function loadProfiles() {
		try {
			var res = await api.get("profiles.asp");
			renderProfiles(res.data.data);
		} catch (err) {
			showError(extractErrorMessage(err));
		}
	}

	async function refreshSessionStatus() {
		try {
			var res = await api.get("session.asp");
			var status = res.data.data;
			if (status.active) {
				setCsrfToken(status.csrfToken);
				setSessionUi(true, status.timeoutMinutes);
			} else {
				setCsrfToken(null);
				setSessionUi(false);
			}
		} catch (err) {
			showError(extractErrorMessage(err));
		}
		await loadProfiles();
	}

	async function handleCreateSession() {
		clearError();
		try {
			var res = await api.post("session.asp");
			var status = res.data.data;
			setCsrfToken(status.csrfToken);
			setSessionUi(true, status.timeoutMinutes);
			await loadProfiles();
		} catch (err) {
			showError(extractErrorMessage(err));
		}
	}

	async function handleEndSession() {
		clearError();
		try {
			await api.delete("session.asp");
			setCsrfToken(null);
			setSessionUi(false);
			await loadProfiles();
		} catch (err) {
			showError(extractErrorMessage(err));
		}
	}

	async function handleDelete(id, fullName) {
		var confirmed = await confirmController.confirm("Delete profile for " + fullName + "? This can't be undone.");
		if (!confirmed) return;
		clearError();
		try {
			await api.delete("profiles.asp?id=" + encodeURIComponent(id));
			await loadProfiles();
		} catch (err) {
			showError(extractErrorMessage(err));
		}
	}

	async function handleFormSubmit(e) {
		e.preventDefault();
		clearError();
		var payload = App.buildProfilePayload({
			firstName: fieldFirstName.value,
			lastName: fieldLastName.value,
			email: fieldEmail.value,
			bio: fieldBio.value
		});
		var id = fieldId.value;
		try {
			if (id) {
				await api.put("profiles.asp?id=" + encodeURIComponent(id), payload);
			} else {
				await api.post("profiles.asp", payload);
			}
			profileDialog.close();
			await loadProfiles();
		} catch (err) {
			showError(extractErrorMessage(err));
		}
	}

	btnCreateSession.addEventListener("click", handleCreateSession);
	btnEndSession.addEventListener("click", handleEndSession);
	btnNewProfile.addEventListener("click", function () {
		openDialog(null);
	});
	btnCancelProfile.addEventListener("click", function () {
		profileDialog.close();
	});
	profileForm.addEventListener("submit", handleFormSubmit);

	refreshSessionStatus();
})();
