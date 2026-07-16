<%@ Language="JScript" %>
<!--#include virtual="/lib/http.asp" -->
<!--#include virtual="/lib/json.asp" -->
<!--#include virtual="/lib/csrf.asp" -->
<!--#include virtual="/lib/db.asp" -->
<%
// Session lifecycle: GET status, POST provision the per-session .accdb,
// DELETE end the session early (also cleaned up automatically by
// Session_OnEnd in global.asa on timeout).

var SESSION_TIMEOUT_MINUTES = 5; // short on purpose so cleanup is easy to observe in a demo

function statusPayload() {
	if (hasActiveSession()) {
		return { active: true, csrfToken: Session("CsrfToken"), timeoutMinutes: Session.Timeout };
	}
	return { active: false };
}

var method = getRequestMethod();

switch (method) {
	case "GET":
		sendData(200, statusPayload());
		break;

	case "POST":
		if (!hasActiveSession()) {
			Session.Timeout = SESSION_TIMEOUT_MINUTES;
			var dbPath = buildSessionDbPath();
			var createError = null;
			try {
				createSessionDatabase(dbPath);
			} catch (e) {
				createError = e;
			}
			if (createError !== null) {
				deleteSessionDatabase(dbPath); // clean up any partially-created file
				sendError(500, "db_create_failed",
					"Could not create the session database. Check that the Microsoft Access " +
					"Database Engine (ACE OLEDB 12.0) is installed and its bitness matches the " +
					"app pool. Detail: " + createError.message);
			}
			Session("DbPath") = dbPath;
			Session("CsrfToken") = newCsrfToken();
			sendData(201, statusPayload());
		} else {
			sendData(200, statusPayload());
		}
		break;

	case "DELETE":
		requireSession();
		requireCsrf();
		deleteSessionDatabase(getDbPath());
		Session("DbPath") = "";
		Session("CsrfToken") = "";
		Session.Abandon();
		sendData(200, statusPayload());
		break;

	default:
		Response.AddHeader("Allow", "GET, POST, DELETE");
		sendError(405, "method_not_allowed", "Supported methods: GET, POST, DELETE.");
}
%>
