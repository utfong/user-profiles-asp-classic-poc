<%@ Language="JScript" %>
<!--#include virtual="/lib/http.asp" -->
<!--#include virtual="/lib/json.asp" -->
<!--#include virtual="/lib/csrf.asp" -->
<!--#include virtual="/lib/db.asp" -->
<%
// Profiles CRUD. Resource selection is via ?id= (no URL Rewrite module
// assumed to be installed), verb selection via the real HTTP method.
// ADO type codes used below (no adovbs.inc-equivalent dependency in JScript
// either): adInteger = 3, adDate = 7, adVarWChar = 202, adLongVarWChar = 203,
// adParamInput = 1.

var FIRSTNAME_MAX_LEN = 100;
var LASTNAME_MAX_LEN = 100;
var EMAIL_MAX_LEN = 255;
// Bio has no max length - it's a MEMO column (see lib/db.asp), unbounded by design.

// Required: firstName, lastName, email. Bio is optional and uncapped.
// Returns "" when valid, otherwise a combined human-readable error message.
function validateProfileFields(firstName, lastName, email) {
	var errors = [];
	if (firstName.replace(/^\s+|\s+$/g, "") === "") errors.push("firstName is required.");
	if (lastName.replace(/^\s+|\s+$/g, "") === "") errors.push("lastName is required.");
	if (email.replace(/^\s+|\s+$/g, "") === "") errors.push("email is required.");
	if (firstName.length > FIRSTNAME_MAX_LEN) errors.push("firstName must be " + FIRSTNAME_MAX_LEN + " characters or fewer.");
	if (lastName.length > LASTNAME_MAX_LEN) errors.push("lastName must be " + LASTNAME_MAX_LEN + " characters or fewer.");
	if (email.length > EMAIL_MAX_LEN) errors.push("email must be " + EMAIL_MAX_LEN + " characters or fewer.");
	return errors.join(" ");
}

// ADO DATETIME fields come back as a raw VarDate, not a JScript Date
// (confirmed empirically - see CLAUDE.md) - wrap with new Date(...) here so
// jsonEncode only ever has to deal with real Date instances.
function rowToObject(rs) {
	return {
		id: rs("Id").Value,
		firstName: rs("FirstName").Value || "",
		lastName: rs("LastName").Value || "",
		email: rs("Email").Value || "",
		bio: rs("Bio").Value || "",
		createdAt: new Date(rs("CreatedAt").Value),
		updatedAt: new Date(rs("UpdatedAt").Value)
	};
}

function fieldOf(payload, key) {
	var v = payload[key];
	return (v === undefined || v === null) ? "" : String(v);
}

function isValidId(s) {
	if (s === "" || isNaN(Number(s))) return false;
	var n = Number(s);
	return !(n <= 0 || n !== Math.floor(n) || n > 2147483647);
}

// Runs a write Command inside Application.Lock/Unlock, guaranteeing Unlock
// even if Execute fails via finally - a direct language-level replacement
// for VBScript's On Error Resume Next + manual Err.Number save/restore
// dance around the same guarantee.
function executeLocked(cmd, conn) {
	var execError = null;
	Application.Lock();
	try {
		cmd.Execute();
	} catch (e) {
		execError = e;
	} finally {
		Application.Unlock();
	}
	if (execError !== null) {
		conn.Close();
		sendError(500, "write_failed", "Database write failed: " + execError.message);
	}
}

// Parses the JSON body into an object, or ends the response with 400.
function parseBodyOrFail() {
	var body = readRawBody();
	try {
		return jsonParse(body);
	} catch (e) {
		sendError(400, "invalid_json", "Request body is not valid JSON.");
	}
}

function handleList() {
	if (!hasActiveSession()) {
		sendData(200, []); // no session yet - nothing to list
	}
	var conn = openConnection();
	var cmd = Server.CreateObject("ADODB.Command");
	cmd.ActiveConnection = conn;
	cmd.CommandText = "SELECT Id, FirstName, LastName, Email, Bio, CreatedAt, UpdatedAt FROM Profiles ORDER BY Id";
	var rs = cmd.Execute();
	var list = [];
	while (!rs.EOF) {
		list.push(rowToObject(rs));
		rs.MoveNext();
	}
	rs.Close();
	conn.Close();
	sendData(200, list);
}

function handleGet(idText) {
	if (!isValidId(idText)) {
		sendError(400, "invalid_id", "id must be a positive integer.");
	}
	if (!hasActiveSession()) {
		sendError(404, "not_found", "Profile " + idText + " not found."); // no session yet - nothing exists
	}
	var idVal = Number(idText);

	var conn = openConnection();
	var cmd = Server.CreateObject("ADODB.Command");
	cmd.ActiveConnection = conn;
	cmd.CommandText = "SELECT Id, FirstName, LastName, Email, Bio, CreatedAt, UpdatedAt FROM Profiles WHERE Id = ?";
	cmd.Parameters.Append(cmd.CreateParameter("Id", 3, 1, 0, idVal));
	var rs = cmd.Execute();
	if (rs.EOF) {
		rs.Close();
		conn.Close();
		sendError(404, "not_found", "Profile " + idVal + " not found.");
	} else {
		var d = rowToObject(rs);
		rs.Close();
		conn.Close();
		sendData(200, d);
	}
}

function handleCreate() {
	var payload = parseBodyOrFail();

	var firstName = fieldOf(payload, "firstName");
	var lastName = fieldOf(payload, "lastName");
	var email = fieldOf(payload, "email");
	var bio = fieldOf(payload, "bio");

	var validationError = validateProfileFields(firstName, lastName, email);
	if (validationError !== "") {
		sendError(400, "validation_error", validationError);
	}

	var conn = openConnection();
	var cmd = Server.CreateObject("ADODB.Command");
	cmd.ActiveConnection = conn;
	cmd.CommandText = "INSERT INTO Profiles (FirstName, LastName, Email, Bio, CreatedAt, UpdatedAt) VALUES (?, ?, ?, ?, ?, ?)";
	cmd.Parameters.Append(cmd.CreateParameter("FirstName", 202, 1, 100, firstName));
	cmd.Parameters.Append(cmd.CreateParameter("LastName", 202, 1, 100, lastName));
	cmd.Parameters.Append(cmd.CreateParameter("Email", 202, 1, 255, email));
	cmd.Parameters.Append(cmd.CreateParameter("Bio", 203, 1, Math.max(1, bio.length), bio)); // adLongVarWChar - Bio is a MEMO column, no 255-char cap. Size must be >=1 - ADO rejects 0 ("Parameter object is improperly defined") for an empty bio.
	var now = (new Date()).getVarDate();
	cmd.Parameters.Append(cmd.CreateParameter("CreatedAt", 7, 1, 0, now));
	cmd.Parameters.Append(cmd.CreateParameter("UpdatedAt", 7, 1, 0, now));

	executeLocked(cmd, conn);

	var idRs = conn.Execute("SELECT @@IDENTITY");
	var newId = Number(idRs(0).Value);
	idRs.Close();

	var getCmd = Server.CreateObject("ADODB.Command");
	getCmd.ActiveConnection = conn;
	getCmd.CommandText = "SELECT Id, FirstName, LastName, Email, Bio, CreatedAt, UpdatedAt FROM Profiles WHERE Id = ?";
	getCmd.Parameters.Append(getCmd.CreateParameter("Id", 3, 1, 0, newId));
	var getRs = getCmd.Execute();
	var d = rowToObject(getRs);
	getRs.Close();
	conn.Close();
	sendData(201, d);
}

function handleUpdate(idText) {
	if (!isValidId(idText)) {
		sendError(400, "invalid_id", "id must be a positive integer.");
	}
	var idVal = Number(idText);

	var payload = parseBodyOrFail();

	var firstName = fieldOf(payload, "firstName");
	var lastName = fieldOf(payload, "lastName");
	var email = fieldOf(payload, "email");
	var bio = fieldOf(payload, "bio");

	var validationError = validateProfileFields(firstName, lastName, email);
	if (validationError !== "") {
		sendError(400, "validation_error", validationError);
	}

	var conn = openConnection();

	var checkCmd = Server.CreateObject("ADODB.Command");
	checkCmd.ActiveConnection = conn;
	checkCmd.CommandText = "SELECT Id FROM Profiles WHERE Id = ?";
	checkCmd.Parameters.Append(checkCmd.CreateParameter("Id", 3, 1, 0, idVal));
	var checkRs = checkCmd.Execute();
	if (checkRs.EOF) {
		checkRs.Close();
		conn.Close();
		sendError(404, "not_found", "Profile " + idVal + " not found.");
	}
	checkRs.Close();

	var cmd = Server.CreateObject("ADODB.Command");
	cmd.ActiveConnection = conn;
	cmd.CommandText = "UPDATE Profiles SET FirstName=?, LastName=?, Email=?, Bio=?, UpdatedAt=? WHERE Id=?";
	cmd.Parameters.Append(cmd.CreateParameter("FirstName", 202, 1, 100, firstName));
	cmd.Parameters.Append(cmd.CreateParameter("LastName", 202, 1, 100, lastName));
	cmd.Parameters.Append(cmd.CreateParameter("Email", 202, 1, 255, email));
	cmd.Parameters.Append(cmd.CreateParameter("Bio", 203, 1, Math.max(1, bio.length), bio)); // adLongVarWChar - Bio is a MEMO column, no 255-char cap. Size must be >=1 - ADO rejects 0 ("Parameter object is improperly defined") for an empty bio.
	cmd.Parameters.Append(cmd.CreateParameter("UpdatedAt", 7, 1, 0, (new Date()).getVarDate()));
	cmd.Parameters.Append(cmd.CreateParameter("Id", 3, 1, 0, idVal));

	executeLocked(cmd, conn);

	var getCmd = Server.CreateObject("ADODB.Command");
	getCmd.ActiveConnection = conn;
	getCmd.CommandText = "SELECT Id, FirstName, LastName, Email, Bio, CreatedAt, UpdatedAt FROM Profiles WHERE Id = ?";
	getCmd.Parameters.Append(getCmd.CreateParameter("Id", 3, 1, 0, idVal));
	var getRs = getCmd.Execute();
	var d = rowToObject(getRs);
	getRs.Close();
	conn.Close();
	sendData(200, d);
}

function handleDelete(idText) {
	if (!isValidId(idText)) {
		sendError(400, "invalid_id", "id must be a positive integer.");
	}
	var idVal = Number(idText);

	var conn = openConnection();
	var checkCmd = Server.CreateObject("ADODB.Command");
	checkCmd.ActiveConnection = conn;
	checkCmd.CommandText = "SELECT Id FROM Profiles WHERE Id = ?";
	checkCmd.Parameters.Append(checkCmd.CreateParameter("Id", 3, 1, 0, idVal));
	var checkRs = checkCmd.Execute();
	if (checkRs.EOF) {
		checkRs.Close();
		conn.Close();
		sendError(404, "not_found", "Profile " + idVal + " not found.");
	}
	checkRs.Close();

	var cmd = Server.CreateObject("ADODB.Command");
	cmd.ActiveConnection = conn;
	cmd.CommandText = "DELETE FROM Profiles WHERE Id = ?";
	cmd.Parameters.Append(cmd.CreateParameter("Id", 3, 1, 0, idVal));

	executeLocked(cmd, conn);

	conn.Close();
	sendData(204, null);
}

// --- Dispatch ---
// GET is a safe/read-only verb and works with or without an active session
// (an inactive session just has nothing to read yet). POST/PUT/DELETE mutate
// data, so they require a session - and therefore CSRF - first.

var method = getRequestMethod();
var id = requestParam(Request.QueryString("id"));

switch (method) {
	case "GET":
		if (id === "") {
			handleList();
		} else {
			handleGet(id);
		}
		break;

	case "POST":
		requireSession();
		requireCsrf();
		handleCreate();
		break;

	case "PUT":
		requireSession();
		requireCsrf();
		if (id === "") {
			sendError(400, "missing_id", "id query parameter is required for update.");
		}
		handleUpdate(id);
		break;

	case "DELETE":
		requireSession();
		requireCsrf();
		if (id === "") {
			sendError(400, "missing_id", "id query parameter is required for delete.");
		}
		handleDelete(id);
		break;

	default:
		Response.AddHeader("Allow", "GET, POST, PUT, DELETE");
		sendError(405, "method_not_allowed", "Supported methods: GET, POST, PUT, DELETE.");
}
%>
