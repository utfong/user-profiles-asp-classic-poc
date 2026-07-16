<%
// Per-session Access (.accdb) database lifecycle and connection helpers.
// Requires lib/http.asp (sendError) to be included for requireSession.

var ACE_PROVIDER = "Microsoft.ACE.OLEDB.12.0";

function getDbPath() {
	var p = Session("DbPath");
	return (p === null || p === undefined) ? "" : String(p);
}

function hasActiveSession() {
	var p = getDbPath();
	if (p === "") return false;
	var fso = Server.CreateObject("Scripting.FileSystemObject");
	return fso.FileExists(p);
}

// Guard for endpoints that require a provisioned session database. Ends the
// response with 400 if there isn't one.
function requireSession() {
	if (!hasActiveSession()) {
		sendError(400, "no_session", "No active session. Create a session first.");
	}
}

function getSessionsFolder() {
	var fso = Server.CreateObject("Scripting.FileSystemObject");
	var appData = Server.MapPath("/App_Data");
	if (!fso.FolderExists(appData)) fso.CreateFolder(appData);
	var sessionsFolder = appData + "\\sessions";
	if (!fso.FolderExists(sessionsFolder)) fso.CreateFolder(sessionsFolder);
	return sessionsFolder;
}

function buildSessionDbPath() {
	return getSessionsFolder() + "\\" + Session.SessionID + ".accdb";
}

function openConnectionAt(dbPath) {
	var conn = Server.CreateObject("ADODB.Connection");
	conn.Open("Provider=" + ACE_PROVIDER + ";Data Source=" + dbPath + ";");
	return conn;
}

function openConnection() {
	return openConnectionAt(getDbPath());
}

// Creates a new .accdb at dbPath (must not already exist) with the Profiles
// schema. Uses ADOX to create the file itself, then plain DDL for the table.
// try/finally guarantees the connection closes even if CREATE TABLE fails -
// a direct language-level replacement for needing to remember a manual
// cleanup call on every error path.
function createSessionDatabase(dbPath) {
	var catalog = Server.CreateObject("ADOX.Catalog");
	catalog.Create("Provider=" + ACE_PROVIDER + ";Data Source=" + dbPath + ";");
	catalog = null;

	var conn = openConnectionAt(dbPath);
	try {
		conn.Execute(
			"CREATE TABLE Profiles (" +
			"Id COUNTER PRIMARY KEY, " +
			"FirstName TEXT(100), " +
			"LastName TEXT(100), " +
			"Email TEXT(255), " +
			"Bio MEMO, " +
			"CreatedAt DATETIME, " +
			"UpdatedAt DATETIME)"
		);
	} finally {
		conn.Close();
	}
}

// Shared by the explicit "end session" endpoint and global.asa's Session_OnEnd
// cleanup - safe to call even if the file is already gone. Deleting a .accdb
// immediately after Close() can transiently fail (Jet/ACE file-handle
// release lag, confirmed empirically) - swallow it, same as VBScript's
// On Error Resume Next did for the same reason.
function deleteSessionDatabase(dbPath) {
	if (dbPath === "") return;
	var fso = Server.CreateObject("Scripting.FileSystemObject");
	try {
		if (fso.FileExists(dbPath)) fso.DeleteFile(dbPath, true);
	} catch (e) {
		// safe no-op
	}
}
%>
