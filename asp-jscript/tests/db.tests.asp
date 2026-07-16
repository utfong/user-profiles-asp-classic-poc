<%
// Tests for lib/db.asp against a throwaway .accdb - never touches the real
// visitor's session file, since createSessionDatabase/openConnectionAt take
// an explicit path rather than reading it off Session.
//
// Two JScript/COM quirks (confirmed via a throwaway probe page before this
// host was built, see asp-jscript/CLAUDE.md) that this file depends on:
//   - CreateParameter's Size argument must be a real value (e.g. 0), never
//     omitted the way VBScript's "a, , b" skipped-argument syntax allows.
//   - A JScript Date passed as a parameter Value for an adDate column must
//     be converted with .getVarDate() first, or ADO throws a type error.

// Classic ASP JScript has no native sleep - busy-wait against Date.getTime()
// is the standard workaround.
function briefDelay(ms) {
	var start = (new Date()).getTime();
	while ((new Date()).getTime() - start < ms) { /* busy wait */ }
}

function runDbTests() {
	var fso = Server.CreateObject("Scripting.FileSystemObject");
	var testPath = getSessionsFolder() + "\\unittest-" + newCsrfToken() + ".accdb";

	beginTest("db", "createSessionDatabase creates a file with the Profiles schema");
	var createError = null;
	try {
		createSessionDatabase(testPath);
	} catch (e) {
		createError = e;
	}
	assertTrue(createError === null, "createSessionDatabase should not throw" + (createError ? " (" + createError.message + ")" : ""));
	assertTrue(fso.FileExists(testPath), "accdb file should exist on disk");

	if (!fso.FileExists(testPath)) {
		recordResult(false, "cannot continue db tests - database was not created");
		return;
	}

	var conn = openConnectionAt(testPath);

	beginTest("db", "Insert and read back a row via parameterized Command");
	var cmd = Server.CreateObject("ADODB.Command");
	cmd.ActiveConnection = conn;
	cmd.CommandText = "INSERT INTO Profiles (FirstName, LastName, Email, Bio, CreatedAt, UpdatedAt) VALUES (?, ?, ?, ?, ?, ?)";
	cmd.Parameters.Append(cmd.CreateParameter("FirstName", 202, 1, 100, "Test"));
	cmd.Parameters.Append(cmd.CreateParameter("LastName", 202, 1, 100, "User"));
	cmd.Parameters.Append(cmd.CreateParameter("Email", 202, 1, 255, "test@example.com"));
	cmd.Parameters.Append(cmd.CreateParameter("Bio", 203, 1, 5, "hello"));
	cmd.Parameters.Append(cmd.CreateParameter("CreatedAt", 7, 1, 0, (new Date()).getVarDate()));
	cmd.Parameters.Append(cmd.CreateParameter("UpdatedAt", 7, 1, 0, (new Date()).getVarDate()));
	cmd.Execute();

	var rs = conn.Execute("SELECT FirstName, LastName, CreatedAt FROM Profiles");
	assertFalse(rs.EOF, "row should exist after insert");
	if (!rs.EOF) {
		assertEqual("Test", rs("FirstName").Value, "firstname persisted");
		assertEqual("User", rs("LastName").Value, "lastname persisted");

		beginTest("db", "DATETIME column round-trips through new Date(rawValue)");
		var rawCreatedAt = rs("CreatedAt").Value;
		assertTrue(!(rawCreatedAt instanceof Date), "raw ADO value is not itself a JScript Date instance (confirmed quirk)");
		var wrapped = new Date(rawCreatedAt);
		assertTrue(wrapped instanceof Date, "new Date(rawValue) produces a real Date instance");
		assertTrue(!isNaN(wrapped.getFullYear()), "wrapped Date has a valid year");
	}
	rs.Close();

	beginTest("db", "Bio column (MEMO) accepts text longer than 255 chars");
	var longBio = "";
	for (var i = 0; i < 400; i++) longBio += "x";
	var cmd2 = Server.CreateObject("ADODB.Command");
	cmd2.ActiveConnection = conn;
	cmd2.CommandText = "INSERT INTO Profiles (FirstName, LastName, Email, Bio, CreatedAt, UpdatedAt) VALUES (?, ?, ?, ?, ?, ?)";
	cmd2.Parameters.Append(cmd2.CreateParameter("FirstName", 202, 1, 100, "Long"));
	cmd2.Parameters.Append(cmd2.CreateParameter("LastName", 202, 1, 100, "Bio"));
	cmd2.Parameters.Append(cmd2.CreateParameter("Email", 202, 1, 255, ""));
	cmd2.Parameters.Append(cmd2.CreateParameter("Bio", 203, 1, longBio.length, longBio));
	cmd2.Parameters.Append(cmd2.CreateParameter("CreatedAt", 7, 1, 0, (new Date()).getVarDate()));
	cmd2.Parameters.Append(cmd2.CreateParameter("UpdatedAt", 7, 1, 0, (new Date()).getVarDate()));
	var longBioError = null;
	try {
		cmd2.Execute();
	} catch (e2) {
		longBioError = e2;
	}
	assertTrue(longBioError === null, "400-char bio insert should not throw" + (longBioError ? " (" + longBioError.message + ")" : ""));

	conn.Close();

	beginTest("db", "deleteSessionDatabase removes the file");
	// The underlying OS file handle can take a moment to release after
	// conn.Close() (confirmed empirically - see asp-jscript/CLAUDE.md). Retry
	// a few times rather than asserting immediate success: deleteSessionDatabase
	// itself already swallows this by design (best-effort delete, matching
	// asp-vbscript's On Error Resume Next equivalent), so a single immediate
	// check is testing OS timing, not application correctness.
	var deleted = false;
	for (var attempt = 0; attempt < 10 && !deleted; attempt++) {
		if (attempt > 0) briefDelay(100);
		deleteSessionDatabase(testPath);
		deleted = !fso.FileExists(testPath);
	}
	assertTrue(deleted, "test accdb should be gone after deleteSessionDatabase (retried up to 10x with delay for OS file-handle release lag)");

	beginTest("db", "deleteSessionDatabase is a safe no-op on an already-deleted file");
	var noopError = null;
	try {
		deleteSessionDatabase(testPath);
	} catch (e3) {
		noopError = e3;
	}
	assertTrue(noopError === null, "calling delete again should not throw");
}
%>
