<%
' Tests for lib/db.asp against a throwaway .accdb - never touches the real
' visitor's session file, since CreateSessionDatabase/OpenConnectionAt take
' an explicit path rather than reading it off Session.

Sub RunDbTests()
	Dim fso, testPath
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	testPath = GetSessionsFolder() & "\unittest-" & NewCsrfToken() & ".accdb"

	BeginTest "db", "CreateSessionDatabase creates a file with the Profiles schema"
	Dim createErrNum, createErrDesc
	On Error Resume Next
	CreateSessionDatabase testPath
	createErrNum = Err.Number
	createErrDesc = Err.Description
	Err.Clear
	On Error Goto 0
	AssertTrue (createErrNum = 0), "CreateSessionDatabase should not raise (" & createErrDesc & ")"
	AssertTrue fso.FileExists(testPath), "accdb file should exist on disk"

	If Not fso.FileExists(testPath) Then
		RecordResult False, "cannot continue db tests - database was not created"
		Exit Sub
	End If

	Dim conn
	Set conn = OpenConnectionAt(testPath)

	BeginTest "db", "Insert and read back a row via parameterized Command"
	Dim cmd
	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = conn
	cmd.CommandText = "INSERT INTO Profiles (FirstName, LastName, Email, Bio, CreatedAt, UpdatedAt) VALUES (?, ?, ?, ?, ?, ?)"
	cmd.Parameters.Append cmd.CreateParameter("FirstName", 202, 1, 100, "Test")
	cmd.Parameters.Append cmd.CreateParameter("LastName", 202, 1, 100, "User")
	cmd.Parameters.Append cmd.CreateParameter("Email", 202, 1, 255, "test@example.com")
	cmd.Parameters.Append cmd.CreateParameter("Bio", 203, 1, 5, "hello")
	cmd.Parameters.Append cmd.CreateParameter("CreatedAt", 7, 1, , Now())
	cmd.Parameters.Append cmd.CreateParameter("UpdatedAt", 7, 1, , Now())
	cmd.Execute

	Dim rs
	Set rs = conn.Execute("SELECT FirstName, LastName FROM Profiles")
	AssertFalse rs.EOF, "row should exist after insert"
	If Not rs.EOF Then
		AssertEqual "Test", rs("FirstName").Value, "firstname persisted"
		AssertEqual "User", rs("LastName").Value, "lastname persisted"
	End If
	rs.Close

	BeginTest "db", "Bio column (MEMO) accepts text longer than 255 chars"
	Dim longBio, cmd2, longErrNum, longErrDesc
	longBio = String(400, "x")
	Set cmd2 = Server.CreateObject("ADODB.Command")
	Set cmd2.ActiveConnection = conn
	cmd2.CommandText = "INSERT INTO Profiles (FirstName, LastName, Email, Bio, CreatedAt, UpdatedAt) VALUES (?, ?, ?, ?, ?, ?)"
	cmd2.Parameters.Append cmd2.CreateParameter("FirstName", 202, 1, 100, "Long")
	cmd2.Parameters.Append cmd2.CreateParameter("LastName", 202, 1, 100, "Bio")
	cmd2.Parameters.Append cmd2.CreateParameter("Email", 202, 1, 255, "")
	cmd2.Parameters.Append cmd2.CreateParameter("Bio", 203, 1, Len(longBio), longBio)
	cmd2.Parameters.Append cmd2.CreateParameter("CreatedAt", 7, 1, , Now())
	cmd2.Parameters.Append cmd2.CreateParameter("UpdatedAt", 7, 1, , Now())
	On Error Resume Next
	cmd2.Execute
	longErrNum = Err.Number
	longErrDesc = Err.Description
	Err.Clear
	On Error Goto 0
	AssertTrue (longErrNum = 0), "400-char bio insert should not raise (" & longErrDesc & ")"

	conn.Close

	BeginTest "db", "DeleteSessionDatabase removes the file"
	DeleteSessionDatabase testPath
	AssertFalse fso.FileExists(testPath), "test accdb should be gone after DeleteSessionDatabase"

	BeginTest "db", "DeleteSessionDatabase is a safe no-op on an already-deleted file"
	Dim noopErrNum
	On Error Resume Next
	DeleteSessionDatabase testPath
	noopErrNum = Err.Number
	Err.Clear
	On Error Goto 0
	AssertTrue (noopErrNum = 0), "calling delete again should not raise"
End Sub
%>
