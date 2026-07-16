<%
' Per-session Access (.accdb) database lifecycle and connection helpers.
' Requires lib/http.asp (SendError) to be included for RequireSession.

Const ACE_PROVIDER = "Microsoft.ACE.OLEDB.12.0"

Function GetDbPath()
	Dim p
	p = ""
	If Not IsEmpty(Session("DbPath")) And Not IsNull(Session("DbPath")) Then
		p = Session("DbPath")
	End If
	GetDbPath = p
End Function

Function HasActiveSession()
	Dim fso, p
	p = GetDbPath()
	If p = "" Then
		HasActiveSession = False
	Else
		Set fso = Server.CreateObject("Scripting.FileSystemObject")
		HasActiveSession = fso.FileExists(p)
	End If
End Function

' Guard for endpoints that require a provisioned session database. Ends the
' response with 400 if there isn't one.
Sub RequireSession()
	If Not HasActiveSession() Then
		SendError 400, "no_session", "No active session. Create a session first."
	End If
End Sub

Function GetSessionsFolder()
	Dim fso, appData, sessionsFolder
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	appData = Server.MapPath("/App_Data")
	If Not fso.FolderExists(appData) Then
		fso.CreateFolder appData
	End If
	sessionsFolder = appData & "\sessions"
	If Not fso.FolderExists(sessionsFolder) Then
		fso.CreateFolder sessionsFolder
	End If
	GetSessionsFolder = sessionsFolder
End Function

Function BuildSessionDbPath()
	BuildSessionDbPath = GetSessionsFolder() & "\" & Session.SessionID & ".accdb"
End Function

Function OpenConnectionAt(ByVal dbPath)
	Dim conn
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open "Provider=" & ACE_PROVIDER & ";Data Source=" & dbPath & ";"
	Set OpenConnectionAt = conn
End Function

Function OpenConnection()
	Set OpenConnection = OpenConnectionAt(GetDbPath())
End Function

' Creates a new .accdb at dbPath (must not already exist) with the Profiles
' schema. Uses ADOX to create the file itself, then plain DDL for the table.
Sub CreateSessionDatabase(ByVal dbPath)
	Dim catalog, conn
	Set catalog = Server.CreateObject("ADOX.Catalog")
	catalog.Create "Provider=" & ACE_PROVIDER & ";Data Source=" & dbPath & ";"
	Set catalog = Nothing

	Set conn = OpenConnectionAt(dbPath)
	conn.Execute "CREATE TABLE Profiles (" & _
		"Id COUNTER PRIMARY KEY, " & _
		"FirstName TEXT(100), " & _
		"LastName TEXT(100), " & _
		"Email TEXT(255), " & _
		"Bio MEMO, " & _
		"CreatedAt DATETIME, " & _
		"UpdatedAt DATETIME)"
	conn.Close
	Set conn = Nothing
End Sub

' Shared by the explicit "end session" endpoint and global.asa's Session_OnEnd
' cleanup - safe to call even if the file is already gone.
Sub DeleteSessionDatabase(ByVal dbPath)
	Dim fso
	If dbPath = "" Then Exit Sub
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	On Error Resume Next
	If fso.FileExists(dbPath) Then
		fso.DeleteFile dbPath, True
	End If
	On Error Goto 0
End Sub
%>
