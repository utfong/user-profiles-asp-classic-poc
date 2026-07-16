<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include virtual="/lib/http.asp" -->
<!--#include virtual="/lib/json.asp" -->
<!--#include virtual="/lib/csrf.asp" -->
<!--#include virtual="/lib/db.asp" -->
<%
' Profiles CRUD. Resource selection is via ?id= (no URL Rewrite module
' assumed to be installed), verb selection via the real HTTP method.
' ADO type codes used below (no adovbs.inc dependency):
'   adInteger = 3, adDate = 7, adVarWChar = 202, adLongVarWChar = 203, adParamInput = 1

Const FIRSTNAME_MAX_LEN = 100
Const LASTNAME_MAX_LEN = 100
Const EMAIL_MAX_LEN = 255
' Bio has no max length - it's a MEMO column (see lib/db.asp), unbounded by design.

' Required: firstName, lastName, email. Bio is optional and uncapped.
' Returns "" when valid, otherwise a combined human-readable error message.
Function ValidateProfileFields(ByVal firstName, ByVal lastName, ByVal email)
	Dim errors
	errors = ""
	If Trim(firstName) = "" Then errors = errors & "firstName is required. "
	If Trim(lastName) = "" Then errors = errors & "lastName is required. "
	If Trim(email) = "" Then errors = errors & "email is required. "
	If Len(firstName) > FIRSTNAME_MAX_LEN Then errors = errors & "firstName must be " & FIRSTNAME_MAX_LEN & " characters or fewer. "
	If Len(lastName) > LASTNAME_MAX_LEN Then errors = errors & "lastName must be " & LASTNAME_MAX_LEN & " characters or fewer. "
	If Len(email) > EMAIL_MAX_LEN Then errors = errors & "email must be " & EMAIL_MAX_LEN & " characters or fewer. "
	ValidateProfileFields = Trim(errors)
End Function

' ADO rejects a Size of 0 for a variable-length parameter ("Parameter object
' is improperly defined") - always pass at least 1, even for an empty string.
Function BioParamSize(ByVal bio)
	If Len(bio) < 1 Then
		BioParamSize = 1
	Else
		BioParamSize = Len(bio)
	End If
End Function

Function NullToEmpty(ByVal v)
	If IsNull(v) Then
		NullToEmpty = ""
	Else
		NullToEmpty = v
	End If
End Function

Function RowToDict(ByVal rs)
	Dim d
	Set d = Server.CreateObject("Scripting.Dictionary")
	d.Add "id", CLng(rs("Id").Value)
	d.Add "firstName", NullToEmpty(rs("FirstName").Value)
	d.Add "lastName", NullToEmpty(rs("LastName").Value)
	d.Add "email", NullToEmpty(rs("Email").Value)
	d.Add "bio", NullToEmpty(rs("Bio").Value)
	d.Add "createdAt", rs("CreatedAt").Value
	d.Add "updatedAt", rs("UpdatedAt").Value
	Set RowToDict = d
End Function

Function DictGet(ByVal dict, ByVal key)
	If dict.Exists(key) And Not IsNull(dict.Item(key)) Then
		DictGet = dict.Item(key)
	Else
		DictGet = ""
	End If
End Function

Function IsValidId(ByVal s)
	If s = "" Or Not IsNumeric(s) Then
		IsValidId = False
	ElseIf CDbl(s) <= 0 Or CDbl(s) <> Int(CDbl(s)) Or CDbl(s) > 2147483647 Then
		IsValidId = False ' also rejects non-integers and values CLng can't hold, to avoid an overflow error
	Else
		IsValidId = True
	End If
End Function

' Runs a write Command inside Application.Lock/Unlock, guaranteeing Unlock
' even if Execute fails - otherwise a single failed write would leave the
' whole application locked and hang every other request on the site.
Sub ExecuteLocked(ByVal cmd, ByVal conn)
	Dim execErrNum, execErrDesc
	Application.Lock
	On Error Resume Next
	cmd.Execute
	execErrNum = Err.Number
	execErrDesc = Err.Description
	Err.Clear
	On Error Goto 0
	Application.Unlock
	If execErrNum <> 0 Then
		conn.Close
		SendError 500, "write_failed", "Database write failed: " & execErrDesc
	End If
End Sub

' Parses the JSON body into a Dictionary, or ends the response with 400.
Function ParseBodyOrFail()
	Dim body, payload
	body = ReadRawBody()
	On Error Resume Next
	Set payload = JsonParse(body)
	If Err.Number <> 0 Then
		Err.Clear
		On Error Goto 0
		SendError 400, "invalid_json", "Request body is not valid JSON."
	End If
	On Error Goto 0
	Set ParseBodyOrFail = payload
End Function

Sub HandleList()
	Dim conn, cmd, rs, list(), count
	If Not HasActiveSession() Then
		ReDim list(-1) ' no session yet - nothing to list
		SendData 200, list
	End If
	Set conn = OpenConnection()
	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = conn
	cmd.CommandText = "SELECT Id, FirstName, LastName, Email, Bio, CreatedAt, UpdatedAt FROM Profiles ORDER BY Id"
	Set rs = cmd.Execute()
	ReDim list(-1)
	count = 0
	Do While Not rs.EOF
		ReDim Preserve list(count)
		Set list(count) = RowToDict(rs)
		count = count + 1
		rs.MoveNext
	Loop
	rs.Close
	conn.Close
	SendData 200, list
End Sub

Sub HandleGet(ByVal idText)
	If Not IsValidId(idText) Then
		SendError 400, "invalid_id", "id must be a positive integer."
	End If
	If Not HasActiveSession() Then
		SendError 404, "not_found", "Profile " & idText & " not found." ' no session yet - nothing exists
	End If
	Dim idVal
	idVal = CLng(idText)

	Dim conn, cmd, rs
	Set conn = OpenConnection()
	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = conn
	cmd.CommandText = "SELECT Id, FirstName, LastName, Email, Bio, CreatedAt, UpdatedAt FROM Profiles WHERE Id = ?"
	cmd.Parameters.Append cmd.CreateParameter("Id", 3, 1, , idVal)
	Set rs = cmd.Execute()
	If rs.EOF Then
		rs.Close
		conn.Close
		SendError 404, "not_found", "Profile " & idVal & " not found."
	Else
		Dim d
		Set d = RowToDict(rs)
		rs.Close
		conn.Close
		SendData 200, d
	End If
End Sub

Sub HandleCreate()
	Dim payload
	Set payload = ParseBodyOrFail()

	Dim firstName, lastName, email, bio
	firstName = DictGet(payload, "firstName")
	lastName = DictGet(payload, "lastName")
	email = DictGet(payload, "email")
	bio = DictGet(payload, "bio")

	Dim createValidationError
	createValidationError = ValidateProfileFields(firstName, lastName, email)
	If createValidationError <> "" Then
		SendError 400, "validation_error", createValidationError
	End If

	Dim conn, cmd
	Set conn = OpenConnection()
	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = conn
	cmd.CommandText = "INSERT INTO Profiles (FirstName, LastName, Email, Bio, CreatedAt, UpdatedAt) VALUES (?, ?, ?, ?, ?, ?)"
	cmd.Parameters.Append cmd.CreateParameter("FirstName", 202, 1, 100, firstName)
	cmd.Parameters.Append cmd.CreateParameter("LastName", 202, 1, 100, lastName)
	cmd.Parameters.Append cmd.CreateParameter("Email", 202, 1, 255, email)
	cmd.Parameters.Append cmd.CreateParameter("Bio", 203, 1, BioParamSize(bio), bio) ' adLongVarWChar - Bio is a MEMO column, no 255-char cap
	cmd.Parameters.Append cmd.CreateParameter("CreatedAt", 7, 1, , Now())
	cmd.Parameters.Append cmd.CreateParameter("UpdatedAt", 7, 1, , Now())

	ExecuteLocked cmd, conn

	Dim idRs, newId
	Set idRs = conn.Execute("SELECT @@IDENTITY")
	newId = CLng(idRs(0).Value)
	idRs.Close

	Dim getCmd, getRs, d
	Set getCmd = Server.CreateObject("ADODB.Command")
	Set getCmd.ActiveConnection = conn
	getCmd.CommandText = "SELECT Id, FirstName, LastName, Email, Bio, CreatedAt, UpdatedAt FROM Profiles WHERE Id = ?"
	getCmd.Parameters.Append getCmd.CreateParameter("Id", 3, 1, , newId)
	Set getRs = getCmd.Execute()
	Set d = RowToDict(getRs)
	getRs.Close
	conn.Close
	SendData 201, d
End Sub

Sub HandleUpdate(ByVal idText)
	If Not IsValidId(idText) Then
		SendError 400, "invalid_id", "id must be a positive integer."
	End If
	Dim idVal
	idVal = CLng(idText)

	Dim payload
	Set payload = ParseBodyOrFail()

	Dim firstName, lastName, email, bio
	firstName = DictGet(payload, "firstName")
	lastName = DictGet(payload, "lastName")
	email = DictGet(payload, "email")
	bio = DictGet(payload, "bio")

	Dim updateValidationError
	updateValidationError = ValidateProfileFields(firstName, lastName, email)
	If updateValidationError <> "" Then
		SendError 400, "validation_error", updateValidationError
	End If

	Dim conn
	Set conn = OpenConnection()

	Dim checkCmd, checkRs
	Set checkCmd = Server.CreateObject("ADODB.Command")
	Set checkCmd.ActiveConnection = conn
	checkCmd.CommandText = "SELECT Id FROM Profiles WHERE Id = ?"
	checkCmd.Parameters.Append checkCmd.CreateParameter("Id", 3, 1, , idVal)
	Set checkRs = checkCmd.Execute()
	If checkRs.EOF Then
		checkRs.Close
		conn.Close
		SendError 404, "not_found", "Profile " & idVal & " not found."
	End If
	checkRs.Close

	Dim cmd
	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = conn
	cmd.CommandText = "UPDATE Profiles SET FirstName=?, LastName=?, Email=?, Bio=?, UpdatedAt=? WHERE Id=?"
	cmd.Parameters.Append cmd.CreateParameter("FirstName", 202, 1, 100, firstName)
	cmd.Parameters.Append cmd.CreateParameter("LastName", 202, 1, 100, lastName)
	cmd.Parameters.Append cmd.CreateParameter("Email", 202, 1, 255, email)
	cmd.Parameters.Append cmd.CreateParameter("Bio", 203, 1, BioParamSize(bio), bio) ' adLongVarWChar - Bio is a MEMO column, no 255-char cap
	cmd.Parameters.Append cmd.CreateParameter("UpdatedAt", 7, 1, , Now())
	cmd.Parameters.Append cmd.CreateParameter("Id", 3, 1, , idVal)

	ExecuteLocked cmd, conn

	Dim getCmd, getRs, d
	Set getCmd = Server.CreateObject("ADODB.Command")
	Set getCmd.ActiveConnection = conn
	getCmd.CommandText = "SELECT Id, FirstName, LastName, Email, Bio, CreatedAt, UpdatedAt FROM Profiles WHERE Id = ?"
	getCmd.Parameters.Append getCmd.CreateParameter("Id", 3, 1, , idVal)
	Set getRs = getCmd.Execute()
	Set d = RowToDict(getRs)
	getRs.Close
	conn.Close
	SendData 200, d
End Sub

Sub HandleDelete(ByVal idText)
	If Not IsValidId(idText) Then
		SendError 400, "invalid_id", "id must be a positive integer."
	End If
	Dim idVal
	idVal = CLng(idText)

	Dim conn, checkCmd, checkRs
	Set conn = OpenConnection()
	Set checkCmd = Server.CreateObject("ADODB.Command")
	Set checkCmd.ActiveConnection = conn
	checkCmd.CommandText = "SELECT Id FROM Profiles WHERE Id = ?"
	checkCmd.Parameters.Append checkCmd.CreateParameter("Id", 3, 1, , idVal)
	Set checkRs = checkCmd.Execute()
	If checkRs.EOF Then
		checkRs.Close
		conn.Close
		SendError 404, "not_found", "Profile " & idVal & " not found."
	End If
	checkRs.Close

	Dim cmd
	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = conn
	cmd.CommandText = "DELETE FROM Profiles WHERE Id = ?"
	cmd.Parameters.Append cmd.CreateParameter("Id", 3, 1, , idVal)

	ExecuteLocked cmd, conn

	conn.Close
	SendData 204, Empty
End Sub

' --- Dispatch ---
' GET is a safe/read-only verb and works with or without an active session
' (an inactive session just has nothing to read yet). POST/PUT/DELETE mutate
' data, so they require a session - and therefore CSRF - first.

Dim method, id
method = GetRequestMethod()
id = Request.QueryString("id")

Select Case method
	Case "GET"
		If id = "" Then
			HandleList
		Else
			HandleGet id
		End If

	Case "POST"
		RequireSession
		RequireCsrf
		HandleCreate

	Case "PUT"
		RequireSession
		RequireCsrf
		If id = "" Then
			SendError 400, "missing_id", "id query parameter is required for update."
		End If
		HandleUpdate id

	Case "DELETE"
		RequireSession
		RequireCsrf
		If id = "" Then
			SendError 400, "missing_id", "id query parameter is required for delete."
		End If
		HandleDelete id

	Case Else
		Response.AddHeader "Allow", "GET, POST, PUT, DELETE"
		SendError 405, "method_not_allowed", "Supported methods: GET, POST, PUT, DELETE."
End Select
%>
