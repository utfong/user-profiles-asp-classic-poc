<%
' Live HTTP round-trip tests: the API is self-called over real HTTP via
' MSXML2.ServerXMLHTTP, because verb routing and header-based CSRF checking
' can only be observed through an actual request, not in-process. This
' automates the same flow that was previously verified by hand with
' PowerShell (session lifecycle, CRUD, CSRF rejection, cleanup).
'
' Note: response bodies here are checked with InStr/Extract* string helpers
' rather than JsonParse, because every real API response is envelope-wrapped
' ({"data": ...} / {"error": {...}}) and JsonParse only supports flat
' objects by design (see lib/json.asp) - it can't parse its own API's
' nested envelope.

Function GetSelfBaseUrl()
	Dim scheme
	If Request.ServerVariables("HTTPS") = "on" Then
		scheme = "https"
	Else
		scheme = "http"
	End If
	GetSelfBaseUrl = scheme & "://" & Request.ServerVariables("HTTP_HOST")
End Function

Function ExtractCookiePair(ByVal setCookieHeader)
	If setCookieHeader = "" Then
		ExtractCookiePair = ""
	Else
		ExtractCookiePair = Trim(Split(setCookieHeader, ";")(0))
	End If
End Function

Function ExtractJsonStringField(ByVal json, ByVal fieldName)
	Dim q, marker, startPos, endPos
	q = Chr(34)
	marker = q & fieldName & q & ":" & q
	startPos = InStr(json, marker)
	If startPos = 0 Then
		ExtractJsonStringField = ""
		Exit Function
	End If
	startPos = startPos + Len(marker)
	endPos = InStr(startPos, json, q)
	If endPos = 0 Then
		ExtractJsonStringField = ""
	Else
		ExtractJsonStringField = Mid(json, startPos, endPos - startPos)
	End If
End Function

Function ExtractJsonNumberField(ByVal json, ByVal fieldName)
	Dim marker, startPos, endPos, c
	marker = Chr(34) & fieldName & Chr(34) & ":"
	startPos = InStr(json, marker)
	If startPos = 0 Then
		ExtractJsonNumberField = ""
		Exit Function
	End If
	startPos = startPos + Len(marker)
	endPos = startPos
	Do While endPos <= Len(json)
		c = Mid(json, endPos, 1)
		If (c >= "0" And c <= "9") Or c = "-" Or c = "." Then
			endPos = endPos + 1
		Else
			Exit Do
		End If
	Loop
	ExtractJsonNumberField = Mid(json, startPos, endPos - startPos)
End Function

' Returns Nothing (and records a failure) instead of raising, so one flaky
' request can't take down the whole test run.
Function HttpCall(ByVal method, ByVal url, ByVal cookieHeader, ByVal csrfToken, ByVal jsonBody)
	Dim http
	On Error Resume Next
	Set http = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
	http.Open method, url, False
	If jsonBody <> "" Then http.setRequestHeader "Content-Type", "application/json"
	If cookieHeader <> "" Then http.setRequestHeader "Cookie", cookieHeader
	If csrfToken <> "" Then http.setRequestHeader "X-CSRF-Token", csrfToken
	If jsonBody <> "" Then
		http.send jsonBody
	Else
		http.send
	End If
	If Err.Number <> 0 Then
		RecordResult False, "HTTP " & method & " " & url & " failed to complete: " & Err.Description
		Set http = Nothing
	End If
	Err.Clear
	On Error Goto 0
	Set HttpCall = http
End Function

Sub RunApiTests()
	Dim baseUrl, http, cookieHeader, csrf, newId, q, longName
	baseUrl = GetSelfBaseUrl()
	q = Chr(34)

	BeginTest "api", "GET /api/profiles.asp works with no session at all"
	Set http = HttpCall("GET", baseUrl & "/api/profiles.asp", "", "", "")
	If http Is Nothing Then Exit Sub
	AssertEqual 200, http.status, "status"
	AssertTrue InStr(http.responseText, q & "data" & q & ":[]") > 0, "empty list when no session"

	BeginTest "api", "POST /api/profiles.asp without a session is rejected"
	Set http = HttpCall("POST", baseUrl & "/api/profiles.asp", "", "", _
		"{" & q & "firstName" & q & ":" & q & "X" & q & "," & q & "lastName" & q & ":" & q & "Y" & q & "}")
	If http Is Nothing Then Exit Sub
	AssertEqual 400, http.status, "status"
	AssertTrue InStr(http.responseText, "no_session") > 0, "error code"

	BeginTest "api", "POST /api/session.asp creates a session"
	Set http = HttpCall("POST", baseUrl & "/api/session.asp", "", "", "")
	If http Is Nothing Then Exit Sub
	AssertEqual 201, http.status, "status"
	cookieHeader = ExtractCookiePair(http.getResponseHeader("Set-Cookie"))
	csrf = ExtractJsonStringField(http.responseText, "csrfToken")
	AssertTrue cookieHeader <> "", "session cookie should be set"
	AssertEqual 32, Len(csrf), "csrf token length"

	BeginTest "api", "POST /api/profiles.asp without email is rejected"
	Set http = HttpCall("POST", baseUrl & "/api/profiles.asp", cookieHeader, csrf, _
		"{" & q & "firstName" & q & ":" & q & "Test" & q & "," & q & "lastName" & q & ":" & q & "User" & q & "}")
	If Not (http Is Nothing) Then
		AssertEqual 400, http.status, "status"
		AssertTrue InStr(http.responseText, "validation_error") > 0, "error code"
		AssertTrue InStr(http.responseText, "email is required") > 0, "message mentions missing email"
	End If

	BeginTest "api", "POST /api/profiles.asp with an oversized firstName is rejected"
	' 101 chars - one over api/profiles.asp's FIRSTNAME_MAX_LEN (not in scope
	' here; that constant lives in an entry-point page, not an includable lib).
	longName = String(101, "A")
	Set http = HttpCall("POST", baseUrl & "/api/profiles.asp", cookieHeader, csrf, _
		"{" & q & "firstName" & q & ":" & q & longName & q & "," & _
		q & "lastName" & q & ":" & q & "User" & q & "," & _
		q & "email" & q & ":" & q & "t@example.com" & q & "}")
	If Not (http Is Nothing) Then
		AssertEqual 400, http.status, "status"
		AssertTrue InStr(http.responseText, "firstName must be 100 characters or fewer") > 0, "message mentions max length"
	End If

	BeginTest "api", "POST /api/profiles.asp creates a profile"
	Set http = HttpCall("POST", baseUrl & "/api/profiles.asp", cookieHeader, csrf, _
		"{" & q & "firstName" & q & ":" & q & "Test" & q & "," & _
		q & "lastName" & q & ":" & q & "User" & q & "," & _
		q & "email" & q & ":" & q & "t@example.com" & q & "," & _
		q & "bio" & q & ":" & q & "hi" & q & "}")
	If http Is Nothing Then
		HttpCall "DELETE", baseUrl & "/api/session.asp", cookieHeader, csrf, ""
		Exit Sub
	End If
	AssertEqual 201, http.status, "status"
	newId = ExtractJsonNumberField(http.responseText, "id")
	AssertTrue newId <> "", "should return a new id"

	BeginTest "api", "PUT without a CSRF header is rejected and does not modify data"
	Set http = HttpCall("PUT", baseUrl & "/api/profiles.asp?id=" & newId, cookieHeader, "", _
		"{" & q & "firstName" & q & ":" & q & "TAMPERED" & q & "," & _
		q & "lastName" & q & ":" & q & "User" & q & "," & _
		q & "email" & q & ":" & q & "t@example.com" & q & "," & _
		q & "bio" & q & ":" & q & "hi" & q & "}")
	If Not (http Is Nothing) Then
		AssertEqual 403, http.status, "status"
	End If

	BeginTest "api", "GET after the failed tamper still shows the original data"
	Set http = HttpCall("GET", baseUrl & "/api/profiles.asp?id=" & newId, cookieHeader, "", "")
	If Not (http Is Nothing) Then
		AssertTrue InStr(http.responseText, "TAMPERED") = 0, "tampered value should not have persisted"
		AssertTrue InStr(http.responseText, q & "firstName" & q & ":" & q & "Test" & q) > 0, "original value should remain"
	End If

	BeginTest "api", "PUT with a valid CSRF header succeeds"
	Set http = HttpCall("PUT", baseUrl & "/api/profiles.asp?id=" & newId, cookieHeader, csrf, _
		"{" & q & "firstName" & q & ":" & q & "Updated" & q & "," & _
		q & "lastName" & q & ":" & q & "User" & q & "," & _
		q & "email" & q & ":" & q & "t@example.com" & q & "," & _
		q & "bio" & q & ":" & q & "hi" & q & "}")
	If Not (http Is Nothing) Then
		AssertEqual 200, http.status, "status"
		AssertTrue InStr(http.responseText, q & "firstName" & q & ":" & q & "Updated" & q) > 0, "update applied"
	End If

	BeginTest "api", "DELETE /api/profiles.asp removes the profile"
	Set http = HttpCall("DELETE", baseUrl & "/api/profiles.asp?id=" & newId, cookieHeader, csrf, "")
	If Not (http Is Nothing) Then
		AssertEqual 204, http.status, "status"
	End If

	BeginTest "api", "GET after delete returns 404"
	Set http = HttpCall("GET", baseUrl & "/api/profiles.asp?id=" & newId, cookieHeader, "", "")
	If Not (http Is Nothing) Then
		AssertEqual 404, http.status, "status"
	End If

	BeginTest "api", "DELETE /api/session.asp ends the session"
	Set http = HttpCall("DELETE", baseUrl & "/api/session.asp", cookieHeader, csrf, "")
	If Not (http Is Nothing) Then
		AssertEqual 200, http.status, "status"
		AssertTrue InStr(http.responseText, q & "active" & q & ":false") > 0, "session reports inactive"
	End If

	BeginTest "api", "session's accdb file is removed from disk after ending"
	Dim sessionIdFromCookie, accdbPath, fso2
	sessionIdFromCookie = Mid(cookieHeader, InStr(cookieHeader, "=") + 1)
	accdbPath = GetSessionsFolder() & "\" & sessionIdFromCookie & ".accdb"
	Set fso2 = Server.CreateObject("Scripting.FileSystemObject")
	AssertFalse fso2.FileExists(accdbPath), "session accdb should be deleted from disk"
End Sub
%>
