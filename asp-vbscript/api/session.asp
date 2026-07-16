<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include virtual="/lib/http.asp" -->
<!--#include virtual="/lib/json.asp" -->
<!--#include virtual="/lib/csrf.asp" -->
<!--#include virtual="/lib/db.asp" -->
<%
' Session lifecycle: GET status, POST provision the per-session .accdb,
' DELETE end the session early (also cleaned up automatically by
' Session_OnEnd in global.asa on timeout).

Const SESSION_TIMEOUT_MINUTES = 5 ' short on purpose so cleanup is easy to observe in a demo

Function StatusPayload()
	Dim d
	Set d = Server.CreateObject("Scripting.Dictionary")
	If HasActiveSession() Then
		d.Add "active", True
		d.Add "csrfToken", Session("CsrfToken")
		d.Add "timeoutMinutes", Session.Timeout
	Else
		d.Add "active", False
	End If
	Set StatusPayload = d
End Function

Dim method
method = GetRequestMethod()

Select Case method
	Case "GET"
		SendData 200, StatusPayload()

	Case "POST"
		If Not HasActiveSession() Then
			Dim dbPath
			On Error Resume Next
			Session.Timeout = SESSION_TIMEOUT_MINUTES
			dbPath = BuildSessionDbPath()
			CreateSessionDatabase dbPath
			If Err.Number <> 0 Then
				Dim errDesc
				errDesc = Err.Description
				Err.Clear
				On Error Goto 0
				DeleteSessionDatabase dbPath ' clean up any partially-created file
				SendError 500, "db_create_failed", _
					"Could not create the session database. Check that the Microsoft Access " & _
					"Database Engine (ACE OLEDB 12.0) is installed and its bitness matches the " & _
					"app pool. Detail: " & errDesc
			End If
			On Error Goto 0
			Session("DbPath") = dbPath
			Session("CsrfToken") = NewCsrfToken()
			SendData 201, StatusPayload()
		Else
			SendData 200, StatusPayload()
		End If

	Case "DELETE"
		RequireSession
		RequireCsrf
		DeleteSessionDatabase GetDbPath()
		Session("DbPath") = ""
		Session("CsrfToken") = ""
		Session.Abandon
		SendData 200, StatusPayload()

	Case Else
		Response.AddHeader "Allow", "GET, POST, DELETE"
		SendError 405, "method_not_allowed", "Supported methods: GET, POST, DELETE."
End Select
%>
