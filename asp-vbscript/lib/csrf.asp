<%
' CSRF synchronizer-token helpers.
' Note: token is generated with Rnd(), not a CSPRNG - classic ASP/VBScript has
' no built-in crypto RNG without an extra COM component. Fine for a POC; a
' real deployment should source this from something like CAPICOM or a native
' helper DLL instead.

Function NewCsrfToken()
	Dim i, token
	Randomize
	token = ""
	For i = 1 To 32
		token = token & Hex(Int(Rnd() * 16))
	Next
	NewCsrfToken = LCase(token)
End Function

Function IsCsrfValid()
	Dim headerToken, sessionToken
	headerToken = Request.ServerVariables("HTTP_X_CSRF_TOKEN")
	sessionToken = ""
	If Not IsEmpty(Session("CsrfToken")) Then
		sessionToken = Session("CsrfToken")
	End If
	IsCsrfValid = (sessionToken <> "" And headerToken <> "" And headerToken = sessionToken)
End Function

' Guard for write endpoints: ends the response with 403 if the token is
' missing or doesn't match. Requires lib/http.asp (SendError) to be included.
Sub RequireCsrf()
	If Not IsCsrfValid() Then
		SendError 403, "invalid_csrf", "Missing or invalid CSRF token."
	End If
End Sub
%>
