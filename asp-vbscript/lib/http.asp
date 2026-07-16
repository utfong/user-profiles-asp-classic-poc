<%
' Request/response helpers for the JSON API.
' Requires lib/json.asp to already be included (JsonEncode).

Function GetRequestMethod()
	GetRequestMethod = UCase(Request.ServerVariables("REQUEST_METHOD"))
End Function

' Classic ASP has no native JSON body parsing - Request.Form only understands
' multipart/form-data and x-www-form-urlencoded. Read the raw bytes ourselves
' and decode as UTF-8 text so JsonParse can work on it.
Function ReadRawBody()
	Dim totalBytes, bytes, stream, raw
	totalBytes = Request.TotalBytes
	If totalBytes > 0 Then
		bytes = Request.BinaryRead(totalBytes)
		Set stream = Server.CreateObject("ADODB.Stream")
		stream.Type = 1 ' adTypeBinary
		stream.Open
		stream.Write bytes
		stream.Position = 0
		stream.Type = 2 ' adTypeText
		stream.Charset = "utf-8"
		raw = stream.ReadText
		stream.Close
	Else
		raw = ""
	End If
	ReadRawBody = raw
End Function

Function StatusText(ByVal code)
	Select Case code
		Case 200 : StatusText = "200 OK"
		Case 201 : StatusText = "201 Created"
		Case 204 : StatusText = "204 No Content"
		Case 400 : StatusText = "400 Bad Request"
		Case 403 : StatusText = "403 Forbidden"
		Case 404 : StatusText = "404 Not Found"
		Case 405 : StatusText = "405 Method Not Allowed"
		Case 409 : StatusText = "409 Conflict"
		Case 500 : StatusText = "500 Internal Server Error"
		Case Else : StatusText = CStr(code) & " Error"
	End Select
End Function

' Writes { "data": <payload> } and ends the response.
Sub SendData(ByVal statusCode, ByVal payload)
	Dim envelope
	Set envelope = Server.CreateObject("Scripting.Dictionary")
	envelope.Add "data", payload
	SendJson statusCode, envelope
End Sub

' Writes { "error": { "code": ..., "message": ... } } and ends the response.
Sub SendError(ByVal statusCode, ByVal errCode, ByVal message)
	Dim errObj, envelope
	Set errObj = Server.CreateObject("Scripting.Dictionary")
	errObj.Add "code", errCode
	errObj.Add "message", message
	Set envelope = Server.CreateObject("Scripting.Dictionary")
	envelope.Add "error", errObj
	SendJson statusCode, envelope
End Sub

Sub SendJson(ByVal statusCode, ByVal payload)
	Response.Status = StatusText(statusCode)
	Response.ContentType = "application/json"
	Response.CharSet = "UTF-8"
	If statusCode <> 204 Then
		Response.Write JsonEncode(payload)
	End If
	Response.End
End Sub
%>
