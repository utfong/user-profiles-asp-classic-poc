<%
' Tests for lib/json.asp. Pure logic - runs entirely in-process, no HTTP
' round trip needed. Tricky strings (embedded quotes/backslashes) are built
' with Chr() rather than hand-doubled quote literals, to avoid transcription
' mistakes in the test itself.

Sub RunJsonTests()
	BeginTest "json", "JsonEncode escapes double quotes"
	Dim q, backslash
	q = Chr(34) ' "
	backslash = Chr(92) ' \
	' JSON escapes an embedded quote as \" (backslash+quote) - not to be
	' confused with VBScript's own "" doubling convention for writing a
	' literal quote inside a source-code string literal.
	AssertEqual q & "hello " & backslash & q & "world" & backslash & q & q, JsonEncode("hello " & q & "world" & q), "quotes should be escaped"

	BeginTest "json", "JsonEncode numbers are unquoted"
	AssertEqual "42", JsonEncode(42), "integer should not be quoted"

	BeginTest "json", "JsonEncode booleans"
	AssertEqual "true", JsonEncode(True), "True -> true"
	AssertEqual "false", JsonEncode(False), "False -> false"

	BeginTest "json", "JsonEncode null/empty"
	AssertEqual "null", JsonEncode(Null), "Null -> null"

	BeginTest "json", "JsonEncode empty array"
	Dim emptyArr()
	ReDim emptyArr(-1)
	AssertEqual "[]", JsonEncode(emptyArr), "empty array -> []"

	BeginTest "json", "JsonEncode array of dictionaries"
	Dim d1, d2, arr(1)
	Set d1 = Server.CreateObject("Scripting.Dictionary")
	d1.Add "a", 1
	Set d2 = Server.CreateObject("Scripting.Dictionary")
	d2.Add "a", 2
	Set arr(0) = d1
	Set arr(1) = d2
	AssertEqual "[{" & q & "a" & q & ":1},{" & q & "a" & q & ":2}]", JsonEncode(arr), "array of dicts"

	BeginTest "json", "JsonParse reads string/number/boolean/null fields"
	Dim src, parsed
	src = "{" & q & "firstName" & q & ":" & q & "Ada" & q & "," & _
		q & "age" & q & ":36," & _
		q & "active" & q & ":true," & _
		q & "note" & q & ":null}"
	Set parsed = JsonParse(src)
	AssertEqual "Ada", parsed.Item("firstName"), "string field"
	AssertEqual "36", parsed.Item("age"), "number field"
	AssertEqual "True", CStr(parsed.Item("active")), "boolean field"
	AssertTrue IsNull(parsed.Item("note")), "null field parses as Null"

	BeginTest "json", "JsonParse handles escaped newline and quote characters"
	Dim src2, parsed2
	src2 = "{" & q & "text" & q & ":" & q & "line1" & backslash & "nsay " & backslash & q & "hi" & backslash & q & q & "}"
	Set parsed2 = JsonParse(src2)
	AssertEqual "line1" & vbLf & "say " & q & "hi" & q, parsed2.Item("text"), "escaped newline and quotes"

	BeginTest "json", "JsonEncode/JsonParse round trip for a profile-shaped object"
	Dim d3, encoded, parsed3
	Set d3 = Server.CreateObject("Scripting.Dictionary")
	d3.Add "firstName", "Grace"
	d3.Add "lastName", "Hopper"
	d3.Add "email", "grace@example.com"
	encoded = JsonEncode(d3)
	Set parsed3 = JsonParse(encoded)
	AssertEqual "Grace", parsed3.Item("firstName"), "round trip firstName"
	AssertEqual "Hopper", parsed3.Item("lastName"), "round trip lastName"
	AssertEqual "grace@example.com", parsed3.Item("email"), "round trip email"
End Sub
%>
