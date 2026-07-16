<%
' JSON helpers for Classic ASP / VBScript.
' JsonEncode: Dictionary -> object, VBScript Array -> array, scalars as expected.
' JsonParse: flat objects only (string/number/boolean/null values) - matches the
' shape of the profile create/update payloads this app actually sends.

Function JsonEncode(ByVal val)
	If IsArray(val) Then
		JsonEncode = JsonEncodeArray(val)
	ElseIf IsObject(val) Then
		If val Is Nothing Then
			JsonEncode = "null"
		ElseIf TypeName(val) = "Dictionary" Then
			JsonEncode = JsonEncodeObject(val)
		Else
			Err.Raise vbObjectError + 513, "JsonEncode", "Unsupported object type: " & TypeName(val)
		End If
	ElseIf IsNull(val) Or IsEmpty(val) Then
		JsonEncode = "null"
	ElseIf VarType(val) = vbBoolean Then
		If val Then
			JsonEncode = "true"
		Else
			JsonEncode = "false"
		End If
	ElseIf VarType(val) = vbString Then
		JsonEncode = """" & JsonEscape(val) & """"
	ElseIf IsDate(val) Then
		JsonEncode = """" & JsonEscape(FormatIso8601(CDate(val))) & """"
	ElseIf IsNumeric(val) Then
		JsonEncode = CStr(val)
	Else
		JsonEncode = """" & JsonEscape(CStr(val)) & """"
	End If
End Function

Function JsonEncodeArray(ByVal arr)
	Dim sb, i, isFirst
	sb = "["
	isFirst = True
	For i = LBound(arr) To UBound(arr)
		If Not isFirst Then sb = sb & ","
		sb = sb & JsonEncode(arr(i))
		isFirst = False
	Next
	JsonEncodeArray = sb & "]"
End Function

Function JsonEncodeObject(ByVal dict)
	Dim sb, k, isFirst
	sb = "{"
	isFirst = True
	For Each k In dict.Keys
		If Not isFirst Then sb = sb & ","
		sb = sb & """" & JsonEscape(CStr(k)) & """:" & JsonEncode(dict.Item(k))
		isFirst = False
	Next
	JsonEncodeObject = sb & "}"
End Function

Function JsonEscape(ByVal s)
	Dim result, i, c, code
	result = ""
	For i = 1 To Len(s)
		c = Mid(s, i, 1)
		Select Case c
			Case "\"
				result = result & "\\"
			Case """"
				result = result & "\"""
			Case Chr(8)
				result = result & "\b"
			Case Chr(9)
				result = result & "\t"
			Case Chr(10)
				result = result & "\n"
			Case Chr(12)
				result = result & "\f"
			Case Chr(13)
				result = result & "\r"
			Case Else
				code = AscW(c)
				If code < 32 Then
					result = result & "\u" & Right("0000" & Hex(code), 4)
				Else
					result = result & c
				End If
		End Select
	Next
	JsonEscape = result
End Function

Function FormatIso8601(ByVal d)
	FormatIso8601 = Year(d) & "-" & Right("0" & Month(d), 2) & "-" & Right("0" & Day(d), 2) & "T" & _
		Right("0" & Hour(d), 2) & ":" & Right("0" & Minute(d), 2) & ":" & Right("0" & Second(d), 2)
End Function

' --- Parsing (flat objects only) ---

Class ClsFlatJsonParser
	Private m_text, m_pos, m_len

	Public Function ParseObject(ByVal jsonText)
		Dim dict, key
		Set dict = Server.CreateObject("Scripting.Dictionary")
		m_text = jsonText
		m_pos = 1
		m_len = Len(m_text)
		SkipWs
		Expect "{"
		SkipWs
		If Peek() = "}" Then
			m_pos = m_pos + 1
		Else
			Do
				SkipWs
				key = ParseString()
				SkipWs
				Expect ":"
				SkipWs
				dict.Add key, ParseScalar()
				SkipWs
				If Peek() = "," Then
					m_pos = m_pos + 1
				Else
					Exit Do
				End If
			Loop
			SkipWs
			Expect "}"
		End If
		Set ParseObject = dict
	End Function

	Private Sub Expect(ByVal ch)
		If Peek() <> ch Then
			Err.Raise vbObjectError + 514, "JsonParse", "Expected '" & ch & "' at position " & m_pos
		End If
		m_pos = m_pos + 1
	End Sub

	Private Function Peek()
		If m_pos <= m_len Then
			Peek = Mid(m_text, m_pos, 1)
		Else
			Peek = ""
		End If
	End Function

	Private Sub SkipWs
		Do While m_pos <= m_len And ( _
			Mid(m_text, m_pos, 1) = " " Or Mid(m_text, m_pos, 1) = vbTab Or _
			Mid(m_text, m_pos, 1) = vbCr Or Mid(m_text, m_pos, 1) = vbLf)
			m_pos = m_pos + 1
		Loop
	End Sub

	Private Function ParseScalar()
		Dim c
		c = Peek()
		If c = """" Then
			ParseScalar = ParseString()
		ElseIf c = "t" Then
			ExpectLiteral "true"
			ParseScalar = True
		ElseIf c = "f" Then
			ExpectLiteral "false"
			ParseScalar = False
		ElseIf c = "n" Then
			ExpectLiteral "null"
			ParseScalar = Null
		ElseIf c = "{" Or c = "[" Then
			Err.Raise vbObjectError + 515, "JsonParse", "Nested objects/arrays are not supported"
		Else
			ParseScalar = ParseNumber()
		End If
	End Function

	Private Sub ExpectLiteral(ByVal lit)
		If Mid(m_text, m_pos, Len(lit)) <> lit Then
			Err.Raise vbObjectError + 516, "JsonParse", "Invalid literal at position " & m_pos
		End If
		m_pos = m_pos + Len(lit)
	End Sub

	Private Function ParseString()
		Dim result, c, esc, hexDigits
		Expect """"
		result = ""
		Do While Peek() <> """"
			If m_pos > m_len Then
				Err.Raise vbObjectError + 517, "JsonParse", "Unterminated string"
			End If
			c = Mid(m_text, m_pos, 1)
			If c = "\" Then
				m_pos = m_pos + 1
				esc = Mid(m_text, m_pos, 1)
				Select Case esc
					Case """" : result = result & """"
					Case "\" : result = result & "\"
					Case "/" : result = result & "/"
					Case "b" : result = result & Chr(8)
					Case "f" : result = result & Chr(12)
					Case "n" : result = result & vbLf
					Case "r" : result = result & vbCr
					Case "t" : result = result & vbTab
					Case "u"
						hexDigits = Mid(m_text, m_pos + 1, 4)
						result = result & ChrW(CLng("&H" & hexDigits))
						m_pos = m_pos + 4
					Case Else
						Err.Raise vbObjectError + 518, "JsonParse", "Invalid escape sequence"
				End Select
				m_pos = m_pos + 1
			Else
				result = result & c
				m_pos = m_pos + 1
			End If
		Loop
		m_pos = m_pos + 1 ' closing quote
		ParseString = result
	End Function

	Private Function ParseNumber()
		Dim startPos, numText
		startPos = m_pos
		If Peek() = "-" Then m_pos = m_pos + 1
		Do While m_pos <= m_len And IsDigitChar(Mid(m_text, m_pos, 1))
			m_pos = m_pos + 1
		Loop
		If Peek() = "." Then
			m_pos = m_pos + 1
			Do While m_pos <= m_len And IsDigitChar(Mid(m_text, m_pos, 1))
				m_pos = m_pos + 1
			Loop
		End If
		If Peek() = "e" Or Peek() = "E" Then
			m_pos = m_pos + 1
			If Peek() = "+" Or Peek() = "-" Then m_pos = m_pos + 1
			Do While m_pos <= m_len And IsDigitChar(Mid(m_text, m_pos, 1))
				m_pos = m_pos + 1
			Loop
		End If
		numText = Mid(m_text, startPos, m_pos - startPos)
		If numText = "" Or numText = "-" Then
			Err.Raise vbObjectError + 519, "JsonParse", "Invalid number at position " & startPos
		End If
		ParseNumber = CDbl(numText)
	End Function

	Private Function IsDigitChar(ByVal c)
		IsDigitChar = (c >= "0" And c <= "9")
	End Function
End Class

Function JsonParse(ByVal jsonText)
	Dim parser
	Set parser = New ClsFlatJsonParser
	Set JsonParse = parser.ParseObject(jsonText)
End Function
%>
