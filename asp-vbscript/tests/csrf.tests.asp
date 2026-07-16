<%
' Tests for the pure-logic half of lib/csrf.asp. IsCsrfValid()'s actual
' header-matching behavior depends on the real incoming request's headers,
' which can't be varied in-process within a single page execution - that's
' covered by the live HTTP round trip in tests/api.tests.asp instead.

Function IsHexString(ByVal s)
	Dim i, c
	IsHexString = True
	For i = 1 To Len(s)
		c = Mid(s, i, 1)
		If InStr("0123456789abcdef", c) = 0 Then
			IsHexString = False
			Exit Function
		End If
	Next
End Function

Sub RunCsrfTests()
	BeginTest "csrf", "NewCsrfToken returns a 32-char lowercase hex string"
	Dim token
	token = NewCsrfToken()
	AssertEqual 32, Len(token), "token length"
	AssertTrue IsHexString(token), "token should be lowercase hex only"

	BeginTest "csrf", "NewCsrfToken does not return the same value twice in a row"
	Dim token2
	token2 = NewCsrfToken()
	AssertTrue token <> token2, "two consecutive tokens should differ"
End Sub
%>
