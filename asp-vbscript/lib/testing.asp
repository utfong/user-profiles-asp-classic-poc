<%
' Minimal xUnit-style assertion/reporting framework for the tests/ suite.
' Not reentrant - designed for a single test-runner page (tests/run.asp)
' that calls each RunXTests() sub once per request.

Dim TestResults(), TestResultCount, CurrentSuiteName, CurrentTestName

Sub ResetTests()
	ReDim TestResults(-1)
	TestResultCount = 0
	CurrentSuiteName = ""
	CurrentTestName = ""
End Sub

Sub BeginTest(ByVal suiteName, ByVal testName)
	CurrentSuiteName = suiteName
	CurrentTestName = testName
End Sub

Sub RecordResult(ByVal passed, ByVal detail)
	Dim r
	Set r = Server.CreateObject("Scripting.Dictionary")
	r.Add "suite", CurrentSuiteName
	r.Add "test", CurrentTestName
	r.Add "passed", passed
	r.Add "detail", detail
	ReDim Preserve TestResults(TestResultCount)
	Set TestResults(TestResultCount) = r
	TestResultCount = TestResultCount + 1
End Sub

Sub AssertEqual(ByVal expected, ByVal actual, ByVal message)
	If CStr(expected) = CStr(actual) Then
		RecordResult True, message
	Else
		RecordResult False, message & " -- expected [" & CStr(expected) & "] but got [" & CStr(actual) & "]"
	End If
End Sub

Sub AssertTrue(ByVal condition, ByVal message)
	If condition Then
		RecordResult True, message
	Else
		RecordResult False, message & " -- expected True but got False"
	End If
End Sub

Sub AssertFalse(ByVal condition, ByVal message)
	AssertTrue Not condition, message
End Sub

Sub RenderTestReport()
	Dim i, r, passCount, failCount, testStatus

	passCount = 0
	failCount = 0
	For i = 0 To TestResultCount - 1
		If TestResults(i).Item("passed") Then
			passCount = passCount + 1
		Else
			failCount = failCount + 1
		End If
	Next

	If failCount = 0 Then
		testStatus = "pass"
	Else
		testStatus = "fail"
	End If

	' Headers first, before any Write, regardless of buffering settings.
	Response.ContentType = "text/html"
	Response.CharSet = "UTF-8"
	Response.AddHeader "X-Test-Status", testStatus
	Response.AddHeader "X-Test-Pass-Count", CStr(passCount)
	Response.AddHeader "X-Test-Fail-Count", CStr(failCount)

	Response.Write "<!DOCTYPE html><html><head><meta charset=""utf-8""><title>ASP VBScript Test Results</title>"
	Response.Write "<style>body{font-family:Consolas,Menlo,monospace;background:#0b0f19;color:#e4e7ec;padding:24px}"
	Response.Write ".pass{color:#4ade80}.fail{color:#f87171}h1{font-size:1.15rem}"
	Response.Write ".summary{font-weight:bold;margin-top:16px;font-size:1.1rem;padding:10px 14px;border-radius:6px;display:inline-block}"
	Response.Write ".summary.pass{background:#0f2a1c}.summary.fail{background:#3a1414}</style></head><body>"
	Response.Write "<h1>ASP VBScript Test Results</h1><pre>"

	For i = 0 To TestResultCount - 1
		Set r = TestResults(i)
		If r.Item("passed") Then
			Response.Write "<span class=""pass"">PASS</span>  "
		Else
			Response.Write "<span class=""fail"">FAIL</span>  "
		End If
		Response.Write "[" & Server.HTMLEncode(r.Item("suite")) & "] " & Server.HTMLEncode(r.Item("test")) & _
			" -- " & Server.HTMLEncode(r.Item("detail")) & vbCrLf
	Next

	Response.Write "</pre>"
	Response.Write "<div class=""summary " & testStatus & """>" & passCount & " passed, " & failCount & _
		" failed (" & TestResultCount & " total)</div>"
	Response.Write "</body></html>"
End Sub
%>
