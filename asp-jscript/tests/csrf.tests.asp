<%
// Tests for the pure-logic half of lib/csrf.asp. isCsrfValid()'s actual
// header-matching behavior depends on the real incoming request's headers,
// which can't be varied in-process within a single page execution - that's
// covered by the live HTTP round trip in tests/api.tests.asp instead.

function isHexString(s) {
	for (var i = 0; i < s.length; i++) {
		if ("0123456789abcdef".indexOf(s.charAt(i)) === -1) return false;
	}
	return true;
}

function runCsrfTests() {
	beginTest("csrf", "newCsrfToken returns a 32-char lowercase hex string");
	var token = newCsrfToken();
	assertEqual(32, token.length, "token length");
	assertTrue(isHexString(token), "token should be lowercase hex only");

	beginTest("csrf", "newCsrfToken does not return the same value twice in a row");
	var token2 = newCsrfToken();
	assertTrue(token !== token2, "two consecutive tokens should differ");
}
%>
