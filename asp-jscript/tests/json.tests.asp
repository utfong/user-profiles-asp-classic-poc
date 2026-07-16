<%
// Tests for lib/json.asp. Pure logic - runs entirely in-process, no HTTP
// round trip needed. Unlike asp-vbscript's version, tricky strings don't need
// building via Chr(34)/Chr(92): JScript string literals have real \" and \\
// escapes, so there's no VBScript-style "" quoting-convention collision to
// work around here.

function runJsonTests() {
	beginTest("json", "jsonEncode escapes double quotes and backslashes");
	assertEqual('"hello \\"world\\""', jsonEncode('hello "world"'), "quotes should be escaped");

	beginTest("json", "jsonEncode numbers are unquoted");
	assertEqual("42", jsonEncode(42), "integer should not be quoted");

	beginTest("json", "jsonEncode booleans");
	assertEqual("true", jsonEncode(true), "true -> true");
	assertEqual("false", jsonEncode(false), "false -> false");

	beginTest("json", "jsonEncode null/undefined");
	assertEqual("null", jsonEncode(null), "null -> null");
	assertEqual("null", jsonEncode(undefined), "undefined -> null");

	beginTest("json", "jsonEncode empty array");
	assertEqual("[]", jsonEncode([]), "empty array -> []");

	beginTest("json", "jsonEncode array of objects");
	assertEqual('[{"a":1},{"a":2}]', jsonEncode([{ a: 1 }, { a: 2 }]), "array of objects");

	beginTest("json", "jsonEncode nested objects/arrays (supported here, unlike asp-vbscript's flat-only parser)");
	var nested = { firstName: "Ada", tags: ["x", "y"], meta: { active: true, count: 2 } };
	assertEqual('{"firstName":"Ada","tags":["x","y"],"meta":{"active":true,"count":2}}', jsonEncode(nested), "nested encode");

	beginTest("json", "jsonParse reads string/number/boolean/null fields");
	var parsed = jsonParse('{"firstName":"Ada","age":36,"active":true,"note":null}');
	assertEqual("Ada", parsed.firstName, "string field");
	assertEqual("36", parsed.age, "number field");
	assertEqual("true", String(parsed.active), "boolean field");
	assertTrue(parsed.note === null, "null field parses as null");

	beginTest("json", "jsonParse handles escaped newline and quote characters");
	var parsed2 = jsonParse('{"text":"line1\\nsay \\"hi\\""}');
	assertEqual('line1\nsay "hi"', parsed2.text, "escaped newline and quotes");

	beginTest("json", "jsonParse handles nested objects and arrays");
	var parsed3 = jsonParse('{"firstName":"Ada","tags":["x","y"],"meta":{"active":true,"count":2}}');
	assertEqual("Ada", parsed3.firstName, "top-level string field");
	assertEqual("2", String(parsed3.tags.length), "nested array length");
	assertEqual("y", parsed3.tags[1], "nested array element");
	assertEqual("true", String(parsed3.meta.active), "nested object field (boolean)");
	assertEqual("2", String(parsed3.meta.count), "nested object field (number)");

	beginTest("json", "jsonEncode/jsonParse round trip for a profile-shaped object");
	var encoded = jsonEncode({ firstName: "Grace", lastName: "Hopper", email: "grace@example.com" });
	var parsed4 = jsonParse(encoded);
	assertEqual("Grace", parsed4.firstName, "round trip firstName");
	assertEqual("Hopper", parsed4.lastName, "round trip lastName");
	assertEqual("grace@example.com", parsed4.email, "round trip email");

	beginTest("json", "jsonEncode/jsonParse round trip for a Date value");
	var d = new Date(2026, 0, 15, 9, 30, 5); // Jan 15 2026, 09:30:05 (month is 0-based)
	var encodedDate = jsonEncode({ createdAt: d });
	assertEqual('{"createdAt":"2026-01-15T09:30:05"}', encodedDate, "Date encodes as ISO 8601");
	var parsedDate = jsonParse(encodedDate);
	assertEqual("2026-01-15T09:30:05", parsedDate.createdAt, "Date round-trips as an ISO string");
}
%>
