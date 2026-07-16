<%
// JSON helpers for Classic ASP / JScript. No native JSON global exists in
// this engine (typeof JSON === "undefined" here, despite @_jscript_version
// reporting 11), so both directions are hand-rolled, same as asp-vbscript's
// lib/json.asp. Unlike that version, jsonParse is a full recursive-descent
// parser and objects/arrays nest freely: in JScript, explicitly *rejecting*
// nesting would take more code (an extra branch in parseValue that throws on
// "{"/"[") than just recursing into it, so there's no win in hand-limiting it
// to this app's own flat firstName/lastName/email/bio shape the way
// VBScript's Dictionary-based design was forced to.

function jsonEncode(value) {
	if (value === null || value === undefined) return "null";
	if (value instanceof Date) return '"' + jsonEscape(formatIso8601(value)) + '"';
	// ADO DATETIME columns come back from a Recordset as a raw COM VarDate,
	// not a JScript Date (typeof "date", fails instanceof Date, has no
	// methods of its own) - rowToObject() normalizes those to real Date
	// objects via `new Date(rawValue)` before they ever reach here, but this
	// branch is a defensive fallback in case a raw VarDate slips through.
	if (typeof value === "date") return '"' + jsonEscape(formatIso8601(new Date(value))) + '"';
	if (value instanceof Array) return jsonEncodeArray(value);
	var t = typeof value;
	if (t === "boolean") return value ? "true" : "false";
	if (t === "number") return isFinite(value) ? String(value) : "null";
	if (t === "string") return '"' + jsonEscape(value) + '"';
	if (t === "object") return jsonEncodeObject(value);
	return '"' + jsonEscape(String(value)) + '"';
}

function jsonEncodeArray(arr) {
	var parts = [];
	for (var i = 0; i < arr.length; i++) parts.push(jsonEncode(arr[i]));
	return "[" + parts.join(",") + "]";
}

function jsonEncodeObject(obj) {
	var parts = [];
	for (var key in obj) {
		if (!obj.hasOwnProperty(key)) continue;
		parts.push('"' + jsonEscape(key) + '":' + jsonEncode(obj[key]));
	}
	return "{" + parts.join(",") + "}";
}

function jsonEscape(s) {
	var result = "";
	for (var i = 0; i < s.length; i++) {
		var c = s.charAt(i);
		var code = s.charCodeAt(i);
		switch (c) {
			case "\\": result += "\\\\"; break;
			case '"': result += '\\"'; break;
			case "\b": result += "\\b"; break;
			case "\t": result += "\\t"; break;
			case "\n": result += "\\n"; break;
			case "\f": result += "\\f"; break;
			case "\r": result += "\\r"; break;
			default:
				if (code < 32) {
					var hex = code.toString(16);
					while (hex.length < 4) hex = "0" + hex;
					result += "\\u" + hex;
				} else {
					result += c;
				}
		}
	}
	return result;
}

function pad2(n) {
	return (n < 10 ? "0" : "") + n;
}

function formatIso8601(d) {
	return d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate()) + "T" +
		pad2(d.getHours()) + ":" + pad2(d.getMinutes()) + ":" + pad2(d.getSeconds());
}

// Recursive-descent parser. Parser state (pos/len) lives in closures local to
// each jsonParse() call, not a Class with private fields the way VBScript's
// ClsFlatJsonParser needed: VBScript has no closures, so parser state had to
// live in an object's fields; JScript's function-scoped closures make that
// object unnecessary.
function jsonParse(text) {
	var pos = 0;
	var len = text.length;

	function fail(msg) {
		throw new Error("jsonParse: " + msg + " at position " + pos);
	}

	function peek() {
		return pos < len ? text.charAt(pos) : "";
	}

	function skipWs() {
		while (pos < len) {
			var c = text.charAt(pos);
			if (c === " " || c === "\t" || c === "\r" || c === "\n") {
				pos++;
			} else {
				break;
			}
		}
	}

	function expect(ch) {
		if (peek() !== ch) fail("expected '" + ch + "'");
		pos++;
	}

	function isDigitChar(c) {
		return c >= "0" && c <= "9";
	}

	function parseValue() {
		skipWs();
		var c = peek();
		if (c === "{") return parseObject();
		if (c === "[") return parseArray();
		if (c === '"') return parseString();
		if (c === "t") return parseLiteral("true", true);
		if (c === "f") return parseLiteral("false", false);
		if (c === "n") return parseLiteral("null", null);
		if (c === "-" || isDigitChar(c)) return parseNumber();
		fail("unexpected character '" + c + "'");
	}

	function parseObject() {
		var obj = {};
		expect("{");
		skipWs();
		if (peek() === "}") {
			pos++;
			return obj;
		}
		while (true) {
			skipWs();
			var key = parseString();
			skipWs();
			expect(":");
			obj[key] = parseValue();
			skipWs();
			if (peek() === ",") {
				pos++;
				continue;
			}
			break;
		}
		skipWs();
		expect("}");
		return obj;
	}

	function parseArray() {
		var arr = [];
		expect("[");
		skipWs();
		if (peek() === "]") {
			pos++;
			return arr;
		}
		while (true) {
			arr.push(parseValue());
			skipWs();
			if (peek() === ",") {
				pos++;
				continue;
			}
			break;
		}
		skipWs();
		expect("]");
		return arr;
	}

	function parseLiteral(lit, val) {
		if (text.substr(pos, lit.length) !== lit) fail("invalid literal");
		pos += lit.length;
		return val;
	}

	function parseString() {
		expect('"');
		var result = "";
		while (peek() !== '"') {
			if (pos >= len) fail("unterminated string");
			var c = text.charAt(pos);
			if (c === "\\") {
				pos++;
				var esc = text.charAt(pos);
				switch (esc) {
					case '"': result += '"'; break;
					case "\\": result += "\\"; break;
					case "/": result += "/"; break;
					case "b": result += "\b"; break;
					case "f": result += "\f"; break;
					case "n": result += "\n"; break;
					case "r": result += "\r"; break;
					case "t": result += "\t"; break;
					case "u":
						var hexDigits = text.substr(pos + 1, 4);
						result += String.fromCharCode(parseInt(hexDigits, 16));
						pos += 4;
						break;
					default:
						fail("invalid escape sequence");
				}
				pos++;
			} else {
				result += c;
				pos++;
			}
		}
		pos++; // closing quote
		return result;
	}

	function parseNumber() {
		var startPos = pos;
		if (peek() === "-") pos++;
		while (pos < len && isDigitChar(text.charAt(pos))) pos++;
		if (peek() === ".") {
			pos++;
			while (pos < len && isDigitChar(text.charAt(pos))) pos++;
		}
		if (peek() === "e" || peek() === "E") {
			pos++;
			if (peek() === "+" || peek() === "-") pos++;
			while (pos < len && isDigitChar(text.charAt(pos))) pos++;
		}
		var numText = text.substring(startPos, pos);
		if (numText === "" || numText === "-") fail("invalid number at position " + startPos);
		return parseFloat(numText);
	}

	var result = parseValue();
	skipWs();
	if (pos !== len) fail("unexpected trailing characters");
	return result;
}
%>
