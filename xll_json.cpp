#include "xll_json.h"

using namespace fms;
using namespace xll;

using xcstr = xll::traits<XLOPERX>::xcstr;
using xchar = xll::traits<XLOPERX>::xchar;

AddIn xai_json_type(
	Function(XLL_LPOPER, "xll_json_type", "JSON.TYPE")
	.Arguments({
		Arg(XLL_LPOPER, "value", "is a JSON value or handle."),
		})
	.FunctionHelp("Return the type of a JSON value.")
	.Category("JSON")
	.HelpTopic("https://www.json.org/json-en.html")
	.Documentation(R"(
Return a string indicating if the JSON value is an <code>object</code>, <code>array</code>,
<code>string</code>, <code>number</code>, <code>boolean</code>, or <code>null</code>.
)")
);
LPOPER WINAPI xll_json_type(LPOPER pjson)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		o = ErrNA;

		if (pjson->is_num()) {
			handle<OPER> h_(pjson->val.num);
			if (h_) {
				pjson = h_.ptr();
			}
		}

		if (pjson->is_multi()) {
			if (pjson->rows() == 2) {
				o = "object";
			}
			else if (pjson->rows() == 1) {
				o = "array";
			}
		}
		else {
			if (pjson->is_str()) {
				o = "string";
			}
			else if (pjson->is_num()) {
				o = "num";
			}
			else if (pjson->is_bool()) {
				o = "boolean";
			}
			else if (*pjson == ErrNull) {
				o = "null";
			}
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}

AddIn xai_parse_json(
	Function(XLL_LPOPER, "xll_parse_json", "JSON.PARSE")
	.Arguments({
		Arg(XLL_LPOPER, "json", "is a JSON string or handle."),
		})
	.FunctionHelp("Return a parsed JSON value.")
	.Category("JSON")
	.HelpTopic("https://www.json.org/json-en.html")
	.Documentation(xll_parse_json_doc)
);
LPOPER WINAPI xll_parse_json(LPOPER pjson)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		o = ErrNA;
		if (pjson->is_num()) {
			handle<fms::view<char>> h_(pjson->val.num);
			ensure(h_);
			// convert from char to wchar if needed
			o = json::parse::view<XLOPERX, char>(fms::view<const char>(h_->buf, h_->len));
		}
		else {
			ensure(pjson->is_str());
			o = json::parse::view<XLOPERX>(fms::view<const TCHAR>(pjson->val.str + 1, pjson->val.str[0]));
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}

// all subkeys of object matching pattern seperated by "."
inline OPER match_keys(const OPER& v, const OPER& pattern, const OPER& prefix = OPER("."))
{
	OPER o;

	ensure(v.rows() == 2);

	XLOPERX vi;
	vi.xltype = xltypeMulti;
	vi.val.array.rows = 1;
	vi.val.array.columns = 1;
	for (unsigned i = 0; i < v.columns(); ++i) {
		vi.val.array.lparray = v.val.array.lparray + i;
		if (Excel(xlfMatch, pattern, vi, OPER(0))) {
			o.push_back(prefix & v(0, i));
			if (v(1, i).rows() == 2) { // object
				o.push_back(match_keys(v(1, i), pattern, prefix & v(0, 1) & OPER(".")));
			}
		}
	}

	return o;
}

AddIn xai_json_keys(
	Function(XLL_LPOPER, "xll_json_keys", "JSON.KEYS")
	.Arguments({
		Arg(XLL_LPOPER, "json", "is a JSON range or handle."),
		Arg(XLL_LPOPER, "_pattern", "is a regular expression to match keys. Default is \"*\"."),
		})
	.FunctionHelp("Parse JSON string into a range.")
	.Category("JSON")
	.HelpTopic("https://www.json.org/json-en.html")
	.Documentation(R"xyzyx(
Return all keys in JSON object matching <code>_pattern</code> using <code>jq</code> dotted naming.
The pattern can include wildcard characters <code>?</code> to match any single character
or <code>*</code> to match zero or more characters.
)xyzyx")
);
LPOPER WINAPI xll_json_keys(LPOPER pjson, LPOPER pkeys)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		if (pkeys->is_missing()) {
			static OPER star("*");
			pkeys = &star;
		}

		if (pjson->is_num()) {
			handle<OPER> h_(pjson->val.num);
			ensure(h_);

			pjson = h_.ptr();
		}

		o = match_keys(*pjson, *pkeys, OPER("."));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = ErrNA;
	}

	return &o;
}


AddIn xai_json_index(
	Function(XLL_LPOPER, "xll_json_index", "JSON.INDEX")
	.Arguments({
		Arg(XLL_LPOPER, "json", "is a JSON range or handle."),
		Arg(XLL_LPOPER, "key", "is the key to lookup."),
		})
	.FunctionHelp("Parse JSON string into a range.")
	.Category("JSON")
	.HelpTopic("https://www.json.org/json-en.html")
	.Documentation(R"xyzyx(
Lookup multi-level index in JSON range using <code>key</code>s. Only works at second and greater
depths if <code>json</code> is a handle to a range.
)xyzyx")
);
LPOPER WINAPI xll_json_index(LPOPER pjson, LPOPER pindex)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		o = ErrNA;
		if (pjson->is_num()) {
			handle<OPER> h_(pjson->val.num);
			ensure(h_);

			pjson = h_.ptr();
		}

		o = json::index(*pjson, *pindex);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}

#ifdef _DEBUG

Auto<OpenAfter> aoa_test_parse_json([]() { return xll::json::parse::test(); });

#endif // _DEBUG