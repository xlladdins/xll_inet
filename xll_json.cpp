#ifdef _DEBUG
#include <cassert>
#endif
#include "xll_json.h"

//using namespace fms;
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
		switch(json::type(*pjson)) {
		case json::TYPE::Object:
			o = "object";
			break;
		case json::TYPE::Array:
			o = "array";
			break;
		case json::TYPE::String:
			o = "string";
			break;
		case json::TYPE::Number:
			o = "number";
			break;
		case json::TYPE::Boolean:
			o = "boolean";
			break;
		case json::TYPE::Null:
			o = "null";
			break;
		default:
			o = ErrValue;
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}

#ifdef _DEBUG

Auto<OpenAfter> xaoa_json_type_test([]() {
	try {
		{
			OPER o = json::object("key", "val");
			ensure(*xll_json_type(&o) == "object");
			o = json::object();
			ensure(*xll_json_type(&o) == "object");
		}
		{
			OPER o = json::array("i0", 1, true);
			ensure(*xll_json_type(&o) == "array");
			o = json::array();
			ensure(*xll_json_type(&o) == "array");
		}
		{
			OPER o = json::string("foo");
			ensure(*xll_json_type(&o) == "string");
			o = json::string("");
			ensure(*xll_json_type(&o) == "string");
			o = json::string(OPER{});
			ensure(o == "\"\"");
			ensure(*xll_json_type(&o) == "string");
			o = "bar";
			ensure(*xll_json_type(&o) == "string");
		}
		{
			OPER o(1.23);
			ensure(*xll_json_type(&o) == "number");
		}
		{
			OPER o(true);
			ensure(*xll_json_type(&o) == "boolean");
		}
		{
			ensure(*xll_json_type((LPOPER)&ErrNull) == "null");
		}
		{
			OPER o(REF(2, 3));
			ensure(*xll_json_type(&o) == ErrValue);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}

	return TRUE;
});

#endif // _DEBUG

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
			handle<fms::char_view<char>> h_(pjson->val.num);
			ensure(h_);
			// convert from char to wchar if needed
			o = json::parse::view<XLOPERX, char>(*h_);
		}
		else {
			ensure(pjson->is_str());
			o = json::parse::view<XLOPERX>(fms::char_view<const TCHAR>(pjson->val.str + 1, pjson->val.str[0]));
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}

AddIn xai_json_stringify(
	Function(XLL_LPOPER, "xll_json_stringify", "JSON.STRINGIFY")
	.Arguments({
		Arg(XLL_LPOPER, "value", "is a JSON value or handle."),
		})
	.FunctionHelp("Convert JSON value to a string.")
	.Category("JSON")
	.HelpTopic("https://www.json.org/json-en.html")
	.Documentation(xll_parse_json_doc)
);
LPOPER WINAPI xll_json_stringify(LPOPER pjson)
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
		auto str = json::stringify<XLOPERX, TCHAR>(*pjson);
		o = OPER(str.c_str(), static_cast<TCHAR>(str.length()));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}

// match * and ? wildcards
static inline bool glob(const XLOPERX& o, const XLOPERX& pat)
{
	// phony up a 1x1 multi
	static XLOPERX x = { 
		.val = {.array = {.rows = 1, .columns = 1}},
		.xltype = xltypeMulti
	};
	static XLOPERX zero = { .val = {.num = 0}, .xltype = xltypeNum };

	x.val.array.lparray = const_cast<XLOPERX*>(&o);

	return Excel(xlfMatch, pat, x, zero).is_num();
}

// all subkeys of object matching pattern separated by "."
inline OPER match_keys(const OPER& v, const OPER& pat, OPER& o)
{
	static OPER dot(".");
	static OPER z("0");

	if (v.is_multi()) {
		OPER pre;
		if (auto n = o.size()) {
			pre = o[n - 1];
		}

		if (json::type(v) == json::TYPE::Object) {
			for (unsigned i = 0; i < v.columns(); ++i) {
				if (glob(v(0, i), pat)) {
					o.push_back(pre & dot & v(0, i));
					if (v(1, i).is_multi()) {
						match_keys(v(1, i), pat, o);
					}
				}
			}
		}
		else if (json::type(v) == json::TYPE::Array) {
			for (unsigned i = 0; i < v.size(); ++i) {
				auto oi = Excel(xlfText, OPER(i), z);
				if (glob(oi, pat)) { 
					o.push_back(pre & dot & oi);
					if (v[i].is_multi()) {
						match_keys(v[i], pat, o);
					}
				}
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

		o = OPER{};
		o = match_keys(*pjson, *pkeys, o);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = ErrNA;
	}

	return &o;
}

AddIn xai_json_index(
	Function(XLL_LPOPER, "xll_json_index", "JSON.VALUE")
	.Arguments({
		Arg(XLL_LPOPER, "object", "is a JSON object or handle."),
		Arg(XLL_LPOPER, "key", "is the key to lookup."),
		})
	.Uncalced()
	.FunctionHelp("Return value corresponding to key.")
	.Category("JSON")
	.HelpTopic("https://www.json.org/json-en.html")
	.Documentation(R"xyzyx(
Return the value in <code>object</code> determined by the <code>key</code>.
If the value is an object or array then a handle is returned.
If the <code>key</code> is a string starting with a period (<code>.</code>)
then it is interpreted as a <code>jq</code> style dotted key.
If <code>key</code> is an array of strings the
lookup recurses along the specified keys in order.
<p>
If the value 
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

		if (pindex->is_str() and pindex->val.str[0] and pindex->val.str[1] == _T('.')) {
			OPER i = json::to_index(*pindex);
			o = json::index(*pjson, i);
		}
		else {
			o = json::index(*pjson, *pindex);
		}

		if (o.is_multi()) {
			handle<OPER> h_(new OPER(o));
			o = h_.get();
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}


#ifdef _DEBUG

Auto<OpenAfter> xaoa_json_test([]() {
	try {
		//_crtBreakAlloc = 52474;
		ensure(0 == xll::json::test());
		//ensure(0 == xll::json::parse::test());
		ensure(0 == xll::json::index_test<XLOPERX>());
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}

	return TRUE;
});

#endif // _DEBUG