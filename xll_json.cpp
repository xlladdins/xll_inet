#include "xll_parse_json.h"

using namespace fms;
using namespace xll;

using xcstr = xll::traits<XLOPERX>::xcstr;
using xchar = xll::traits<XLOPERX>::xchar;

AddIn xai_parse_json(
	Function(XLL_HANDLEX, "xll_parse_json", "\\JSON.PARSE")
	.Arguments({
		Arg(XLL_LPOPER, "json", "is a JSON string or handle."),
		})
	.Uncalced()
	.FunctionHelp("Return a handle to a parsed JSON string.")
	.Category("JSON")
	.HelpTopic("https://www.json.org/json-en.html")
	.Documentation(xll_parse_json_doc)
	.Documentation(R"(
<p>
A multi data type is a two dimension range of <code>OPER</code>s. Each
element of the range can be another multi, but Excel has no way of
displying these. 
</p>
)")
);
HANDLEX WINAPI xll_parse_json(LPOPER pjson)
{
#pragma XLLEXPORT
	HANDLEX h = INVALID_HANDLEX;

	try {
		if (pjson->is_num()) {
			handle<fms::view<char>> h_(pjson->val.num);
			ensure(h_);
			auto v = fms::view<const char>(h_->buf, h_->len);
			// convert from char to wchar if needed
			handle<OPER> j_(new OPER(xll::parse::json::value<XLOPERX, char>(v)));

			h = h_.get();
		}
		else {
			ensure(pjson->is_str());
			handle<OPER> h_(new OPER(xll::parse::json::value<XLOPERX>(fms::view<const TCHAR>(pjson->val.str + 1, pjson->val.str[0]))));

			h = h_.get();
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return h;
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
Return all keys in JSON object matching <code>_pattern</code>.
The pattern can include wildcard characters <code>?</code> to match any single character
or <code>*</code> to match zero or more characters.
)xyzyx")
);
LPOPER WINAPI xll_json_keys(LPOPER pjson, LPOPER pkeys)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		o = OPER{};
		if (pjson->is_num()) {
			handle<OPER> h_(pjson->val.num);
			ensure(h_);

			pjson = h_.ptr();
		}
		ensure(pjson->rows() == 2);
		const OPER& json = *pjson;
		if (pkeys->is_missing()) {
			static OPER star("*");
			pkeys = &star;
		}
		for (unsigned i = 0; i < json.columns(); ++i) {
			if (Excel(xlfMatch, *pkeys, json(0, i), OPER(0))) {
				o.push_back(json(0, i));
			}
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = ErrNA;
	}

	return &o;
}
// JSON.VALUE(index, ...) values matching index

AddIn xai_json_index(
	Function(XLL_LPOPER, "xll_json_index", "JSON.VALUE")
	.Arguments({
		Arg(XLL_LPOPER, "json", "is a JSON range or handle."),
		Arg(XLL_LPOPER, "key", "is the key to lookup."),
		})
	.FunctionHelp("Parse JSON string into a range.")
	.Category("JSON")
	.HelpTopic("https://www.json.org/json-en.html")
	.Documentation(R"xyzyx(
Lookup multi-level index in JSON range. Only works at second and greater
depths if <code>json</code> is a handle returned by <code>\\JSON.PARSE</code>.
)xyzyx")
);
LPOPER WINAPI xll_json_index(LPOPER pjson, LPOPER pindex)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		if (pjson->is_num()) {
			handle<OPER> h_(pjson->val.num);
			ensure(h_);

			pjson = h_.ptr();
		}

		o = parse::json::index(*pjson, *pindex);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = ErrNA;
	}

	return &o;
}

#ifdef _DEBUG

Auto<OpenAfter> aoa_test_parse_json([]() { return parse::json::test<XLOPERX>(); });

#endif // _DEBUG