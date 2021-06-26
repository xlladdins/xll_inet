#include "xll_parse_json.h"

using namespace xll;

using namespace xll;

using xcstr = xll::traits<XLOPERX>::xcstr;
using xchar = xll::traits<XLOPERX>::xchar;

AddIn xai_parse_json(
	Function(XLL_LPOPER, "xll_parse_json", "JSON.PARSE")
	.Arguments({
		Arg(XLL_PSTRING, "string", "is a JSON string."),
		})
	.FunctionHelp("Parse JSON string into a range.")
	.Category("CSV")
	.HelpTopic("https://www.json.org/json-en.html")
	.Documentation(xll_parse_json_doc)
	.Documentation(R"(
<p>
An multi data type is a two dimension range of <code>OPER</code>s. Each
element of the range can be another multi, but Excel has no way of
displying these. 
</p>
)")
);
LPOPER WINAPI xll_parse_json(xcstr csv)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		o = xll::parse::json::object<XLOPERX>(view(csv + 1, csv[0]));
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
		Arg(XLL_PSTRING, "index", "is a range of indices to return."),
		})
		.FunctionHelp("Parse JSON string into a range.")
	.Category("CSV")
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
		XLOPERX x;
		if (pjson->is_num()) {
			handle<OPER> h_(pjson->val.num);
			ensure(h_);
			ensure(h_->rows() == 2);
			x = *h_.ptr(); // current JSON node
		}
		else {
			x = *pjson;
		}

		for (const OPER& i : *pindex) {
			DWORD match;
			if (i.is_num()) {
				match = (DWORD)i.val.num - 1; // 1-base index
			}
			else {
				ensure(x.val.array.rows == 2);
				x.val.array.rows = 1; // only keys
				OPER mi = Excel(xlfMatch, i, x, OPER(0)); // exact
				x.val.array.rows = 2;
				ensure(mi);
				match = (DWORD)mi.val.num - 1;
			}
			x = index(x, 1, match);
		}

		o = x;
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