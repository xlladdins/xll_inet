// xll_csv.cpp - Parse CSV strings
#include "xll_csv.h"

using namespace xll;

AddIn xai_csv_parse(
	Function(XLL_LPOPER, "xll_csv_parse", "CSV.PARSE")
	.Arguments({
		Arg(XLL_CSTRING, "string", "is a string of comma separated values."),
		Arg(XLL_CSTRING, "_fs", "is the optional field separator. Default is ','."),
		Arg(XLL_CSTRING, "_rs", "is the optional record separator. Default is '\\n'."),
		Arg(XLL_CSTRING, "_esc", "is the optional escape character. Default is '\"'."),
		})
	.FunctionHelp("Parse CSV string into a range.")
	.Category("CSV")
	.Documentation(R"xyzyx(
Convert <code>string</code> to a range. It uses the Excel function
<a href="https://support.microsoft.com/en-us/office/value-function-257d0108-07dc-437d-ae1c-bc2d3953d8c2"<code>VALUE</code>
to parse anything that looks like a number, date, or time. The strings
<code>"TRUE"</code> and <code>"FALSE"</code> are converted to Boolean values.
)xyzyx")
);
LPOPER WINAPI xll_csv_parse(xcstr csv, xcstr fs, xcstr rs, xcstr esc)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		if (!*fs) {
			fs = _T(",");
		}
		if (!*rs) {
			rs = _T("\\n");
		}
		if (!*esc) {
			esc = _T("\"");
		}

		o = csv_parse(csv, *fs, rs, *esc);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = ErrNA;
	}

	return &o;
}

#ifdef _DEBUG

int test_csv_parse()
{
	try {
		xll_test_item();
		xll_test_csv_skip();
		xll_test_csv_parse_field();

		xcstr buf = _T("ab,cd\ne,fg");
		OPER o = csv_parse(buf);
		ensure(o.rows() == 2);
		ensure(o.columns() == 2);
		ensure(o(0, 0) == OPER(_T("ab")));
		ensure(o(1, 1) == OPER(_T("fg")));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}

	return TRUE;
}
Auto<OpenAfter> xaoa_test_csv_parse(test_csv_parse);

#endif // _DEBUG
