// xll_csv.cpp - Parse CSV strings
#include "xll_csv.h"

using namespace xll;

AddIn xai_csv_parse(
	Function(XLL_LPOPER, "xll_csv_parse", "CSV.PARSE")
	.Arguments({
		Arg(XLL_CSTRING, "string", "is a string of comma separated values."),
		Arg(XLL_CSTRING, "_fs", "is the optional field separator. Default is ','."),
		Arg(XLL_CSTRING, "_rs", "is the optional record separator. Default is ';'."),
		Arg(XLL_CSTRING, "_esc", "is the optional escape character. Default is '\"'."),
		})
	.FunctionHelp("Parse CSV string into a range.")
	.Category("CSV")
	.Documentation(R"()")
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
			rs = _T(";");
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
		xll_test_csv_field_parse();

		xcstr buf = _T("ab,cd\r\ne,fg");
		OPER o = csv_parse(buf);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}

	return TRUE;
}
Auto<OpenAfter> xaoa_test_csv_parse(test_csv_parse);

#endif // _DEBUG
