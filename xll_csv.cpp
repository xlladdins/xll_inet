// xll_csv.cpp - Parse CSV strings
#include "xll_parse_csv.h"

using namespace xll;

using xcstr = xll::traits<XLOPERX>::xcstr;
using xchar = xll::traits<XLOPERX>::xchar;

AddIn xai_csv_parse(
	Function(XLL_LPOPER, "xll_csv_parse", "CSV.PARSE")
	.Arguments({
		Arg(XLL_PSTRING, "string", "is a string of comma separated values."),
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
LPOPER WINAPI xll_csv_parse(xcstr csv)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		xchar rs = '\n';
		xchar fs = ',';
		xchar l = '\"';
		xchar r = '\"';
		xchar e = '\\';
		o = xll::parse::csv::parse<XLOPERX>(csv + 1, csv[0], rs, fs, l, r, e);
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
		xll::parse::csv::test<TCHAR>();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}

	return TRUE;
}
Auto<OpenAfter> xaoa_test_csv_parse(test_csv_parse);

#endif // _DEBUG
