// xll_csv.cpp - Parse CSV strings
#include "xll_csv.h"

using namespace xll;

using xcstr = xll::traits<XLOPERX>::xcstr;
using xchar = xll::traits<XLOPERX>::xchar;

AddIn xai_csv_parse(
	Function(XLL_LPOPER, "xll_csv_parse", "CSV.PARSE")
	.Arguments({
		Arg(XLL_LPOPER, "string", "is a string of comma separated values or handle."),
		Arg(XLL_CSTRING, "_rs", "is an optional record separator. Default is newline '\\n'."),
		Arg(XLL_CSTRING, "_fs", "is an optional field separator. Default is comma ','."),
		Arg(XLL_CSTRING, "_esc", "is an optional escape character. Default is backslash '\\'."),
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
LPOPER WINAPI xll_csv_parse(LPOPER pcsv, xcstr _rs, xcstr _fs, xcstr _e)
{
#pragma XLLEXPORT
	static OPER o;

	try {

		if (pcsv->is_num()) {
			handle<fms::view<char>> h_(pcsv->as_num());
			ensure(h_);

			char rs = static_cast<char>(*_rs ? *_rs : '\n');
			char fs = static_cast<char>(*_fs ? *_fs : ',');
			char e = static_cast<char>(*_e ? *_e : '\\');

			o = xll::csv::parse<XLOPERX, char>(h_->buf, h_->len, rs, fs, e);
		}
		else {
			ensure(pcsv->is_str());

			xchar rs = *_rs ? *_rs : '\n';
			xchar fs = *_fs ? *_fs : ',';
			xchar e = *_e ? *_e : '\\';

			o = xll::csv::parse<XLOPERX>(pcsv->val.str + 1, pcsv->val.str[0], rs, fs, e);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = ErrNA;
	}

	return &o;
}

#ifdef _DEBUG

Auto<OpenAfter> xaoa_test_csv_parse([]() { return xll::csv::test<TCHAR>() == 0; });

#endif // _DEBUG
