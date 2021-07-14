// xll_csv.cpp - Parse CSV strings
#include "xll_csv.h"

using namespace xll;

#define XLTYPE_TOPIC "https://support.microsoft.com/en-us/office/type-function-45b4e688-4bc3-48b3-a105-ffa892995899"
#define X(a, b, c) XLL_CONST(LONG, XLTYPE_##b, a, c, "XLL", XLTYPE_TOPIC)
XLTYPE(X)
#undef XLTYPE_TOPIC
#undef X

using xcstr = xll::traits<XLOPERX>::xcstr;
using xchar = xll::traits<XLOPERX>::xchar;

AddIn xai_csv_parse(
	Function(XLL_LPOPER, "xll_csv_parse", "CSV.PARSE")
	.Arguments({
		Arg(XLL_LPOPER, "string", "is a string of comma separated values or handle."),
		Arg(XLL_CSTRING, "_rs", "is an optional record separator. Default is newline '\\n'."),
		Arg(XLL_CSTRING, "_fs", "is an optional field separator. Default is comma ','."),
		Arg(XLL_CSTRING, "_esc", "is an optional escape character. Default is backslash '\\'."),
		Arg(XLL_WORD, "_offset", "is an optional number of lines to skip. Default is 0."),
		Arg(XLL_WORD, "_count", "is an optional number of lines to return. Default is all.")
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
LPOPER WINAPI xll_csv_parse(LPOPER pcsv, xcstr _rs, xcstr _fs, xcstr _e, unsigned off, unsigned count)
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

			auto v = fms::view<char>(h_->buf, h_->len);
			o = xll::csv::parse<XLOPERX,char>(v, rs, fs, e, off, count);
		}
		else {
			ensure(pcsv->is_str());

			xchar rs = *_rs ? *_rs : '\n';
			xchar fs = *_fs ? *_fs : ',';
			xchar e = *_e ? *_e : '\\';

			auto v = fms::view<xchar>(pcsv->val.str + 1, pcsv->val.str[0]);
			o = xll::csv::parse<XLOPERX>(v, rs, fs, e, off, count);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = ErrNA;
	}

	return &o;
}

AddIn xai_range_convert(
	Function(XLL_LPOPER, "xll_range_convert", "RANGE.CONVERT")
	.Arguments({
		Arg(XLL_LPOPER, "range", "is a range or handle to a range."),
		Arg(XLL_LPOPER, "types", "is a one row range of conversion types."),
		Arg(XLL_BOOL, "_header", "is an optional boolean indicating the first row is a header.")
		})
	.FunctionHelp("Convert columns of range.")
	.Category("RANGE")
	.Documentation(R"(
Convert columns of <code>range</code> based on index. The <code>types</code>
should have the same size as the number of columns and contain numbers
corresponding to <code>xltype*</code> values.
)")
);
LPOPER WINAPI xll_range_convert(LPOPER prange, const LPOPER ptypes, BOOL header)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		LPOPER po = &o;
		if (prange->is_num()) {
			handle<OPER> h_(prange->val.num);
			ensure(h_);
			po = h_.ptr();
		}
		else {
			o = *prange;
		}
		ensure(po->columns() == ptypes->size());

		for (unsigned i = 0; i < po->columns(); ++i) {
			const auto& type = (ptypes)[i];
			if (type) {
				for (unsigned j = header; j < po->rows(); ++j) {
					parse::convert((*po)(i, j), type);
				}
			}
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return (LPOPER)&ErrValue;
	}

	return &o;
}

#ifdef _DEBUG

Auto<OpenAfter> xaoa_test_csv_parse([]() { return xll::csv::test<TCHAR>() == 0; });

#endif // _DEBUG
