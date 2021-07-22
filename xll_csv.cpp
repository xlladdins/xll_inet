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

AddIn xai_xltype(
	Function(XLL_LONG, "xll_xltype", "XLTYPE")
	.Arguments({
		{XLL_LPOPER, "cell", "is a cell."},
		})
	.Category("XLL")
	.FunctionHelp("Returns the type of a cell as a XLTYPE_* enumeration.")
	.Documentation(R"(
Unrecognized types are returned as 0.
)")
);
LONG WINAPI xll_xltype(const LPOPER po)
{
#pragma XLLEXPORT

	return po->xltype;
}

AddIn xai_xltype_convert(
	Function(XLL_LPOPER, "xll_xltype_convert", "XLTYPE.CONVERT")
	.Arguments({
		Arg(XLL_LPOPER, "cell", "is a cell to convert."),
		Arg(XLL_LONG, "type", "is a type from the XLTYPE_* enumeration.")
		})
	.Category("XLL")
	.FunctionHelp("Convert cell to type from XLTYPE_* enumeration.")
	.Documentation(R"(
If a conversion fails <code>#VALUE!</code> is returned.
)")
);
LPOPER WINAPI xll_xltype_convert(LPOPER po, LONG type)
{
#pragma XLLEXPORT
	static OPER o;

	o = *po;
	parse::convert(o, type);

	return &o;
}

AddIn xai_csv_parse(
	Function(XLL_LPOPER, "xll_csv_parse", "CSV.PARSE")
	.Arguments({
		Arg(XLL_LPOPER, "csv", "is a string of comma separated values or handle."),
		Arg(XLL_CSTRING, "_rs", "is an optional record separator. Default is newline '\\n'."),
		Arg(XLL_CSTRING, "_fs", "is an optional field separator. Default is comma ','."),
		Arg(XLL_CSTRING, "_esc", "is an optional escape character. Default is backslash '\\'."),
		Arg(XLL_LPOPER, "_types", "is an optional array of types from XLTYPE_*."),
		Arg(XLL_WORD, "_drop", "is an optional number of rows to skip. Default is 0."),
		Arg(XLL_LONG, "_take", "is an optional number of rows to return. Default is all.")
		})
	.FunctionHelp("Parse CSV string into a range.")
	.Category("CSV")
	.Documentation(R"xyzyx(
Convert comma separated values to a range. 
)xyzyx")
);
LPOPER WINAPI xll_csv_parse(LPOPER pcsv, xcstr _rs, xcstr _fs, xcstr _e, const LPOPER ptypes, unsigned drop, LONG take)
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
			csv::parser<XLOPERX, char> records(v, rs, fs, e);

			unsigned n = ptypes->size();
			for (auto record : records.record_iterator()) {
				if (drop > 0 and drop--) continue;

				unsigned i = 0;
				OPER row;
				for (auto field : records.field_iterator(record)) {
					auto fi = OPER(field.buf, field.len);
					if (i < n) {
						parse::convert(fi, (*ptypes)[i]);
					}
					row.push_right(fi);
				}
				o.push_bottom(row);

				if (take > 0 and !--take) break;
			}
		}
		else {
			ensure(pcsv->is_str());

			xchar rs = *_rs ? *_rs : _T('\n');
			xchar fs = *_fs ? *_fs : _T(',');
			xchar e = *_e ? *_e : _T('\\');

			auto v = fms::view<xchar>(pcsv->val.str + 1, pcsv->val.str[0]);
			csv::parser<XLOPERX, xchar> records(v, rs, fs, e);

			unsigned n = ptypes->size();
			for (auto record : records.record_iterator()) {
				if (drop > 0 and drop--) continue;

				unsigned i = 0;
				OPER row;
				for (auto field : records.field_iterator(record)) {
					auto fi = OPER(field.buf, field.len);
					if (i < n) {
						parse::convert(fi, (*ptypes)[i]);
					}
					row.push_right(fi);
				}
				o.push_bottom(row);

				if (take > 0 and !--take) break;
			}
		}
		if (take < 0) {
			o.drop(take);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = ErrNA;
	}

	return &o;
}
/*
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

		for (unsigned j = 0; j < po->columns(); ++j) {
			const auto& type = (*ptypes)[j];
			if (type) {
				for (unsigned i = header; i < po->rows(); ++i) {
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
*/
#ifdef _DEBUG

Auto<OpenAfter> xaoa_test_csv_parse([]() { return xll::csv::test<TCHAR>() == 0; });

#endif // _DEBUG
