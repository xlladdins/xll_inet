// xll_csv.cpp - Parse CSV strings
#include "xll_csv.h"

using namespace xll;

#define XLTYPE_TOPIC "https://support.microsoft.com/en-us/office/type-function-45b4e688-4bc3-48b3-a105-ffa892995899"
#define X(a, b, c) XLL_CONST(LONG, XLTYPE##b, a, c, "XLL", XLTYPE_TOPIC)
XLTYPE(X)
#undef XLTYPE_TOPIC
#undef X

using xcstr = xll::traits<XLOPERX>::xcstr;
using xchar = xll::traits<XLOPERX>::xchar;

AddIn xai_csf3(Macro("xll_csf3", "C.S.F3").Documentation(R"(
Like <code>Ctrl-Shift-F3</code> this macro defines names using
selected text. It looks for a handle in the cell below
the first item in the selected range and defines names
corresponding to the columns in the array or range
associated with the handle.
)"));
int WINAPI xll_csf3()
{
#pragma XLLEXPORT
	try {
		OPER sel = Excel(xlfSelection);
		OPER names = Excel(xlCoerce, sel);

		OPER cell = Excel(xlfOffset, sel, OPER(1), OPER(0), OPER(1), OPER(1));
		cell = Excel(xlfAbsref, OPER(REF(0, 0)), cell);
		OPER handlex = Excel(xlCoerce, cell);
		ensure(handlex.is_num());
	
		// !!!get.name(cell) and delete name
		// define unique name for handlex
		OPER name(std::tmpnam(nullptr));
		name = name.safe();
		Excel(xlcDefineName, name, cell, OPER(1));

		OPER index;
		{
			handle<FPX> a_(handlex.as_num());
			if (a_) {
				index = OPER("=array.index(") & name & OPER(",,");
			}
		}
		if (!index) {
			handle<OPER> r_(handlex.as_num());
			if (r_) {
				index = OPER("=range.index(") & name & OPER(",,");
			}
		}
		ensure(index);

		// num to text
		auto text = [](unsigned i) {
			return Excel(xlfText, OPER(i), OPER("0"));
		};
		for (unsigned i = 0; i < names.size(); ++i) {
			OPER namei = index & text(i) & OPER(")");
			Excel(xlcDefineName, names[i], namei, OPER(1));
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}

	return TRUE;
}

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
		Arg(XLL_LPOPER, "cell", "is a cell or range to convert."),
		Arg(XLL_LPOPER, "type", "is a type or array of types from the XLTYPE_* enumeration.")
		})
	.Category("XLL")
	.FunctionHelp("Convert cell to type from XLTYPE_* enumeration.")
	.Documentation(R"(
If a conversion fails <code>#VALUE!</code> is returned.
)")
);
LPOPER WINAPI xll_xltype_convert(LPOPER po, LPOPER ptype)
{
#pragma XLLEXPORT
	static OPER o;

	o = ErrValue;
	try {
		if (po->is_num()) {
			handle<OPER> o_(po->as_num());
			ensure(o_);
			po = o_.ptr();
		}

		o = *po; // todo: operate on handle and return handle

		if (po->rows() == 1) {
			ensure(o.size() == ptype->size());

			for (unsigned i = 0; i < o.size(); ++i) {
				parse::convert(o[i], (*ptype)[i]);
			}
		}
		else {
			// 2-d range
			ensure(o.columns() == size(*ptype));

			for (unsigned i = 0; i < o.rows(); ++i) {
				for (unsigned j = 0; j < o.columns(); ++j) {
					parse::convert(o(i, j), (*ptype)[j]);
				}
			}
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}

AddIn xai_csv_parse(
	Function(XLL_LPOPER, "xll_csv_parse", "CSV.PARSE")
	.Arguments({
		Arg(XLL_LPOPER, "csv", "is a string of comma separated values or handle."),
		Arg(XLL_CSTRING, "_rs", "is an optional record separator. Default is newline '\\n'."),
		Arg(XLL_CSTRING, "_fs", "is an optional field separator. Default is comma ','."),
		Arg(XLL_CSTRING, "_esc", "is an optional escape character. Default is backslash '\\'."),
		})
	.FunctionHelp("Parse CSV string into a range.")
	.Category("CSV")
	.Documentation(R"xyzyx(
Convert comma separated values to a range. 
)xyzyx")
);
LPOPER WINAPI xll_csv_parse(LPOPER pcsv, xcstr _rs, xcstr _fs, xcstr _e)
{
#pragma XLLEXPORT
	static OPER o;

	o = ErrNA;
	try {
		if (pcsv->is_num()) {
			handle<fms::view<char>> h_(pcsv->as_num());
			ensure(h_);

			char rs = static_cast<char>(*_rs ? *_rs : '\n');
			char fs = static_cast<char>(*_fs ? *_fs : ',');
			char e = static_cast<char>(*_e ? *_e : '\\');
			auto v = fms::view<char>(h_->buf, h_->len);

			o = csv::parse<XLOPERX, char>(v, rs, fs, e);
		}
		else {
			ensure(pcsv->is_str());

			xchar rs = *_rs ? *_rs : _T('\n');
			xchar fs = *_fs ? *_fs : _T(',');
			xchar e = *_e ? *_e : _T('\\');
			auto v = fms::view<xchar>(pcsv->val.str + 1, pcsv->val.str[0]);
			
			o = csv::parse<XLOPERX, xchar>(v, rs, fs, e);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}

AddIn xai_csv_parse_timeseries(
	Function(XLL_LPOPER, "xll_csv_parse_timeseries", "CSV.PARSE.TIMESERIES")
	.Arguments({
		Arg(XLL_HANDLEX, "view", "is handle to a view."),
		Arg(XLL_CSTRING, "_rs", "is an optional record separator. Default is newline '\\n'."),
		Arg(XLL_CSTRING, "_fs", "is an optional field separator. Default is comma ','."),
		Arg(XLL_CSTRING, "_esc", "is an optional escape character. Default is backslash '\\'."),
		})
	.FunctionHelp("Parse view into a timeseries.")
	.Category("CSV")
	.Documentation(R"xyzyx(
Convert comma separated values to a range. First column must be a date.
)xyzyx")
);
LPOPER WINAPI xll_csv_parse_timeseries(HANDLEX csv, xcstr _rs, xcstr _fs, xcstr _e)
{
#pragma XLLEXPORT
	static OPER o;

	o = ErrNA;
	try {
		handle<fms::view<char>> h_(csv);
		ensure(h_);

		char rs = static_cast<char>(*_rs ? *_rs : '\n');
		char fs = static_cast<char>(*_fs ? *_fs : ',');
		char e = static_cast<char>(*_e ? *_e : '\\');
		auto v = fms::view<char>(h_->buf, h_->len);

		o = csv::parse_timeseries<XLOPERX>(v, rs, fs, e);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}

