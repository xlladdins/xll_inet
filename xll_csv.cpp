// xll_csv.cpp - Parse CSV strings
#include "xll_parse.h"

using namespace xll;

#define XLTYPE_TOPIC "https://support.microsoft.com/en-us/office/type-function-45b4e688-4bc3-48b3-a105-ffa892995899"
#define X(a, b, c) XLL_CONST(LONG, XLTYPE##b, a, c, "XLL", XLTYPE_TOPIC)
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

#if 0

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
#endif // 0

AddIn xai_csv_parse(
	Function(XLL_LPOPER, "xll_csv_parse", "CSV.PARSE")
	.Arguments({
		Arg(XLL_HANDLEX, "csv", "is a handle to a string of comma separated values."),
		Arg(XLL_CSTRING4, "_rs", "is an optional record separator. Default is newline '\\n'."),
		Arg(XLL_CSTRING4, "_fs", "is an optional field separator. Default is comma ','."),
		Arg(XLL_CSTRING4, "_esc", "is an optional escape character. Default is backslash '\\'."),
		})
	.FunctionHelp("Parse handle to a CSV string into a range.")
	.Category("CSV")
	.Documentation(R"xyzyx(
Convert comma separated values to a range. 
)xyzyx")
);
LPOPER WINAPI xll_csv_parse(HANDLEX hcsv, const char* _rs, const char* _fs, const char* _e)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		handle<fms::view<char>> h_(hcsv);
		ensure(h_);

		char rs = *_rs ? *_rs : '\n';
		char fs = *_fs ? *_fs : ',';
		char e = *_e ? *_e : '\\';
		auto v = fms::char_view<char>(h_->buf, h_->len);

		unsigned r = 0;
		unsigned c = 0;
		auto records = fms::parse::splitable(v, rs, '"', '"', e);
		for (auto record : records) {
			unsigned i = 0;
			auto fields = fms::parse::splitable(record, fs, '"', '"', e);
			for (auto field : fields) {
				if (r == 0) {
					c = static_cast<unsigned>(std::distance(fields.begin(), fields.end()));
					o.resize(1, c);
				}
				else if (i == 0) {
					o.resize(r + 1, c);
				}
				ensure(i < c);
				o(r, i) = OPER(field.buf, field.len);
				++i;
			}
			++r;
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = ErrNA;
	}

	return &o;
}

AddIn xai_csv_parse_timeseries(
	Function(XLL_FPX, "xll_csv_parse_timeseries", "CSV.PARSE.TIMESERIES")
	.Arguments({
		Arg(XLL_HANDLEX, "view", "is handle to a view."),
		Arg(XLL_CSTRING4, "_rs", "is an optional record separator. Default is newline '\\n'."),
		Arg(XLL_CSTRING4, "_fs", "is an optional field separator. Default is comma ','."),
		Arg(XLL_CSTRING4, "_esc", "is an optional escape character. Default is backslash '\\'."),
		})
	.FunctionHelp("Parse view into a timeseries.")
	.Category("CSV")
	.Documentation(R"xyzyx(
Convert comma separated values to a range. First column must be a date.
)xyzyx")
);
_FPX* WINAPI xll_csv_parse_timeseries(HANDLEX csv, const char* _rs, const char* _fs, const char* _e)
{
#pragma XLLEXPORT
	static FPX o;

	try {
		handle<fms::view<char>> h_(csv);
		ensure(h_);

		char rs = /*static_cast<char>*/(*_rs ? *_rs : '\n');
		char fs = static_cast<char>(*_fs ? *_fs : ',');
		char e = static_cast<char>(*_e ? *_e : '\\');
		auto v = fms::char_view<const char>(h_->buf, h_->len);

		unsigned r = 0;
		unsigned c = 0;
		auto records = fms::parse::splitable<const char>(v, rs, '"', '"', e);
		for (auto record : records) {
			if (!std::isdigit(record.front())) {
				continue; // skip character data
			}
			unsigned i = 0;
			auto fields = fms::parse::splitable<const char>(record, fs, '"', '"', e);
			for (auto field : fields) {
				if (r == 0) {
					c = static_cast<unsigned>(std::distance(fields.begin(), fields.end()));
					o.resize(1, c);
				}
				else if (i == 0) {
					o.resize(r + 1, c);
				}
				ensure(i < c);
				double x;
				if (i == 0) {
					auto [y, m, d] = fms::parse::to_ymd(field);
					x = Excel(xlfDate, OPER(y), OPER(m), OPER(d)).as_num();
					ensure(!field.is_error());
					if (field and field.drop(1)) {
						auto [hh, mm, ss] = fms::parse::to_hms(field);
						ensure(!field.is_error());
						if (field) {
							auto [ho, mo] = fms::parse::to_off(field);
							ensure(!field.is_error());
							ensure(!field);
							hh += ho;
							mm += mo;
						}
						x += Excel(xlfTime, OPER(hh), OPER(mm), OPER(ss)).as_num();
					}
				}
				else {
					x = fms::parse::to<double>(field);
				}
				o(r, i) = x;
				++i;
			}
			++r;
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o.resize(0, 0);
	}

	return o.get();
}
