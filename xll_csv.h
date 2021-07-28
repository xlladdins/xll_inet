// xll_csv.h - Comma Separated Value
#pragma once
#include <compare>
#include <iterator>
//#include "ISO8601.h"
#include "xll/xll/xll.h"
//#include "xll/xll/codec.h"
#include "xll_parse.h"

#ifndef CATEGORY
#define CATEGORY "XLL"
#endif

namespace xll::csv {

	template<class X, class T>
	struct parser {
		fms::view<T> buf;
		T rs, fs, esc;

		parser()
		{ }
		parser(const fms::view<T>& buf, T rs, T fs, T esc)
			: buf(buf), rs(rs), fs(fs), esc(esc)
		{ }
		parser(const parser&) = default;
		parser& operator=(const parser&) = default;
		~parser()
		{ }

		parse::iterator<T> record_iterator()
		{
			return parse::iterator<T>(buf, rs, 0, 0, esc);
		}
		parse::iterator<T> field_iterator(parse::iterator<T>::value_type& record)
		{
			return parse::iterator<T>(record, fs, 0, 0, esc);
		}
	};

	template<class X, class T>
	inline XOPER<X> parse(const fms::view<T>& buf, T rs, T fs, T esc)
	{
		XOPER<X> o;

		for (auto record : parse::iterator<T>(buf, rs, 0, 0, esc)) {
			XOPER<X> row;

			for (auto field : parse::iterator<T>(record, fs, 0, 0, esc)) {
				row.push_right(XOPER<X>(field.buf, static_cast<T>(field.len)));
			}

			if (row.size() < o.columns()) {
				row.resize(1, o.columns()); // pad
			}
			else if (o and row.size() > o.columns()) {
				// widen range
				o.push_right(OPER(o.rows(), row.size() - row.columns()));
			}

			o.push_bottom(row);
		}
		
		return o;
	}

	// assumes first column is date/time and others are num
	template<class X>
	inline XOPER<X> parse_timeseries(fms::view<char>& buf, char rs, char fs, char esc)
	{
		XOPER<X> o;
		XOPER<X> row;

		for (auto record : parse::iterator<char>(buf, rs, 0, 0, esc)) {
			record.skipws();
			if (!isdigit(record.buf[0])) {
				if (o != Nil) { // ignore headers
					XLL_WARNING(__FUNCTION__ ": record is not numeric");
				}

				continue;
			}

			unsigned i = 0;
			int convert = xlfDatevalue;
			for (auto field : parse::iterator<char>(record, fs, 0, 0, esc)) {
				char c = field.buf[-1];

				field.buf[-1] = static_cast<char>(field.len);
				XLOPER fi = { .val = { .str = field.buf - 1 }, .xltype = xltypeStr };
				fi = Excel4(convert, fi);
				
				field.buf[-1] = c;

				ensure(fi.xltype == xltypeNum);
				if (row.size() <= i) {
					row.resize(1, i + 1);
				}
				row[i].xltype = xltypeNum;
				row[i].val.num = fi.val.num;

				convert = xlfValue;
				++i;
			}

			if (row.size() < o.columns()) {
				row.resize(1, o.columns()); // pad
			}
			else if (o and row.size() > o.columns()) {
				// widen range
				XLL_INFO(__FUNCTION__ ": widening range");
				o.push_right(OPER(o.rows(), row.size() - row.columns()));
			}

			o.push_bottom(row);
		}

		return o;
	}

#ifdef _DEBUG

	template<class T>
	inline int test_parse_csv()
	{
		{
			fms::view<const T> v(_T("a,b\nc,d"));
			OPER o = csv::parse<XLOPERX,const T>(v, _T('\n'), _T(','), _T('\\'));
			ensure(o.rows() == 2);
			ensure(o.columns() == 2);
			ensure(o(0, 0) == "a");
			ensure(o(0, 1) == "b");
			ensure(o(1, 0) == "c");
			ensure(o(1, 1) == "d");
		}
		/*
		{
			fms::view v(_T("abc,FALSE\n1.23,2001-2-3"));
			OPER o = parse<XLOPERX>(v.buf, v.len, _T('\n'), _T(','), _T('\\'));
			ensure(o.rows() == 2);
			ensure(o.columns() == 2);
			ensure(o(0, 0) == "abc");
			ensure(decode<XLOPERX>(o(0, 1)) == false);
			ensure(decode<XLOPERX>(o(1, 0)) == 1.23);
			ensure(decode<XLOPERX>(o(1, 1)) == Excel(xlfDate, OPER(2001), OPER(2), OPER(3)));
		}
		*/

		return 0;
	}

	template<class T>
	inline int test()
	{
		test_parse_csv<TCHAR>();

		return 0;
	}

#endif // _DEBUG

} // namespace xll::csv


