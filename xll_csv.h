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


