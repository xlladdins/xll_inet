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

	/*
	//!!! use xll::codec
	template<class X, class T = typename traits<X>::xchar>
	inline XOPER<X> field(fms::view<const T> field)
	{
		return decode<X, T>(field);
	}
	*/

	// use view!!!
	template<class X, class T>
	inline XOPER<X> parse(fms::view<T> buf, T rs, T fs, T e, unsigned off = 0, unsigned count = 0)
	{
		XOPER<X> o;
		
		for (const auto& r : parse::iterator<T>(buf, rs, 0, 0, e)) {
			OPER row;

			if (off != 0) {
				--off;

				continue;
			}

			for (const auto& f : parse::iterator<T>(r, fs, 0, 0, e)) {
				double num = parse::to_number<T>(f);
				if (!isnan(num)) {
					row.push_right(XOPER<X>(num));
				}
				else {
					row.push_right(XOPER<X>(f.buf, static_cast<T>(f.len)));
				}
			}

			//xlerrNA
			if (row.size() < o.columns()) {
				row.resize(1, o.columns()); // pad
			}
			/*
			else if (o and row.size() > o.columns()) {
				// widen range
				OPER pad(o.rows(), row.size() - row.columns());
				o.push_back(pad, OPER::Side::Right);
			}
			*/
			o.push_bottom(row);

			if (count) {
				if (--count == 0)
					break;
			}
		}
		
		return o;
	}

#ifdef _DEBUG

	template<class T>
	inline int test_parse_csv()
	{
		{
			fms::view<const T> v(_T("a,b\nc,d"));
			OPER o = parse<XLOPERX, const T>(v, _T('\n'), _T(','), _T('\\'));
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


