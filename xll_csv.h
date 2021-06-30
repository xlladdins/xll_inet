// xll_csv.h - Comma Separated Value
#pragma once
#include <compare>
#include <iterator>
#include "ISO8601.h"
#include "xll/xll/xll.h"
//#include "xll/xll/codec.h"
#include "xll_parse.h"

#ifndef CATEGORY
#define CATEGORY "XLL"
#endif

namespace xll::csv {

	//!!! use xll::codec
	template<class X, class T = typename traits<X>::xchar>
	inline XOPER<X> field(fms::view<const T> field)
	{
		return decode<X, T>(field);
	}

	template<class X, class T = typename traits<X>::xchar>
	inline XOPER<X> parse(const T* buf, DWORD len, T rs, T fs, T e, unsigned off = 0, unsigned count = 0)
	{
		XOPER<X> o;
		
		for (const auto& r : iterator<T>(fms::view<const T>(buf, len), rs, 0, 0, e)) {
			OPER row;

			if (off != 0) {
				--off;

				continue;
			}

			for (const auto& f : iterator<T>(r, fs, 0, 0, e)) {
				row.push_back(XOPER<X>(f.buf, static_cast<T>(f.len)));
			}
			row.resize(1, row.size());
			o.push_back(row);

			if (count != 0) {
				--count;
				if (count == 0) {
					break;
				}
			}
		}
		
		return o;
	}

#ifdef _DEBUG

	template<class T>
	inline int test_parse_csv()
	{
		{
			fms::view v(_T("a,b\nc,d"));
			OPER o = parse<XLOPERX>(v.buf, v.len, _T('\n'), _T(','), _T('\\'));
			ensure(o.rows() == 2);
			ensure(o.columns() == 2);
			ensure(o(0, 0) == "a");
			ensure(o(0, 1) == "b");
			ensure(o(1, 0) == "c");
			ensure(o(1, 1) == "d");
		}
		{
			fms::view v(_T("abc,FALSE\n1.23,2001-2-3"));
			OPER o = parse<XLOPERX>(v.buf, v.len, _T('\n'), _T(','), _T('\\'));
			ensure(o.rows() == 2);
			ensure(o.columns() == 2);
			ensure(o(0, 0) == "abc");
			ensure(o(0, 1) == false);
			ensure(o(1, 0) == 1.23);
			ensure(o(1, 1) == Excel(xlfDate, OPER(2001), OPER(2), OPER(3)));
		}

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


