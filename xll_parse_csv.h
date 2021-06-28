// xll_parse_csv.h - Comma Separated Value
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

namespace xll::parse::csv {

	//!!! use xll::codec
	template<class X>
	inline XOPER<X> parse_field(fms::view<const typename traits<X>::xchar> field)
	{
		return decode<X>(field);
	}

	/*
	template<class T>
	using line = chop<T, '\n', '\"', '\"', '\"'>;
	template<class T>
	using row = chop<T, ',', '\"', '\"', '\"'>;
	*/

	template<class X, class T = typename traits<X>::xchar>
	inline XOPER<X> parse(const T* buf, DWORD len, T rs, T fs, T e)
	{
		XOPER<X> o;
		
		for (const auto& row : iterator<T>(fms::view<const T>(buf, len), rs, 0, 0, e)) {
			OPER ro;
			for (const auto& field : iterator<T>(row, fs, 0, 0, e)) {
				ro.push_back(parse_field<X>(field));
			}
			ro.resize(1, ro.size());
			o.push_back(ro);
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


