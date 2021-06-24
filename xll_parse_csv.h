// xll_parse_csv.h - Comma Separated Value
#pragma once
#include <compare>
#include <iterator>
#include "ISO8601.h"
#include "xll/xll/xll.h"
#include "xll_parse.h"

#ifndef CATEGORY
#define CATEGORY "XLL"
#endif

namespace xll::parse::csv {

	//!!! use xll::codec
	template<class T>
	inline OPER parse_field(view<T> field)
	{
		OPER o(field.buf, static_cast<T>(field.len));
		o = Excel(xlfTrim, o);

		if (o.val.str[0] == 4 and 0 == _tcsncmp(o.val.str + 1, _T("TRUE"), 4))
			return OPER(true);
		if (o.val.str[0] == 5 and 0 == _tcsncmp(o.val.str + 1, _T("FALSE"), 5))
			return OPER(false);

		// number or date?
		OPER o_(Excel(xlfValue, o));
		if (o_) {
			o = o_;
		}
		// ISO 8601???

		return o;
	}

#ifdef _DEBUG

	inline int test_parse_field()
	{
		ensure(parse_field(view(_T("abc"))) == OPER("abc"));
		ensure(parse_field(view(_T(" abc "))) == OPER("abc"));
		ensure(parse_field(view(_T("1.23"))) == OPER(1.23));
		ensure(parse_field(view(_T("-123"))) == OPER(-123));
		ensure(parse_field(view(_T("2020-1-2"))) == Excel(xlfDate, OPER(2020), OPER(1), OPER(2)));
		ensure(parse_field(view(_T("Jan 2, 2020"))) == Excel(xlfDate, OPER(2020), OPER(1), OPER(2)));
		ensure(parse_field(view(_T("1:30"))) == OPER(1.5 / 24));
		ensure(parse_field(view(_T("TRUE"))) == OPER(true));
		ensure(parse_field(view(_T("true"))) != OPER(true));

		return 0;
	}

#endif // _DEBUG
	/*
	template<class T>
	using line = chop<T, '\n', '\"', '\"', '\"'>;
	template<class T>
	using row = chop<T, ',', '\"', '\"', '\"'>;
	*/

	template<class T>
	inline OPER parse(const T* buf, DWORD len, T rs, T fs, T l, T r, T e)
	{
		OPER o;
		
		for (const auto& row : iterator(xll::view<const T>(buf, len), rs, l, r, e)) {
			OPER ro;
			for (const auto& field : iterator(row, fs, l, r, e)) {
				ro.push_back(parse_field(field));
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
			xll::view v(_T("a,b\nc,d"));
			OPER o = parse(v.buf, v.len, _T('\n'), _T(','), _T('\"'), _T('\"'), _T('\\'));
			ensure(o.rows() == 2);
			ensure(o.columns() == 2);
			ensure(o(0, 0) == "a");
			ensure(o(0, 1) == "b");
			ensure(o(1, 0) == "c");
			ensure(o(1, 1) == "d");
		}
		{
			xll::view v(_T("abc,FALSE\n1.23,2001-2-3"));
			OPER o = parse(v.buf, v.len, _T('\n'), _T(','), _T('\"'), _T('\"'), _T('\\'));
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
		test_parse_field();
		test_parse_csv<TCHAR>();

		return 0;
	}

#endif // _DEBUG

} // namespace xll::csv


