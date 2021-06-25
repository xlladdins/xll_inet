// xll_parse_json.h - parse JSON to OPER
// JSON objects map neatly to two colulmn OPERs: the key is the first first column and the value is the second.
// If the value is a multi then it is either an array or an object.
// If the multi has 1 column it is an array. If it has 2 columns it is an object.
#pragma once
#include "xll/xll/xll.h"
#include "xll/xll/codec.h"
#include "xll_parse.h"

namespace xll::parse::json {

	// forward declaration
	template<class X, class T>
	XOPER<X> value(view<const T> v);

	template<class X, class T = typename traits<X>::xchar>
	inline XOPER<X> string(view<const T> v)
	{
		auto str = skip<const T>(v, '"', '"', '\\');
		ensure(!skipws(v));

		return XOPER<X>(str.buf, str.len);
	}

	// object := "{ str : val , ... } "
	template<class X, class T = typename traits<X>::xchar>
	inline XOPER<X> object(view<const T> o)
	{
		XOPER<X> x;

		auto kvs = skip<const T>(o, '{', '}', '\\');
		ensure(!skipws(o));

		while (kvs = skipws(kvs)) {
			// key : val , ...
			auto val = chop<const T>(kvs, ',', '\\');
			auto key = chop<const T>(val, ':', '\\');
			x.push_back(XOPER<X>({ string<X>(key), value<X>(val) }));
		}
		ensure(!kvs);
		

		return x;
	}

	// array := "[ value , ... ] "
	template<class X, class T = typename traits<X>::xchar>
	inline XOPER<X> array(view<const T> v)
	{
		XOPER<X> x;

		auto a = skip<const T>(v, '[', ']', '\\');
		ensure(!skipws(v));

		while (a = skipws(a)) {
			auto elem = chop<const T>(a, ',', '\\');
			x.push_back(value<X>(elem));
		}
		ensure(!a);

		x.resize(x.size(), 1);

		return x;
	}

	template<class X, class T = typename traits<X>::xchar>
	inline XOPER<X> value(view<const T> v)
	{
		v = skipws(v);
		
		if (v.front() == '{') {
			return object<X>(v);
		}
		else if (v.front() == '[') {
			return array<X>(v);
		}
		else if (v.front() == '"') {
			return string<X>(v);
		}

		// more decode tests!!!
		return decode<X>(v.buf, v.len);
	}


#ifdef _DEBUG

	template<class X>
	inline int test()
	{
		{
			XOPER<X> x = string<X>(view(_T("\"str\"")));
			ensure(x == "str");
		}
		{
			XOPER<X> x = string<X>(view(_T("\"s\\\"r\"")));
			ensure(x == "s\\\"r");
		}
		{
			XOPER<X> x = string<X>(view(_T("\"s\nr\"")));
			ensure(x == "s\nr");
		}
		{
			XOPER<X> x = string<X>(view(_T("\"s r\"")));
			ensure(x == "s r");
		}
		{
			XOPER<X> x = array<X>(view(_T("[\"a\", 1.23, FALSE]")));
			ensure(x.size() == 3);
			ensure(x[0] == "a");
			ensure(x[1] == 1.23);
			ensure(x[2] == false);
		}
		{
			XOPER<X> x = object<X>(view(_T("{}")));
			ensure(x == XOPER<X>{});
		}
		{
			XOPER<X> x = object<X>(view(_T("{\"key\":\"value\"}")));
			ensure(x.columns() == 2);
			ensure(x[0] == "key");
			ensure(x[1] == "value");
		}
		{
			XOPER<X> x = object<X>(view(_T("{\"key\":\"value\",\"num\":1.23}")));
			ensure(x.rows() == 2);
			ensure(x.columns() == 2);
			ensure(x(0,0) == "key");
			ensure(x(0,1) == "value");
			ensure(x(1, 0) == "num");
			ensure(x(1, 1) == 1.23);
		}
		{
			XOPER<X> x = object<X>(view(_T("{ \"key\" : \"value\" , \"num\" : 1.23 } ")));
			ensure(x.rows() == 2);
			ensure(x.columns() == 2);
			ensure(x(0, 0) == "key");
			ensure(x(0, 1) == "value");
			ensure(x(1, 0) == "num");
			ensure(x(1, 1) == 1.23);
		}

		return 0;
	}

#endif // _DEBUG

}

