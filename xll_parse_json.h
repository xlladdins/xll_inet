// xll_parse_json.h - parse JSON to OPER
#pragma once
#include "xll/xll/xll.h"
#include "xll/xll/xll_codec.h"
#include "xll_parse.h"

static inline const char xll_parse_json_doc[] = R"xyzyx(
JSON objects map neatly to two row OPERs: the keys are in the first row and the values in the second.
If a value is a multi then it is either an array or an object.
If a value is a multi with 1 row it is an array and if it has 2 rows it is an object.
<p>
Recall a JSON object has zero or more key-value pairs of the form <code>{ "key" : value, ... }</code>.
A scalar JSON value is either a string, number, boolean, or null.
These are are parsed to <code>OPER</code> type <code>xltypeStr</code>, <code>xltypeNum</code>,
<code>xltypeBool</code>, and <code>xltypeErr</code> with <code>.val.err = xlerrNull</code>.
JSON array values have zero or more values of the form <code>[ value, ... ]</code>.
JSON objects are also values so the structure is recursive.
</p>
)xyzyx";

namespace xll::parse::json {

	// multi-level index into JSON value
	template<class X>
	inline XOPER<X> index(const XOPER<X>& v, XOPER<X> i)
	{
		if (v.is_multi()) {
			if (v.rows() == 1) { // array
				ensure(i[0].is_num());
				const XOPER<X>& vi = v[(unsigned)i[0].val.num];

				return i.size() == 1 ? vi : index(vi, i.drop(1));
			}
			else if (v.rows() == 2) { // object
				if (i[0].is_num()) {
					const XOPER<X>& vi = v(1, (unsigned)i[0].val.num);

					return i.size() == 1 ? vi : index(vi, i.drop(1));
				}
				else {
					X v1 = v; // first row of v
					v1.val.array.rows = 1;
					XOPER<X> match = Excel(xlfMatch, v1, i[0], XOPER<X>(0)); // exact
					ensure(match.is_num());
					const XOPER<X>& vi = v(1, (unsigned)match.val.num - 1);

					return i.size() == 1 ? vi : index(vi, i.drop(1));
				}
			}
			else {
				return XOPER<X>(XOPER<X>::Err::NA);
			}
		}

		ensure(i == 0);

		return v;
	}

	// forward declaration
	template<class X, class T>
	XOPER<X> value(fms::view<const T> v);

	// "\"str\"" => "str"
	template<class X, class T = typename traits<X>::xchar>
	inline XOPER<X> string(fms::view<const T> v)
	{
		auto str = skip<T>(v, '"', '"', '\\');
		ensure(!skipws(v));

		return XOPER<X>(str.buf, str.len);
	}

	// object := "{ \"str\" : val , ... } "
	template<class X, class T = typename traits<X>::xchar>
	inline XOPER<X> object(fms::view<const T> v)
	{
		XOPER<X> x;

		auto o = skip<T>(v, '{', '}', '\\');
		ensure(!skipws(v));

		while (o = skipws(o)) {
			// key : val , ...
			auto val = chop<const T>(o, ',', '{', '}', '\\');
			auto key = chop<const T>(val, ':', '"', '"', '\\');
			XOPER<X> kv = XOPER<X>({ string<X>(key), value<X>(val) });
			kv.resize(2, 1);
			// keys are in first row
			x.push_back(kv, XOPER<X>::Side::Right);
		}
		
		return x;
	}

	// array := "[ value , ... ] "
	template<class X, class T = typename traits<X>::xchar>
	inline XOPER<X> array(fms::view<const T> v)
	{
		XOPER<X> x;

		auto a = skip<T>(v, '[', ']', '\\');
		ensure(!skipws(v));

		while (a = skipws(a)) {
			x.push_back(value<X>(chop<T>(a, ',', '{', '}', '\\')));
		}

		x.resize(1, x.size());

		return x;
	}

	template<class X, class T = typename traits<X>::xchar>
	inline XOPER<X> value(fms::view<const T> v)
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
		return decode<X>(v);
	}


#ifdef _DEBUG

	template<class X>
	inline int test()
	{
		{
			XOPER<X> x = string<X>(fms::view(_T("\"str\"")));
			ensure(x == "str");
		}
		{
			XOPER<X> x = string<X>(fms::view(_T("\"s\\\"r\"")));
			ensure(x == "s\\\"r");
		}
		{
			XOPER<X> x = string<X>(fms::view(_T("\"s\nr\"")));
			ensure(x == "s\nr");
		}
		{
			XOPER<X> x = string<X>(fms::view(_T("\"s r\"")));
			ensure(x == "s r");
		}
		{
			XOPER<X> x = array<X>(fms::view(_T("[\"a\", 1.23, FALSE]")));
			ensure(x.size() == 3);
			ensure(x[0] == "a");
			ensure(x[1] == 1.23);
			ensure(x[2] == false);
		}
		{
			XOPER<X> x = object<X>(fms::view(_T("{}")));
			ensure(x == XOPER<X>{});
		}
		{
			XOPER<X> x = object<X>(fms::view(_T("{\"key\":\"value\"}")));
			ensure(x.rows() == 2);
			ensure(x.columns() == 1);
			ensure(x[0] == "key");
			ensure(x[1] == "value");
		}
		{
			XOPER<X> x = object<X>(fms::view(_T("{\"key\":\"value\",\"num\":1.23}")));
			ensure(x.rows() == 2);
			ensure(x.columns() == 2);
			ensure(x(0,0) == "key");
			ensure(x(1,0) == "value");
			ensure(x(0,1) == "num");
			ensure(x(1, 1) == 1.23);
		}
		{
			XOPER<X> x = object<X>(fms::view(_T("{ \"key\" : \"value\" , \"num\" : 1.23 } ")));
			ensure(x.rows() == 2);
			ensure(x.columns() == 2);
			ensure(x(0, 0) == "key");
			ensure(x(1, 0) == "value");
			ensure(x(0, 1) == "num");
			ensure(x(1, 1) == 1.23);
		}
		{
			//XOPER<X> x = object<X>(fms::view(_T("{\"entities\":{\"Q64\":{\"type\":\"item\",\"id\":\"Q64\",\"descriptions\":{\"en\":{\"language\":\"en\",\"value\":\"federal state, capitaland largest city of Germany\"}}}},\"success\":1}")));
			XOPER<X> x = object<X>(fms::view(_T("{\"a\":{\"b\":\"ab\"}}")));
			ensure(x.rows() == 2);
		}

		return 0;
	}

#endif // _DEBUG

}

