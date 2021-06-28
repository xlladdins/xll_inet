// xll_parse_json.h - parse JSON to OPER
#pragma once
#include "xll/xll/xll.h"
#include "xll/xll/xll_codec.h"
#include "xll_parse.h"

static inline const char xll_parse_json_doc[] = R"xyzyx(
JSON objects map neatly to two row <code>OPER</code>s: the keys are in the first row and the values in the second.
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
					XOPER<X> match = Excel(xlfMatch, i[0], v1, XOPER<X>(0)); // exact
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

		return XOPER<X>(str.buf, str.len);
	}

	// object := "{ \"str\" : val , ... } "
	template<class X, class T = typename traits<X>::xchar>
	inline XOPER<X> object(fms::view<const T> v)
	{
		XOPER<X> x;

		auto o = skip<T>(v, '{', '}', '\\');
		ensure(!v.skipws());

		while (o.skipws()) {
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

		while (a.skipws()) {
			x.push_back(value<X>(chop<T>(a, ',', '{', '}', '\\')));
		}

		x.resize(1, x.size());

		return x;
	}

	template<class X, class T = typename traits<X>::xchar>
	inline XOPER<X> value(fms::view<const T> v)
	{
		v.skipws();
		
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
		XOPER<X> val = decode<X,T>(v);
		static XOPER<X> xnull("null");
		static XOPER<X> null(XOPER<X>::Err::Null);

		return val == xnull ? null : val;
	}


#ifdef _DEBUG

#define XLL_PARSE_JSON_VALUE(X) \
	X("null", OPER(OPER::Err::Null)) \
	X("\"str\"", "str") \
	X("\"s\\\"r\"", "s\\\"r") \
	X("\"s\nr\"", "s\nr") \
	X("\"s r\"", "s r") \
	X("[\"a\", 1.23, FALSE]", OPER({OPER("a"), OPER(1.23), OPER(false)})) \
	X("{\"key\":\"value\"}", OPER({OPER("key"), OPER("value")}).resize(2,1)) \
	X("{\"key\":\"value\",\"num\":1.23}", OPER({OPER("key"), OPER("num"), OPER("value"), OPER(1.23)}).resize(2,2)) \
	X(" { \"key\" : \"value\" , \"num\" : 1.23 }", OPER({OPER("key"), OPER("num"), OPER("value"), OPER(1.23)}).resize(2,2)) \
	X("\r{ \"key\" \t\n: \r\"value\"\r ,\n\t \"num\" : 1.23  }", OPER({OPER("key"), OPER("num"), OPER("value"), OPER(1.23)}).resize(2,2)) \
	X("{\"a\":{\"b\":\"ab\"}}", OPER({OPER("a"), OPER({OPER("b"), OPER("ab")}).resize(2,1)}).resize(2,1)) \
	X(" { \"a\":\n{\"b\":\r\n\t\"ab\" } }", OPER({OPER("a"), OPER({OPER("b"), OPER("ab")}).resize(2,1)}).resize(2,1)) \

	template<class X>
	inline int test()
	{
#define PARSE_JSON_CHECK(a, b) { ensure(value<X>(fms::view(_T(a))) == b); }
		XLL_PARSE_JSON_VALUE(PARSE_JSON_CHECK)
#undef PARSE_JSON_CHECK

		OPER x = object<XLOPERX>(fms::view(_T("{\"a\":{\"b\":\"cd\"}}")));
		OPER i({ OPER("a"), OPER("b") });
		ensure(json::index(x, i[0]) == OPER({ OPER("b"), OPER("cd") }).resize(2,1));
		ensure(json::index(x, i) == "cd");

		return 0;
	}
#undef XLL_PARSE_JSON_VALUE
#endif // _DEBUG

}

