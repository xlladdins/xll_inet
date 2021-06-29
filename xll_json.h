// xll_json.h - JSON to and from OPER
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

namespace xll::json {

	// JSON construction
	template<class... Os>
	inline xll::OPER array(const Os&... os)
	{
		return xll::OPER({ xll::OPER(os)... }).resize(1, static_cast<unsigned>(sizeof...(os)));
	}
	// { k, ..., v, ...}
	template<class... Os>
	inline xll::OPER object(const Os&... os)
	{
		static_assert(0 == sizeof...(os) % 2);

		return array(os...).resize(2, static_cast<unsigned>(sizeof...(os)/2));
	}

	inline OPER string(const OPER& s)
	{
		return OPER("\"") & Excel(xlfSubstitute, s, OPER("\""), OPER("\\\"")) & OPER("\"");
	}

	// multi-level index into JSON value
	template<class X>
	inline XOPER<X> index(const XOPER<X>& v, XOPER<X> i)
	{
		if (v.is_multi()) {
			if (v.rows() == 1) { // array
				if (!i[0].is_num()) return XErrNA<X>;
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
					if (!match.is_num()) return match;
					const XOPER<X>& vi = v(1, (unsigned)match.val.num - 1);

					return i.size() == 1 ? vi : index(vi, i.drop(1));
				}
			}
			else {
				return XErrNA<X>;
			}
		}

		if (i != 0 or i[0] != 0) return XErrNA<X>;

		return v;
	}

} // xll::json

namespace xll::json::parse {

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

		XOPER<X> val = decode<X, T>(v);

		return val == "null" ? XErrNull<X> : val;
	}

#ifdef _DEBUG

#define XLL_PARSE_JSON_VALUE(X) \
	X("null", XErrNull<XLOPERX>) \
	X("\"str\"", "str") \
	X("\"s\\\"r\"", "s\\\"r") \
	X("\"s\nr\"", "s\nr") \
	X("\"s r\"", "s r") \
	X("[\"a\", 1.23, FALSE]", json::array("a", 1.23, false)) \
	X("{\"key\":\"value\"}", json::object("key", "value")) \
	X("{\"key\":\"value\",\"num\":1.23}", json::object("key", "num", "value", 1.23)) \
	X(" { \"key\" : \"value\" , \"num\" : 1.23 }", json::object("key", "num", "value", 1.23)) \
	X("\r{ \"key\" \t\n: \r\"value\"\r ,\n\t \"num\" : 1.23  }", json::object("key", "num", "value", 1.23)) \
	X("{\"a\":{\"b\":\"ab\"}}", json::object("a", json::object("b", "ab"))) \
	X(" { \"a\":\n{\"b\":\r\n\t\"ab\" } }", json::object("a", json::object("b", "ab"))) \
	
	//	X("[{\"a\":1},false,null]", json::array(json::object("a", 1), false, XErrNull<XLOPERX>)) \
	//	X("null", OPER::Err::Null) \
	//	X("[{\"a\":1},false,null]", json::array(json::object("a", 1), false, OPER::Err::Null)) \
		
	inline int test()
	{

#define PARSE_JSON_CHECK(a, b) { ensure(parse::value<XLOPERX>(fms::view(_T(a))) == b); }
		XLL_PARSE_JSON_VALUE(PARSE_JSON_CHECK)
#undef PARSE_JSON_CHECK

		OPER x = parse::object<XLOPERX>(fms::view(_T("{\"a\":{\"b\":\"cd\"}}")));
		OPER i({ OPER("a"), OPER("b") });
		ensure(json::index(x, i) == "cd");
		ensure(json::index(x, i[0]) == OPER({ OPER("b"), OPER("cd") }).resize(2, 1));

		return 0;
	}
#undef XLL_PARSE_JSON_VALUE
#endif // _DEBUG

} // namespace json::parse