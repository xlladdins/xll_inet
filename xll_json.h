// xll_json.h - JSON to and from OPER
#pragma once
#include <iosfwd>
#include "xll/xll/xll.h"
#include "xll_parse.h"

static inline const char xll_parse_json_doc[] = R"xyzyx(
The function <code>JSON.PARSE(string)</code> parses the JSON <code>string</code> into an <code>OPER</code>.
JSON objects map neatly to <code>OPER</code>s having two rows: 
the keys are in the first row and the values are in the second.
Keys are strings (<code>xltypeStr</code>) and values are objects, arrays, 
strings (<code>xltypeStr</code>, numbers (<code>xltypeNum</code),
<code>true</code> or <code>false</code> (<code>xltypeBool</code>),
or <code>null</code> (<code>#NULL!</code>).
Objects are zero or more key-value pairs of the form <code>{ "key" : value, ... }</code>
and correspond to multis (<code>xltypeMulti</code>) having 2 rows.
Arrays are zero or more values of the form <code>[ value, ... ]</code>
and correspond to multis having 1 row.
<p>
Multis are recursive but Excel cannot display multis that contain other multis.
Use <code>\RANGE(range)</code> to create an in-memory copy of the parsed string and
call <code>JSON.INDEX(handle, index)</code> to retreive values. The index is an array
of strings indicating the keys to be chosen. If <code>index</code> is a string containing
wildcards then all values with matching keys are returned.
)xyzyx";

namespace xll::json {

	enum class TYPE {
		Null = xltypeErr, // and val.err = xlerrNull 
		Number = xltypeNum,
		String = xltypeStr,
		Boolean = xltypeBool,
		Array = xltypeMulti,
		Object = xltypeMulti|xlbitXLFree,
		Unknown
	};

	template<class X>
	inline TYPE type(const XOPER<X>& x)
	{
		if (x.is_multi()) {
			if (x.rows() == 1) {
				return TYPE::Array;
			}
			else if (x.rows() == 2) {
				return TYPE::Object;
			}
		}
		else if (x.is_num()) {
			return TYPE::Number;
		}
		else if (x.is_str()) {
			return TYPE::String;
		}
		else if (x.is_bool()) {
			return TYPE::Boolean;
		}
		else if (x.is_err() and x.val.err == xlerrNull) {
			return TYPE::Null;
		}

		return TYPE::Unknown;
	}

	// JSON construction
	// [ v, ... ]
	template<class... Os>
	inline xll::OPER array(const Os&... os)
	{
		return xll::OPER({ xll::OPER(os)... }).resize(1, static_cast<unsigned>(sizeof...(os)));
	}
	template<>
	inline xll::OPER array()
	{
		return xll::OPER(1, 1);
	}

	// { k, ..., v, ...}
	template<class... Os>
	inline xll::OPER object(const Os&... os)
	{
		static_assert(0 == sizeof...(os) % 2);

		return array(os...).resize(2, static_cast<unsigned>(sizeof...(os)/2));
	}
	template<>
	inline xll::OPER object()
	{
		return xll::OPER(2, 1);
	}

	// "s" with escaped quotes
	inline OPER string(const OPER& s)
	{
		return OPER("\"") & Excel(xlfSubstitute, s, OPER("\""), OPER("\\\"")) & OPER("\"");
	}

	// lookup 0-based index of i in v
	template<class X>
	inline unsigned match(const XOPER<X>& v, const XOPER<X>& i)
	{
		unsigned vi = (unsigned)-1;

		if (i.is_num()) {
			vi = static_cast<unsigned>(i.val.num);
		}
		else if (i.is_str()) {
			X v1 = v; // first row of v
			v1.val.array.rows = 1;
			XOPER<X> vi0 = Excel(xlfMatch, i, v1, XOPER<X>(0)); // exact
			if (vi0.is_num()) {
				vi = static_cast<unsigned>(vi0.val.num - 1); // 0-based
			}
		}

		return vi;
	}
		
	// multi-level index into JSON value
	template<class X>
	inline const XOPER<X>& index(const XOPER<X>& v, XOPER<X> i)
	{
		auto vi0 = match(v, i[0]);
		ensure(vi0 != (unsigned)-1);

		return i.size() == 1 ? v(1, vi0) : index(v(1, vi0), i.drop(1));
	}

	// append contents of w to v.
	template<class X>
	inline XOPER<X>& concat(XOPER<X>& v, const XOPER<X>& w)
	{
		for (unsigned i = 0; i < w.columns(); ++i) {
			v.push_right(json::object<X>(w(0,i), w(1,i)));
		}

		return v;
	}

	template<class X, class T>
	inline std::basic_ostream<T>& operator<<(std::basic_ostream<T>& os, const xll::XOPER<X>& x)
	{
		switch (type(x)) {
		case TYPE::Array:
			os << '[';
			for (unsigned i = 0; i < x.size(); ++i) {
				if (i)
					os << ',';
				os << x[i];
			}
			os << ']';

			break;
		case TYPE::Object:
			os << '{';
			for (unsigned i = 0; i < x.columns(); ++i) {
				if (i)
					os << ',';
				ensure(x(0, i).is_str());
				os << x(0, i) << ':' << x(1, i);
			}
			os << '}';

			break;
		case TYPE::Null:
			os << 'n' << 'u' << 'l' << 'l';

			break;
		case TYPE::Boolean:
			if (x.val.xbool)
				os << 't' << 'r' << 'u' << 'e';
			else
				os << 'f' << 'a' << 'l' << 's' << 'e';

			break;
		case TYPE::Number:
			os << /*std::fixed <<*/ x.val.num;

			break;
		case TYPE::String:
			os << '"';
			os.write(x.val.str + 1, x.val.str[0]);
			os << '"';

			break;
		default:
			os << '?';
		}

		return os;
	}

	template<class X, class T = traits<X>::xchar>
	inline std::basic_string<T> stringify(const xll::XOPER<X>& x)
	{
		std::basic_ostringstream<T> oss;

		oss << x;

		return oss.str();
	}



#ifdef _DEBUG
#include <sstream>
	
	inline int test()
	{
		{
			auto o = json::array("a", 1, true);
			ensure(o.rows() == 1);
			ensure(o.columns() == 3);
			ensure(o[0] == "a");
			ensure(o[1] == 1);
			ensure(o[2] == true);
			auto s = json::stringify<XLOPERX, TCHAR>(o);
			ensure(s == _T("[\"a\",1,true]"));
		}
		{
			auto o = json::object("a", 1);
			ensure(o.rows() == 2);
			ensure(o.columns() == 1);
			ensure(o[0] == "a");
			ensure(o[1] == 1);
			auto s = json::stringify<XLOPERX, TCHAR>(o);
			ensure(s == _T("{\"a\":1}"));
		}
		{
			auto o = json::string(OPER("a\"b"));
			ensure(o.is_str());
			ensure(o == "\"a\\\"b\"");
		}
		{
			auto o = json::object("a", "b",
				                  json::object("b", 2), 1.23);
			ensure(o.rows() == 2);
			ensure(o.columns() == 2);
			auto i = json::array("a", "b");
 			ensure(index(o, i[0]) == json::object("b", 2));
			ensure(index(o, i[1]) == 1.23);
			ensure(index(o, i) == 2);
			auto s = json::stringify<XLOPERX, TCHAR>(o);
			ensure(s == _T("{\"a\":{\"b\":2},\"b\":1.23}"));
		}
		{
			auto o = json::object("a", "b",
				                  json::object("b", OPER::Err::Null), true);
			auto v = json::object("a", json::object("b", OPER::Err::Null));
			auto w = json::object("b", true);
			v = concat(v, w);
			ensure(v == o);
			auto s = json::stringify<XLOPERX, TCHAR>(o);
			ensure(s == _T("{\"a\":{\"b\":null},\"b\":true}"));
		}

		return 0;
	}

#endif // _DEBUG

} // xll::json



namespace xll::json::parse {

	// forward declaration
	template<class X, class T>
	XOPER<X> value(fms::view<const T>& v);

	template<class X, class T>
	inline XOPER<X> is_true(fms::view<const T>& v)
	{
		if (v.len >= 4
			and v.buf[0] == 't'
			and v.buf[1] == 'r'
			and v.buf[2] == 'u'
			and v.buf[3] == 'e') {
				v.drop(4);
				return XOPER<X>(true);
		}
	
		return XErrNA<X>;
	}

	template<class X, class T>
	inline XOPER<X> is_false(fms::view<const T>& v)
	{
		if (v.len >= 5
			and v.buf[0] == 'f'
			and v.buf[1] == 'a'
			and v.buf[2] == 'l'
			and v.buf[3] == 's'
			and v.buf[4] == 'e') {
				v.drop(5);
				return XOPER<X>(false);
		}

		return XErrNA<X>;
	}

	template<class X, class T>
	inline XOPER<X> is_null(fms::view<const T>& v)
	{
		if (v.len >= 4
			and v.buf[0] == 'n'
			and v.buf[1] == 'u'
			and v.buf[2] == 'l'
			and v.buf[3] == 'l') {
				v.drop(4);
				return XErrNull<X>;
		}

		return XErrNA<X>;
	}

	template<class T>
	struct str { };
	template<>
	struct str<char> {
		static double tod(const char* s, char** e) { return strtod(s, e); }
	};
	template<>
	struct str<wchar_t> {
		static double tod(const wchar_t* s, wchar_t** e) { return wcstod(s, e); }
	};

	template<class X, class T>
	inline XOPER<X> number(fms::view<const T>& v)
	{
		XOPER<X> num = XErrNA<X>;

		T* e;
		double n = str<T>::tod(v.buf, &e);
		if (v.buf != e) {
			num = n;
			v.drop(static_cast<uint32_t>(e - v.buf));
		}
		
		return num;
	}

	// "\"str\"" => "str"
	template<class X, class T = typename traits<X>::xchar>
	inline XOPER<X> string(fms::view<const T>& v)
	{
		auto str = skip<T>(v, '"', '"', '\\');

		return XOPER<X>(str.buf, str.len);
	}

	// object := "{ \"key\" : val , ... } "
	template<class X, class T = typename traits<X>::xchar>
	inline XOPER<X> object(fms::view<const T>& v)
	{
		XOPER<X> x;

		v.skipws().eat('{');
		while (v.skipws()) {
			auto key = string<X,T>(v);
			v.skipws();
			v.eat(':');
			x = concat(x, json::object<X>(key, value<X,T>(v)));
			v.skipws();
			if (v.front() == ',') {
				v.eat(',');
			}
			else {
				v.eat('}');
				break;
			}
		}

		return x;
	}

	// array := "[ value , ... ] "
	template<class X, class T = typename traits<X>::xchar>
	inline XOPER<X> array(fms::view<const T>& v)
	{
		XOPER<X> x;

		v.skipws().eat('[');
		while (v.skipws() and v.front() != ']') {
			auto val = value<X, T>(v);
			if (val.is_multi()) {
				OPER xi(1, 1);
				xi[0] = val;
				x.push_right(xi);
			}
			else {
				x.push_right(val);
			}
			v.skipws();
			if (v.front() == ',') {
				v.eat(',');
			}
			else {
				v.eat(']');
				break;
			}
		}

		x.resize(1, x.size());

		return x;
	}

	template<class X, class T = typename traits<X>::xchar>
	inline XOPER<X> value(fms::view<const T>& v)
	{
		v.skipws();

		if (v.front() == '{') {
			return object<X,T>(v);
		}
		else if (v.front() == '[') {
			return array<X,T>(v);
		}
		else if (v.front() == '"') {
			return string<X,T>(v);
		}

		
		if (auto val = is_null<X>(v); val != XErrNA<X>) return val;
		if (auto val = is_true<X>(v); val != XErrNA<X>) return val;
		if (auto val = is_false<X>(v); val != XErrNA<X>) return val;

		// last chance
		XOPER<X> val = number<X>(v);
		ensure(val != XErrNA<X>);

		return val;
	}

	template<class X, class T = typename traits<X>::xchar>
	inline XOPER<X> view(fms::view<const T> v)
	{
		XOPER<X> x = parse::value<X, T>(v.skipws());
		ensure(!v.skipws());

		return x;
	}


#ifdef _DEBUG


#define XLL_PARSE_JSON_VALUE(X) \
	X("null", XErrNull<XLOPERX>) \
	X("\"str\"", "str") \
	X("\"s\\\"r\"", "s\\\"r") \
	X("\"s\nr\"", "s\nr") \
	X("\"s r\"", "s r") \
	X("[\"a\", 1.23, false]", json::array("a", 1.23, false)) \
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
#define PARSE_JSON_CHECK(a, b) { ensure(parse::view<XLOPERX>(fms::view(_T(a))) == b); }
		XLL_PARSE_JSON_VALUE(PARSE_JSON_CHECK)
#undef PARSE_JSON_CHECK

		{
			fms::view t(_T("true"));
			ensure(is_true<XLOPERX>(t) == true);
			ensure(!t);

			fms::view f(_T("false"));
			ensure(is_false<XLOPERX>(f) == false);
			ensure(!f);

			fms::view n(_T("null"));
			ensure(is_null<XLOPERX>(n) == ErrNull);
			ensure(!n);

			{
				fms::view num(_T("1.23foo"));
				ensure(number<XLOPERX>(num) == 1.23);
				ensure(num.equal(_T("foo")));
			}
			{
				fms::view num(_T("foo1.23"));
				ensure(number<XLOPERX>(num) == ErrNA);
			}
		}
		{
			fms::view str("\"str\"");
			ensure(parse::string<XLOPERX>(str) == "str");
		}
		{
			fms::view str("\"s\\\"r\"");
			ensure(parse::string<XLOPERX>(str) == "s\\\"r");
			ensure(!str);
		}
		{
			OPER x;
			x = parse::view<XLOPERX>(fms::view(_T("{\"a\":{\"b\":\"cd\"}}")));
			OPER i({ OPER("a"), OPER("b") });
			ensure(json::index(x, i) == "cd");
			ensure(json::index(x, i[0]) == OPER({ OPER("b"), OPER("cd") }).resize(2, 1));
		}
		{

		}

		// https://www.wikidata.org/w/api.php?action=wbgetentities&sites=enwiki&titles=Berlin&props=descriptions&languages=en&format=json
		const char wd[] = "{\"entities\":"
			"{\"Q64\":"
				"{\"type\":\"item\",\"id\":\"Q64\",\"descriptions\":"
					"{\"en\":{\"language\":\"en\",\"value\":\"federal state, capitaland largest city of Germany\"}}"
			    "}"
			"},\"success\":1}";
		OPER x;
		x = parse::view<XLOPERX, char>(fms::view<const char>(wd));

		return 0;
	}
#undef XLL_PARSE_JSON_VALUE
#endif // _DEBUG

} // namespace json::parse