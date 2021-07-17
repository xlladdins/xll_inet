// xll_json.h - JSON to and from OPER
#pragma once
#include <iosfwd>
#include "xll/xll/xll.h"
#include "xll_csv.h"

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

	// first row of o
	template<class X>
	const X keys(const XOPER<X>& o)
	{
		ensure(json::type(o) == json::TYPE::Object);

		X k = o;
		k.val.array.rows = 1;

		return k;
	}

	// 0-based index of k in first row of o
	template<class X>
	inline unsigned match(const XOPER<X>& o, const XOPER<X>& k)
	{
		auto i = Excel(xlfMatch, k, keys(o), XOPER<X>(0)); // exact
		ensure(i.is_num());

		return static_cast<unsigned>(i.val.num - 1);
	}

	// convert jq dotted index to array
	// array indices do not have [brackets]
	template<class X, class T = xll::traits<X>::xchar>
	inline XOPER<X> to_index(const XOPER<X>& i)
	{
		ensure(i.is_str());
		ensure(i.val.str[0] and i.val.str[1] == '.');

		auto v = view<T>(i.val.str + 2, i.val.str[0] - 1);

		XOPER<X> is;
		for (const auto& i_ : parse::iterator<T>(v, T('.'), T(0), T(0), T(0))) {
			XOPER<X> _i(i_.buf, i_.len);
			if (_i.val.str[0] and isdigit(_i.val.str[1])) {
				is.push_right(Excel(xlfValue, _i)); // might be #VALUE!
			}
			else {
				is.push_right(_i); // might be ""
			}
		}

		return is;
	}

	// convert array of keys to a jq dotted key
	template<class X, class T = xll::traits<X>::xchar>
	inline XOPER<X> from_index(const XOPER<X>& is)
	{
		static XOPER<X> z("0");
		static XOPER<X> dot(".");
		XOPER<X> index;

		for (auto i : is) {
			if (i.is_str()) {
				index.append(dot).append(i);
			}
			else if (i.is_num()) {
				ensure(i.as_num() >= 0);
				ensure(std::trunc(i.val.num) == i.val.num);

				index.append(dot).append(Excel(xlfText, i, z));
			}
			// else error???
		}

		return index;
	}

	// index into JSON values
	template<class X>
	const XOPER<X>& index(const XOPER<X>& o, XOPER<X> k)
	{
		ensure(!k.is_err());
		ensure(k.rows() == 1 or k.columns() == 1);

		auto type = json::type(o);

		// jq dotted index
		if (k.is_str() and k.val.str[0] and k.val.str[1] == '.') {
			k = to_index(k);
		}

		auto k0 = k[0];
		if (k.size() == 1) {
			if (type == json::TYPE::Array) {
				ensure(k0.is_num());
				ensure(k0.as_num() < o.size());

				return o[static_cast<unsigned>(k0.val.num)];
			}
			else if (type == json::TYPE::Object) {
				ensure(k0.is_str());

				return o(1, match(o, k0));
			}
			else {
				ensure(k0 == 0);

				return o;
			}
		}
		k.drop(1);

		return index(index(o, k0), k);
	}


#ifdef _DEBUG

	template<class X>
	inline int index_test()
	{
		{
			XOPER<X> i(".a.b.c");
			auto is = to_index(i);
			auto i_ = from_index(is);
			ensure(i == i_);
		}
		{
			XOPER<X> i(".a.0.1");
			auto is = to_index(i);
			ensure(is.size() == 3);
			ensure(is[1] == 0);
			ensure(is[2] == 1);
			auto i_ = from_index(is);
			ensure(i == i_);
		}

		return 0;
	}

#endif // _DEBUG

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
 			ensure(json::index(o, i[0]) == json::object("b", 2));
			ensure(json::index(o, OPER(".a")) == json::object("b", 2));
			ensure(json::index(o, i[1]) == 1.23);
			ensure(json::index(o, OPER(".b")) == 1.23);
			ensure(json::index(o, i) == 2);
			ensure(json::index(o, OPER(".a.b")) == 2);
			auto s = json::stringify<XLOPERX, TCHAR>(o);
			ensure(s == _T("{\"a\":{\"b\":2},\"b\":1.23}"));
		}
		{
			auto o = json::object("a", "b",
				                  json::object("b", OPER::Err::Null), true);
			auto v = json::object("a", json::object("b", OPER::Err::Null));
			auto w = json::object("b", true);
			v.push_right(w);
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
	XOPER<X> value(fms::view<T>& v);

	template<class X, class T>
	inline XOPER<X> is_true(fms::view<T>& v)
	{
		if (v.len >= 4
			and v.buf[0] == 't'
			and v.buf[1] == 'r'
			and v.buf[2] == 'u'
			and v.buf[3] == 'e') {
				v.drop(4);
				return XOPER<X>(true);
		}
	
		return XErrValue<X>;
	}

	template<class X, class T>
	inline XOPER<X> is_false(fms::view<T>& v)
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

		return XErrValue<X>;
	}

	template<class X, class T>
	inline XOPER<X> is_null(fms::view<T>& v)
	{
		if (v.len >= 4
			and v.buf[0] == 'n'
			and v.buf[1] == 'u'
			and v.buf[2] == 'l'
			and v.buf[3] == 'l') {
				v.drop(4);
				return XErrNull<X>;
		}

		return XErrValue<X>;
	}

	template<class X, class T>
	inline XOPER<X> number(fms::view<T>& v)
	{
		double num = xll::parse::to_number<T>(v);

		return isnan(num) ? XErrValue<X> : XOPER<X>(num);
	}

	// "\"str\"" => "str"
	template<class X, class T = typename traits<X>::xchar>
	inline XOPER<X> string(fms::view<T>& v)
	{
		auto str = xll::parse::skip<T>(v, '"', '"', '\\');

		return XOPER<X>(str.buf, str.len);
	}

	// object := "{ \"key\" : val , ... } "
	template<class X, class T = typename traits<X>::xchar>
	inline XOPER<X> object(fms::view<T>& v)
	{
		XOPER<X> x;

		v.skipws().eat('{');
		while (v.skipws()) {
			auto key = string<X,T>(v);
			v.skipws();
			v.eat(':');
			x.push_right(json::object<X>(key, value<X,T>(v)));
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
	inline XOPER<X> array(fms::view<T>& v)
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
	inline XOPER<X> value(fms::view<T>& v)
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

		
		if (auto val = is_null<X>(v); val != XErrValue<X>) return val;
		if (auto val = is_true<X>(v); val != XErrValue<X>) return val;
		if (auto val = is_false<X>(v); val != XErrValue<X>) return val;

		// last chance
		XOPER<X> val = number<X>(v);
		ensure(val != XErrValue<X>);

		return val;
	}

	template<class X, class T = typename traits<X>::xchar>
	inline XOPER<X> view(fms::view<T> v)
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
				ensure(number<XLOPERX>(num) == ErrValue);
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
		x = parse::view<XLOPERX, const char>(fms::view(wd));

		return 0;
	}
#undef XLL_PARSE_JSON_VALUE
#endif // _DEBUG

} // namespace json::parse