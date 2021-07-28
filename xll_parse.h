// xll_parse.h - string parsing
#pragma once
#include "xll/xll/fms_view.h"

// phony xltypes
#define xltypeDate (xltypeNum|xlbitXLFree)
#define xltypeTime (xltypeNum|xlbitDLLFree)

#define XLTYPE(X) \
	X(xltypeNum, _NUM, "is a 64-bit floating point number") \
	X(xltypeStr, _STR, "is a string") \
	X(xltypeBool, _BOOL, "is a boolean") \
	X(xltypeRef, _REF, "is a multiple reference") \
	X(xltypeErr, _ERR, "is an error") \
	X(xltypeMulti, _MULTI, "is a two dimensional range") \
	X(xltypeMissing, _MISSING, "is missing function argument") \
	X(xltypeNil, _NIL, "is a null type") \
	X(xltypeSRef, _SREF, "is a single reference") \
	X(xltypeInt, _INT, "is an integer") \
	X(xltypeDate, _DATE, "is a datetime") \
	X(xltypeTime, _TIME, "is a time") \

namespace xll::parse {

	template<class T>
	inline double to_number(fms::view<T>& v);

	// test for numbers
	template<>
	inline double to_number<char>(fms::view<char>& v)
	{
		char* end;
		double num = strtod(v.buf, &end);
		if (end == v.buf) 
			return std::numeric_limits<double>::quiet_NaN();
		v.drop(static_cast<int32_t>(end - v.buf));

		return num;
	}
	template<>
	inline double to_number<const char>(fms::view<const char>& v)
	{
		char* end = const_cast<char*>(v.buf) + v.len;
		double num = strtod(v.buf, &end);
		if (end == v.buf)
			return std::numeric_limits<double>::quiet_NaN();
		v.drop(static_cast<int32_t>(end - v.buf));

		return num;
	}
	template<>
	inline double to_number<wchar_t>(fms::view<wchar_t>& v)
	{
		wchar_t* end = const_cast<wchar_t*>(v.buf) + v.len;
		double num = wcstod(v.buf, &end);
		if (end == v.buf)
			return std::numeric_limits<double>::quiet_NaN();
		v.drop(static_cast<int32_t>(end - v.buf));

		return num;
	}
	template<>
	inline double to_number<const wchar_t>(fms::view<const wchar_t>& v)
	{
		wchar_t* end;
		double num = wcstod(v.buf, &end);
		if (end == v.buf)
			return std::numeric_limits<double>::quiet_NaN();
		v.drop(static_cast<int32_t>(end - v.buf));

		return num;
	}

	template<class T>
	inline double to_number(const fms::view<T>& v)
	{
		fms::view<T> v_(v);

		double num = to_number<T>(v_);

		return !v_.skipws() ? num : std::numeric_limits<double>::quiet_NaN();
	}

	template<class X>
	inline void convert(XOPER<X>& o, int type)
	{
		if (type == xltypeNum || type == xltypeBool || type == xltypeInt) {
			o = Excel(xlfEvaluate, o);
			if (type == xltypeBool and !o.is_bool()) {
				o = !!o;
				ensure(o.is_bool());
			}
			else if (type == xltypeInt) {
				o.val.w = static_cast<traits<X>::xint>(o.val.num);
				o.xltype = xltypeInt;
			}
		}
		else if (type == xltypeDate) {
			o = Excel(xlfDatevalue, o);
			o.xltype = xltypeDate;
		}
		else if (type == xltypeTime) {
			o = Excel(xlfTimevalue, o);
			o.xltype = xltypeTime;
		}
	}
	template<class X>
	inline void convert(XOPER<X>& o, const XOPER<X>& type)
	{
		if (type.is_num() and type.as_num() != 0) {
			convert(o, static_cast<int>(type.as_num()));
		}
	}

	// skip matching left and right chars ignoring escaped
	// "{da\}ta}..." returns escaped "da\}ta" an updates fms::view to "..."
	template<class T>
	inline fms::view<T> skip(fms::view<T>& v, T l, T r, T e)
	{
		// delimiters can't be used as escape
		ensure(l != e and r != e);

		v.eat(l);
		int level = 1;

		DWORD n = 0;
		while (level and n < v.len) {
			if (v.buf[n] == e) {
				++n;
				ensure(n < v.len);
				++n;
				ensure(n < v.len);
			}
			// do first in case l == r
			if (v.buf[n] == r) {
				--level;
			}
			else if (v.buf[n] == l) {
				++level;
			}
			++n;
		};

		ensure(level == 0);

		auto data = fms::view<T>(v.buf, n - 1);
		v.drop(n);

		return data;
	}

#ifdef _DEBUG

	template<class T>
	inline int test_skip()
	{
		{
			auto tskip = [](fms::view<const T> v) { return skip(v, _T('{'), _T('}'), _T('\\')); };

			ensure(tskip(_T("{a}")).equal(fms::view(_T(""))));
			ensure(tskip(_T("{a}}")).equal(fms::view(_T("}"))));
			ensure(tskip(_T("{a{}}bc")).equal(fms::view(_T("bc"))));
			ensure(tskip(_T("{a\\}}b")).equal(fms::view(_T("b"))));
			ensure(tskip(_T("{a\\{}b")).equal(fms::view(_T("b"))));
		}
		{
			auto tskip = [](fms::view<const T> v) { return skip(v, _T('|'), _T('|'), 0); };

			ensure(tskip(_T("|a|")).equal(fms::view(_T(""))));
			ensure(tskip(_T("|a||")).equal(fms::view(_T("|"))));
			ensure(tskip(_T("|a|||bc")).equal(fms::view(_T("bc"))));
			ensure(tskip(_T("|a\\||b")).equal(fms::view(_T("|b"))));

		}

		return 0;
	}

#endif // _DEBUG

	// find index of unescaped separator
	template<class T>
	inline DWORD find(const fms::view<T>& v, T s, T l, T r, T e)
	{
		DWORD n = 0;

		while (n < v.len and v.buf[n] != s) {
			if (v.buf[n] == e) {
				++n;
				ensure(n < v.len);
				++n;
				ensure(n < v.len);
			}
			if (v.buf[n] == l) {
				auto vn = fms::view(v.buf + n, static_cast<DWORD>(v.len - n));
				n += skip<T>(vn, l, r, e).len + 2; // count l and r
			}
			else {
				++n;
			}
		};

		return n;
	}
		
	// next chunk up to separator and advance view 
	template<class T>
	inline fms::view<T> chop(fms::view<T>& v, T s, T l, T r, T e)
	{
		DWORD n = find(v, s, l, r, e);
		auto data = fms::view<T>(v.buf, n);
		v.drop(n);
		if (v) {
			v.eat(s);
		}

		return data;
	}

#ifdef _DEBUG

	template<class T>
	inline int test_chop()
	{
		const auto tchop = [](auto& v) {
			return chop<const T>(v, _T(':'), _T('{'), _T('}'), _T('\\'));
		};
		{
			fms::view<const T> v(_T("a:b"));
			auto n = tchop(v);
			ensure(n.equal(fms::view<const T>(_T("a"))));
			ensure(v.equal(fms::view<const T>(_T("b"))));
			n = tchop(v);
			ensure(n.equal(fms::view<const T>(_T("b"))));
			ensure(v.equal(fms::view<const T>(_T(""))));
			n = tchop(v);
			ensure(n.equal(fms::view<const T>(_T(""))));
			ensure(v.equal(fms::view<const T>(_T(""))));
		}
		{
			fms::view<const T> v(_T(":a:b"));
			auto n = tchop(v);
			ensure(n.equal(fms::view<const T>(_T(""))));
			ensure(v.equal(fms::view<const T>(_T("a:b"))));
			n = tchop(v);
			ensure(n.equal(fms::view<const T>(_T("a"))));
			ensure(v.equal(fms::view<const T>(_T("b"))));
			n = tchop(v);
			ensure(n.equal(fms::view<const T>(_T("b"))));
			ensure(v.equal(fms::view<const T>(_T(""))));
		}
		{
			fms::view<const T> v(_T("{a\\::b"));
			auto n = tchop(v);
			ensure(n.equal(fms::view<const T>(_T("{a\\:"))));
			ensure(v.equal(fms::view<const T>(_T("b"))));
		}
		{
			fms::view<const T> v(_T("{a{b}c}d"));
			auto n = tchop(v);
			ensure(n.equal(fms::view<const T>(_T("{a\\:"))));
			ensure(v.equal(fms::view<const T>(_T("b"))));
		}


		return 0;
	}

#endif // _DEBUG

	// view iterator
	template<class T>
	class iterator {
		fms::view<T> v;
		T s, l, r, e;
		DWORD n;
	public:
		using iterator_category = std::forward_iterator_tag;
		using value_type = fms::view<T>;
		iterator()
			: s(0), l(0), r(0), e(0), n(0)
		{ }
		iterator(const fms::view<T>& _v, T s, T l, T r, T e)
			: v(_v), s(s), l(l), r(r), e(e), n(find<T>(v, s, l, r, e))
		{ }
		~iterator()
		{ }

		bool operator==(const iterator& i) const = default;

		auto begin() const
		{
			return *this;
		}
		auto end() const
		{
			return iterator(fms::view<T>(v.buf + v.len, 0), s, l, r, e);
		}
		value_type operator*() const
		{
			return fms::view<T>(v.buf, n);
		}
		iterator& operator++()
		{
			if (v) {
				v.drop(n);
				if (v) {
					v.eat(s);
				}
				n = find<T>(v, s, l, r, e); // n = chop().len
			}

			return *this;
		}
		iterator& operator++(int)
		{
			auto tmp{ *this };

			operator++();

			return tmp;
		}

	};

#ifdef _DEBUG

	template<class T>
	inline int test_iterator()
	{
		{
			fms::view v(_T("ab\ncd"));
			parse::iterator i(v, _T('\n'), 0, 0, _T('\\'));
			auto b = i.begin();
			ensure((*b).equal(fms::view(_T("ab"))));
			++b;
			ensure((*b).equal(fms::view(_T("cd"))));
			++b;
			auto e = i.end();
			ensure(b == e);
		}
		{
			fms::view v(_T("a\naa\naaa"));
			fms::view a(_T("aaa"));
			parse::iterator is(v, _T('\n'), 0, 0, _T('\\'));
			DWORD n = 1;
			for (const auto& i : is) {
				ensure(i.equal(fms::view(a.buf, n)));
				++n;
			}
		}
		{
			fms::view v(_T("\"ab\"\ncd"));
			parse::iterator i(v, _T('\n'), 0, 0, _T('\\'));
			auto b = i.begin();
			ensure((*b).equal(fms::view(_T("\"ab\""))));
			++b;
			ensure((*b).equal(fms::view(_T("cd"))));
			++b;
			ensure(b == i.end());
		}

		return 0;
	}

#endif // _DEBUG

#ifdef _DEBUG

	template<class T>
	inline int test()
	{
		test_skip<TCHAR>();
		test_chop<TCHAR>();
		test_iterator<TCHAR>();

		return 0;
	}

#endif // _DEBUG

} // xll
