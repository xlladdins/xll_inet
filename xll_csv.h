// xll_csv.h - Comma Separated Value
#pragma once
#include <compare>
#include <iterator>
#include "ISO8601.h"
#include "xll/xll/xll.h"

#ifndef CATEGORY
#define CATEGORY "XLL"
#endif

using xchar = typename xll::traits<XLOPERX>::xchar;
using xcstr = typename xll::traits<XLOPERX>::xcstr;

namespace xll::csv {

	// skip matching left and right chars ignoring escaped
	inline xcstr skip(xcstr s, xchar l, xchar r, xchar esc)
	{
		if (l == esc or r == esc) {
			return nullptr; 
		}
		if (!s or *s != l) {
			return s;
		}

		int level = 1;
		while (level and *++s) {
			if (*s == esc) {
				++s;
				if (!*s) {
					return nullptr;
				}
				++s; // always skip escaped char
			}
			if (*s == r) {
				--level;
			}
			else if (*s == l) {
				++level;
			}
		};

		return level == 0 ? s + (*s == r) : nullptr;
	}

#ifdef _DEBUG

	template<class T>
	inline int test_skip()
	{
		auto tskip = [](const TCHAR* s) { return skip(s, _T('{'), _T('}'), _T('\\')); };

		LPCTSTR s;
		s = tskip(nullptr);
		ensure(s == nullptr);

		s = tskip(_T(""));
		ensure(s && *s == 0);

		s = tskip(_T("a"));
		ensure(s && *s == _T('a'));

		s = tskip(_T("{a}"));
		ensure(s && *s == 0);

		s = tskip(_T("{a}}"));
		ensure(s && *s == _T('}'));

		// '{' not matched
		s = tskip(_T("{a{}"));
		ensure(!s);

		s = tskip(_T("{a{}}b"));
		ensure(s && *s == _T('b'));

		// always skip escape
		s = tskip(_T("{a\\}{}}b"));
		ensure(s && *s == _T('b'));
		s = tskip(_T("{a\\}{}}b"));
		ensure(s && *s == _T('b'));

		return 0;
	}

#endif // _DEBUG


	// offset to separator 
	template<class T>
	inline DWORD find(const xll::view<T>& v, xchar s, xchar l, xchar r, xchar e)
	{
		DWORD n = 0;
		while (n < v.len and v.buf[n] and v.buf[n] != s) {
			if (v.buf[n] == l) {
				n = static_cast<DWORD>(skip(v.buf + n, l, r, e) - v.buf);
			}
			else {
				++n;
			}
		}

		return n;
	}

	// next chunk and advance view 
	template<class T>
	inline xll::view<T> chop(xll::view<T>& v, xchar s, xchar l, xchar r, xchar e)
	{
		DWORD n = find(v, s, l, r, e);
		xll::view<T> v0(v.buf, n);
		v.buf += n;
		v.len -= n;
		if (v.len) {
			ensure(v.buf[0] == s);
			++v.buf;
			--v.len;
		}

		return v0;
	}


#ifdef _DEBUG

	template<class T>
	inline int test_chop()
	{
		const auto tchop = [](auto& v) { 
			return chop<const T>(v, _T(':'), _T('{'), _T('}'), _T('\\')); 
		};
		{
			xll::view<const T> v(_T("a:b"));
			auto n = tchop(v);
			ensure(n.equal(xll::view<const T>(_T("a"))));
			ensure(v.equal(xll::view<const T>(_T("b"))));
			n = tchop(v);
			ensure(n.equal(xll::view<const T>(_T("b"))));
			ensure(v.equal(xll::view<const T>(_T(""))));
			n = tchop(v);
			ensure(n.equal(xll::view<const T>(_T(""))));
			ensure(v.equal(xll::view<const T>(_T(""))));
		}
		{
			xll::view<const T> v(_T(":a:b"));
			auto n = tchop(v);
			ensure(n.equal(xll::view<const T>(_T(""))));
			ensure(v.equal(xll::view<const T>(_T("a:b"))));
			n = tchop(v);
			ensure(n.equal(xll::view<const T>(_T("a"))));
			ensure(v.equal(xll::view<const T>(_T("b"))));
			n = tchop(v);
			ensure(n.equal(xll::view<const T>(_T("b"))));
			ensure(v.equal(xll::view<const T>(_T(""))));
		}
		{
			xll::view<const T> v(_T("{a:b}:b"));
			auto n = tchop(v);
			ensure(n.equal(xll::view<const T>(_T("{a:b}"))));
			ensure(v.equal(xll::view<const T>(_T("b"))));
		}
		{
			xll::view<const T> v(_T("{a\\{b}:b"));
			auto n = tchop(v);
			ensure(n.equal(xll::view<const T>(_T("{a\\{b}"))));
			ensure(v.equal(xll::view<const T>(_T("b"))));
		}
		{
			xll::view<const T> v(_T("{a\\}b}:b"));
			auto n = tchop(v);
			ensure(n.equal(xll::view<const T>(_T("{a\\}b}"))));
			ensure(v.equal(xll::view<const T>(_T("b"))));
		}
		{
			xll::view<const T> v(_T("{a{\\}b}c}:b"));
			auto n = tchop(v);
			ensure(n.equal(xll::view<const T>(_T("{a{\\}b}c}"))));
			ensure(v.equal(xll::view<const T>(_T("b"))));
		}


		return 0;
	}

#endif // _DEBUG

	// view iterator
	template<class T>
	class iterator {
		xll::view<T> v;
		xchar s, l, r, e;
		DWORD n;
	public:
		using iterator_category = std::forward_iterator_tag;
		using value_type = view<T>;
		iterator()
			: s(0), l(0), r(0), e(0), n(0)
		{ }
		iterator(const xll::view<T>& _v, xchar s, xchar l, xchar r, xchar e)
			: v(_v), s(s), l(l), r(r), e(e), n(find(v, s, l, r, e))
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
			return iterator(xll::view<T>(v.buf + v.len, 0), s, l, r, e);
		}
		value_type operator*() const
		{
			return xll::view<T>(v.buf, n);
		}
		iterator& operator++()
		{
			if (v) {
				chop(v, s, l, r, e);
				n = find(v, s, l, r, e);
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
			xll::view v(_T("ab\ncd"));
			xll::csv::iterator i(v, '\n', '"', '"', '\\');
			auto b = i.begin();
			ensure((*b).equal(xll::view(_T("ab"))));
			++b;
			ensure((*b).equal(xll::view(_T("cd"))));
			++b;
			auto e = i.end();
			ensure(b == e);
		}
		{
			xll::view v(_T("a\naa\naaa"));
			xll::view a(_T("aaa"));
			xll::csv::iterator is(v, '\n', '"', '"', '\\');
			DWORD n = 1;
			for (const auto& i : is) {
				ensure(i.equal(xll::view(a.buf, n)));
				++n;
			}
		}
		{
			xll::view v(_T("\"ab\"\ncd"));
			xll::csv::iterator i(v, '\n', '"', '"', '\\');
			auto b = i.begin();
			ensure((*b).equal(xll::view(_T("\"ab\""))));
			++b;
			ensure((*b).equal(xll::view(_T("cd"))));
			++b;
			ensure(b == i.end());
		}

		return 0;
	}

#endif // _DEBUG

	//!!! use xll::codec
	template<class T>
	inline OPER parse_field(view<T> field)
	{
		static OPER t("TRUE"), f("FALSE");

		OPER o(field.buf, static_cast<xchar>(field.len));
		o = Excel(xlfTrim, o);

		// case insensitive
		if (o == t)
			return OPER(true);
		if (o == f)
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
		ensure(parse_field(view(_T("tRue"))) == OPER(true));

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
	inline int test_parse()
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
			xll::view v(_T("a,true\n1.23,2001-2-3"));
			OPER o = parse(v.buf, v.len, _T('\n'), _T(','), _T('\"'), _T('\"'), _T('\\'));
			ensure(o.rows() == 2);
			ensure(o.columns() == 2);
			ensure(o(0, 0) == "a");
			ensure(o(0, 1) == true);
			ensure(o(1, 0) == 1.23);
			ensure(o(1, 1) == Excel(xlfDate, OPER(2001), OPER(2), OPER(3)));
		}

		return 0;
	}

	template<class T>
	inline int test()
	{
		test_skip<TCHAR>();
		test_chop<TCHAR>();
		test_iterator<TCHAR>();
		test_parse<TCHAR>();

		return 0;
	}

#endif // _DEBUG

} // namespace xll::csv


