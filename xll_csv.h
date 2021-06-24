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
		if (!s) {
			return s;
		}
		if (*s != l) {
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

		return 0;
	}

#endif // _DEBUG


	// offset to separator 
	template<class T>
	inline DWORD next(const xll::view<T>& v, xchar s, xchar l, xchar r, xchar e)
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

#ifdef _DEBUG

	template<class T>
	inline int test_next()
	{
		{
			xll::view<const T> v(_T("a:b"));
			DWORD n = next<const T>(v, _T(':'), _T('{'), _T('}'), _T('\\'));
			ensure(n == 1);
		}
		{
			xll::view<const T> v(_T(":a:b"));
			DWORD n = next<const T>(v, _T(':'), _T('{'), _T('}'), _T('\\'));
			ensure(n == 0);
		}

		return 0;
	}

#endif // _DEBUG

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
		iterator(const xll::view<T>& v, xchar s, xchar l, xchar r, xchar e)
			: v(v), s(s), l(l), r(r), e(e), n(next(v, s, l, r, e))
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
			return iterator(xll::view<T>{}, s, l, r, e);
		}
		value_type operator*() const
		{
			return n ? xll::trim<T>(xll::view<T>(v.buf, n), _T(' ')) : xll::view<T>{};
		}
		iterator& operator++()
		{
			if (v) {
				v.buf += n;
				v.len -= n;
				if (v.len and v.buf[0] == s) {
					++v.buf;
					--v.len;
				}
				n = next(v, s, l, r, e);
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
			xll::csv::iterator i(v, '\n', '"', '"', '"');
			auto b = i.begin();
			ensure(*b == xll::view(_T("ab")));
			++b;
			ensure(*b == xll::view(_T("cd")));
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

	inline OPER parse(TCHAR* buf, DWORD len)
	{
		OPER o;

		view<TCHAR> lines(buf, len);

		return o;
	}

#ifdef _DEBUG

	template<class T>
	inline int test()
	{
		test_skip<TCHAR>();
		test_next<TCHAR>();
		test_iterator<TCHAR>();

		return 0;
	}

#endif // _DEBUG

} // namespace xll::csv


