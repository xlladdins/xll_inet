// xll_parse.h - string parsing
#pragma once
#include "xll/xll/view.h"

namespace xll::parse {

	// skip matching left and right chars ignoring escaped
	template<class T>
	inline const T* skip(const T* s, T l, T r, T esc)
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
		// delimiters must not equal escape character
		ensure(nullptr == skip(_T(""), _T('\\'), _T('}'), _T('\\')));
		ensure(nullptr == skip(_T(""), _T('{'), _T('\\'), _T('\\')));

		auto tskip = [](const T* s) { return skip(s, _T('{'), _T('}'), _T('\\')); };

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
	inline DWORD find(const xll::view<const T>& v, T s, T l, T r, T e)
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
	inline xll::view<const T> chop(xll::view<const T>& v, T s, T l, T r, T e)
	{
		DWORD n = find(v, s, l, r, e);
		xll::view<const T> v0(v.buf, n);
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
		xll::view<const T> v;
		T s, l, r, e;
		DWORD n;
	public:
		using iterator_category = std::forward_iterator_tag;
		using value_type = view<const T>;
		iterator()
			: s(0), l(0), r(0), e(0), n(0)
		{ }
		iterator(const xll::view<const T>& _v, T s, T l, T r, T e)
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
			return iterator(xll::view<const T>(v.buf + v.len, 0), s, l, r, e);
		}
		value_type operator*() const
		{
			return xll::view<const T>(v.buf, n);
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
			xll::parse::iterator i(v, _T('\n'), _T('"'), _T('"'), _T('\\'));
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
			xll::parse::iterator is(v, _T('\n'), _T('"'), _T('"'), _T('\\'));
			DWORD n = 1;
			for (const auto& i : is) {
				ensure(i.equal(xll::view(a.buf, n)));
				++n;
			}
		}
		{
			xll::view v(_T("\"ab\"\ncd"));
			xll::parse::iterator i(v, _T('\n'), _T('"'), _T('"'), _T('\\'));
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

} // xll::parse
