// xll_parse.h - string parsing
#pragma once
#include "xll/xll/view.h"

namespace xll::parse {

	// eat c from front with checking
	template<class T>
	inline view<const T> eat(T c, view<const T> v)
	{
		ensure(v.front() == c);
		++v.buf;
		--v.len;

		return v;
	}
	// skip leading white space
	template<class T>
	inline view<const T> skipws(view<const T> v)
	{
		while (isspace(v.front())) {
			v = eat(v.front(), v);
		}

		return v;
	}

	// skip matching left and right chars ignoring escaped
	// "{data}..." returns escaped "data" an updates view to "..."
	template<class T>
	inline view<const T> skip(view<const T>& v, T l, T r, T e)
	{
		// delimiters can't be used as escape
		ensure(l != e and r != e);

		v = eat(l, v);
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

		auto data = view<const T>(v.buf, n - 1);
		v.buf += n;
		v.len -= n;

		return data;
	}

#ifdef _DEBUG

	template<class T>
	inline int test_skip()
	{
		auto tskip = [](view<const T> v) { return skip(v, _T('{'), _T('}'), _T('\\')); };

		ensure(tskip(_T("{a}")).equal(view(_T(""))));
		ensure(tskip(_T("{a}}")).equal(view(_T("}"))));
		ensure(tskip(_T("{a{}}bc")).equal(view(_T("bc"))));
		ensure(tskip(_T("{a\\}}b")).equal(view(_T("b"))));
		ensure(tskip(_T("{a\\{}b")).equal(view(_T("b"))));
		ensure(tskip(_T("{a\"b")).equal(view(_T("b"))));

		return 0;
	}

#endif // _DEBUG

	// find unescaped separator
	template<class T>
	inline DWORD find(xll::view<const T>& v, T s, T l, T r,T e)
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
				view<const T> vs(v.buf + n, (DWORD)(v.buf - n));
				n += skip<T>(vs, l, r, e);
			}
			else {
				++n;
			}
		};

		return n;
	}
		
	// next chunk and advance view 
	// "{data},..." returns "data" and updates view to "..."
	template<class T>
	inline xll::view<const T> chop(xll::view<const T>& v, T s, T l, T r, T e)
	{
		v = eat(l, v);
		DWORD n = find(v, s, l, r, e);
		ensure(v.buf[n] == r);
		auto data = view<const T>(v.buf, n - 1);
		ensure(data.back() == r);
		v.buf += n;
		v.len -= n;
		if (v) {
			v = eat(s, v);
		}

		return data;
	}

#ifdef _DEBUG

	template<class T>
	inline int test_chop()
	{
		const auto tchop = [](auto& v) {
			return chop<const T>(v, _T(':'), _T('\\'));
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
			xll::view<const T> v(_T("{a\\::b"));
			auto n = tchop(v);
			ensure(n.equal(xll::view<const T>(_T("{a\\:"))));
			ensure(v.equal(xll::view<const T>(_T("b"))));
		}
		{
			xll::view<const T> v(_T("{a{b}c}d"));
			auto n = tchop(v);
			ensure(n.equal(xll::view<const T>(_T("{a\\:"))));
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
				v.buf += n;
				v.len -= n;
				if (v) {
					v = eat(s, v);
				}
				n = find<const T>(v, s, l, r, e);
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
			xll::parse::iterator i(v, _T('\n'), 0, 0, _T('\\'));
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
			xll::parse::iterator is(v, _T('\n'), 0, 0, _T('\\'));
			DWORD n = 1;
			for (const auto& i : is) {
				ensure(i.equal(xll::view(a.buf, n)));
				++n;
			}
		}
		{
			xll::view v(_T("\"ab\"\ncd"));
			xll::parse::iterator i(v, _T('\n'), 0, 0, _T('\\'));
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
