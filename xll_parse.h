// xll_parse.h - string parsing
#pragma once
#include "xll/xll/fms_view.h"

namespace xll::parse {

	// eat c from front with checking
	template<class T>
	inline fms::view<const T> eat(T c, fms::view<const T> v)
	{
		ensure(v.front() == c);
		++v.buf;
		--v.len;

		return v;
	}
	// skip leading white space
	template<class T>
	inline fms::view<const T> skipws(fms::view<const T> v)
	{
		while (isspace(v.front())) {
			v = eat(v.front(), v);
		}

		return v;
	}

	// skip matching left and right chars ignoring escaped
	// "{data}..." returns escaped "data" an updates fms::view to "..."
	template<class T>
	inline fms::view<const T> skip(fms::view<const T>& v, T l, T r, T e)
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

		auto data = fms::view<const T>(v.buf, n - 1);
		v.buf += n;
		v.len -= n;

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
	inline DWORD find(fms::view<const T>& v, T s, T l, T r,T e)
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
				n += skip(vn, l, r, e).len + 2; // count l and r
			}
			else {
				++n;
			}
		};

		return n;
	}
		
	// next chunk up to separator and advance view 
	template<class T>
	inline fms::view<const T> chop(fms::view<const T>& v, T s, T l, T r, T e)
	{
		DWORD n = find(v, s, l, r, e);
		auto data = fms::view<const T>(v.buf, n);
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
		fms::view<const T> v;
		T s, l, r, e;
		DWORD n;
	public:
		using iterator_category = std::forward_iterator_tag;
		using value_type = fms::view<const T>;
		iterator()
			: s(0), l(0), r(0), e(0), n(0)
		{ }
		iterator(const fms::view<const T>& _v, T s, T l, T r, T e)
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
			return iterator(fms::view<const T>(v.buf + v.len, 0), s, l, r, e);
		}
		value_type operator*() const
		{
			return fms::view<const T>(v.buf, n);
		}
		iterator& operator++()
		{
			if (v) {
				v.buf += n;
				v.len -= n;
				if (v) {
					v = eat(s, v);
				}
				n = find(v, s, l, r, e); // n = chop().len
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
			xll::parse::iterator i(v, _T('\n'), 0, 0, _T('\\'));
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
			xll::parse::iterator is(v, _T('\n'), 0, 0, _T('\\'));
			DWORD n = 1;
			for (const auto& i : is) {
				ensure(i.equal(fms::view(a.buf, n)));
				++n;
			}
		}
		{
			fms::view v(_T("\"ab\"\ncd"));
			xll::parse::iterator i(v, _T('\n'), 0, 0, _T('\\'));
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

} // xll::parse
