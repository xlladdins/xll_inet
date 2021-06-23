// xll_csv.h - Comma Separated Value
#pragma once
#include "ISO8601.h"
#include "xll/xll/xll.h"

#ifndef CATEGORY
#define CATEGORY "XLL"
#endif

namespace xll {

	using xchar = typename traits<XLOPERX>::xchar;
	using xcstr = typename traits<XLOPERX>::xcstr;

	template<class T>
	struct item {
		T* buf;
		DWORD len;

		item()
			: buf(nullptr), len(0)
		{ }
		item(T* buf, DWORD len)
			: buf(buf), len(len)
		{ }
		template<size_t N>
		item(T(&buf)[N])
			: buf(buf), len(static_cast<DWORD>(N))
		{ }
		item(const item&) = default;
		item& operator=(const item&) = default;
		virtual ~item()
		{ }
		
		explicit operator bool() const
		{
			return len != 0;
		}
		bool operator==(const item& i) const
		{
			return len == i.len and std::equal(buf, buf + len, i.buf);
		}

		T back() const
		{
			return buf[len-1];
		}
	};

	// remove chars that match from the end
	template<class T>
	inline item<T> trim_back(item<T> i, xcstr match = _T(" "))
	{
		while (_tcschr(match, i.back())) {
			--i.len;
		}

		return i;
	}
	// remove chars that match from the end
	template<class T>
	inline item<T> trim_front(item<T> i, xcstr match = _T(" "))
	{
		while (_tcschr(match, *i.buf)) {
			++i.buf;
			--i.len;
		}

		return i;
	}
	template<class T>
	inline item<T> trim(item<T> i, xcstr match = _T(" "))
	{
		return trim_front(trim_back(i, match));
	}

#ifdef _DEBUG

	inline int xll_test_item()
	{
		try {
			{ 
				item i(_T(""));
				ensure(i.len == 1);
				ensure(*i.buf == 0);
				item i2{ i };
				ensure(i2 == i);
				i = i2;
				ensure(!(i != i2));
			}
			{
				item i(_T("abc"));
				ensure(i.len == 4);
				ensure(trim(i) == item(_T("abc")));
			}
			{
				item i(_T(" abc "));
				ensure(i.len == 6);
				ensure(trim(i) == item(_T("abc")));
			}
		}
		catch (...) {
			return FALSE;
		}

		return TRUE;
	}

#endif // _DEBUG

	// "2021-06-22T21:30:48.323279+01:00"

#ifdef _DEBUG
#endif // _DEBUG

	// null or rs. return nullptr if not eol
	inline xcstr csv_eol(xcstr s, xcstr rs = _T("\n"))
	{
		if (*s == 0) {
			return s;
		}
		if (0 == _tcsncmp(s, rs, _tcslen(rs))) {
			return s + _tcslen(rs);
		}
		
		return nullptr;
	}

	// skip matching left and right chars ignoring escape
	inline xcstr csv_skip(xcstr s, xchar l = _T('\"'), xchar r = _T('\"'), xchar esc = _T('\"'))
	{
		if (*s != l) {
			return s;
		}

		++s;
		int level = 1;
		while (level and *s) {
			if (*s == esc) {
				++s;
				if (*s) {
					++s;
				}
			}
			if (*s == r) {
				--level;
			}
			else if (*s == l) {
				++level;
			}
			++s;
		};

		return level == 0 ? s : nullptr;
	}

#ifdef _DEBUG

	inline int xll_test_csv_skip() {

		try {
			ensure(0 == *csv_eol(_T("\n")));
			ensure(_T('a') == *csv_eol(_T("\nabc")));
			ensure(0 == csv_eol(_T("abc\n")));

			LPCTSTR s;
			s = csv_skip(_T(""), _T('{'), _T('}'), _T('\\'));
			ensure(s && *s == 0);
			s = csv_skip(_T("a"), _T('{'), _T('}'), _T('\\'));
			ensure(s && *s == _T('a'));
			s = csv_skip(_T("{a}"), _T('{'), _T('}'), _T('\\'));
			ensure(s && *s == 0);
			s = csv_skip(_T("{a}}"), _T('{'), _T('}'), _T('\\'));
			ensure(s && *s == _T('}'));
			s = csv_skip(_T("{a{}"), _T('{'), _T('}'), _T('\\'));
			ensure(!s);
			s = csv_skip(_T("{a{}}b"), _T('{'), _T('}'), _T('\\'));
			ensure(s && *s == _T('b'));
			s = csv_skip(_T("{a\\}{}}b"), _T('{'), _T('}'), _T('\\'));
			ensure(s && *s == _T('b'));
		}
		catch (...) {
			return FALSE;
		}

		return TRUE;
	}

#endif // _DEBUG

	// view of next record and advance row pointer
	inline auto csv_parse_row_item(xcstr& csv,
		xchar fs = _T(','), xcstr rs = _T("\n"), xchar esc = _T('\"'))
	{
		if (csv_eol(csv, rs)) {
			return item(csv, 0);
		}

		if (*csv == fs) {
			++csv;
		}
		xcstr e = csv;
		while (*e != fs and !csv_eol(e, rs)) {
			e = csv_skip(e, esc, esc, esc);
			++e;
		}

		item i(csv, static_cast<DWORD>(e - csv));
		csv = e;

		return trim(i, _T(" "));
	}

	inline OPER csv_parse_field(item<const xchar> field)
	{
		static OPER t("TRUE"), f("FALSE");

		field = trim(field, _T(" "));

		OPER o(field.buf, (xchar)field.len);

		if (o.val.str[0] > 0) {
			xchar i = 1;
			if (o.val.str[0] > 0) {
				if (_tcschr(_T("+-."), o.val.str[1])) {
					i = 2;
				}
			}
			if (o.val.str[0] >= i and isdigit(o.val.str[i])) {
				o = Excel(xlfValue, o); // number or date
			}
			else if (o == t) {
				o = true;
			}
			else if (o == f) {
				o = false;
			}
		}

		return o;
	}

#ifdef _DEBUG
	int xll_test_csv_parse_field()
	{
		try {
			ensure(csv_parse_field(item(_T("abc"))) == OPER("abc"));
			ensure(csv_parse_field(item(_T(" abc "))) == OPER("abc"));
			ensure(csv_parse_field(item(_T("1.23"))) == OPER(1.23));
			ensure(csv_parse_field(item(_T("2020-1-2"))) == Excel(xlfDate, OPER(2020), OPER(1), OPER(2)));
			ensure(csv_parse_field(item(_T("tRue"))) == OPER(true));
		}
		catch (const std::exception& ex) {
			XLL_ERROR(ex.what());

			return FALSE;
		}

		return TRUE;
	}
#endif // _DEBUG

	inline OPER csv_parse_row(xcstr& csv, xchar fs = _T(','), xcstr rs = _T("\n"), xchar esc = _T('"'))
	{
		OPER row;

		while (auto item = csv_parse_row_item(csv, fs, rs, esc)) {
			row.push_back(csv_parse_field(item));
		}
		csv = csv_eol(csv, rs);
		
		return row.resize(1, row.size());
	}

	inline OPER csv_parse(xcstr& csv, xchar fs = _T(','), xcstr rs = _T("\n"), xchar esc = _T('"'))
	{
		OPER table;

		while (csv and *csv) {
			OPER row = csv_parse_row(csv, fs, rs, esc);
			if (table.columns() > 0) {
				if (row.size() != table.columns()) {
					XLL_WARNING(__FUNCTION__ ": not all rows have the same size");
					row.resize(1, table.columns());
				}
			}
			table.push_back(row);
		}

		return table;
	}


} // namespace xll

