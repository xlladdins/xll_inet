// ISO8601.h - https://www.w3.org/TR/NOTE-datetime-970915
#pragma once
#include <ctime>

namespace ISO8601 {

	template<class T>
	inline constexpr auto digit(T c) 
	{
		return '0' <= c and c <= '9' ? c - '0' : -1;
	};

	// string to number "123" -> 123, -1 on failure
	template<class T>
	inline constexpr auto poly(const T* s, int n, int b = 10)
	{
		int p = 0;

		for (int i = 0; i < n; ++i) {
			if (s[i] == 0)
				return -1;
			int d = digit(s[i]);
			if (d < 0)
				return -1;
			p = b * p + d;
		}

		return p;
	};

	template<class T>
	inline constexpr auto date_year(const T* s) {
		return poly(s, 4);
	};
	template<class T>
	inline constexpr auto date_month(const T* s) {
		int m = poly(s, 2);
		return 0 <= m and m <= 12 ? m : -1;
	};
	template<class T>
	inline constexpr auto date_day(const T* s) {
		int d = poly(s, 2);
		return 0 <= d and d <= 31 ? d : -1;
	};
	template<class T>
	inline constexpr auto date_hour(const T* s) {
		int h = poly(s, 2);
		return 0<= h and h < 24 ? h : -1;
	};
	template<class T>
	inline constexpr auto date_minute(const T* s) {
		int mm = poly(s, 2);
		return 0 <= mm and mm < 60 ? mm : -1;
	};
	template<class T>
	inline constexpr auto date_second(const T* s) {
		int ss = poly(s, 2);
		return 0 <= ss and ss < 61 ? ss : -1; // leap second ok
	};

	inline bool is_valid(struct tm* ptm)
	{
		return ptm->tm_yday != -1;
	}
	template<class T>
	inline struct tm parse(const T* s, int len = 0)
	{
		struct tm tm = { .tm_yday = -1, .tm_isdst = -1 }; // invalid date
		bool zulu = false;

		if (len == 0) {
			for (int i = 0; s[i]; ++i)
				++len;
		}

		// yyyy-mm-dd
		if (len < 4 + 1 + 2 + 1 + 2)
			goto make_time;

		// yyyy-
		tm.tm_year = date_year(s);
		if (tm.tm_year < 0)
			goto make_time;
		tm.tm_year -= 1900;
		s += 4;
		if (*s != '-')
			goto make_time;
		++s;
		len -= 5;

		// mm-
		tm.tm_mon = date_month(s);
		if (tm.tm_mon < 0)
			goto make_time;
		s += 2;
		if (*s != '-')
			goto make_time;
		tm.tm_mon -= 1;
		++s;
		len -= 3;

		// ddT?
		tm.tm_mday = date_day(s);
		if (tm.tm_mday < 0)
			goto make_time;
		s += 2;
		len -= 2;

		if (len == 0) {
			tm.tm_hour = 0;
			tm.tm_min = 0;
			tm.tm_sec = 0;
			tm.tm_isdst = -1;

			goto make_time;
		}

		if (*s != 'T')
			goto make_time;
		++s;
		--len;

		// hh:mm:ss
		//!!! allow hh, hh::mm, or hh:mm:ss ???
		if (len < 2 + 1 + 2 + 1 + 2)
			goto make_time;
		tm.tm_hour = date_hour(s);
		if (tm.tm_hour < 0)
			goto make_time;
		s += 2;
		len -= 2;

		if (*s != ':')
			goto make_time;
		++s;
		--len;
		tm.tm_min = date_minute(s);
		if (tm.tm_min < 0)
			goto make_time;
		s += 2;
		len -= 2;
		
		if (*s != ':')
			goto make_time;
		++s;
		--len;
		tm.tm_sec = date_second(s);
		if (tm.tm_sec < 0)
			goto make_time;
		s += 2;
		len -= 2;
	
		if (len == 0) {
			goto make_time;
		}
		if (*s == 'Z') {
			zulu = true;
			++s;
			--len;
		}
#if 0
		if (*s == '.') {
			++s;
			--len;
			//xcstr s0 = s;
			// secfrac
			while (digit(*s++) >= 0) {
				--len;
			}
			/*
			int n = s - s0;
			if (n)
				ss += p(s0, n) / pow(10, n);
				*/
		}
		if (len == 0)
			goto make_time;

		// [+-]?
		sgn = 1;
		if (*s == '+' or *s == '-') {
			sgn = *s == '+' ? 1 : -1;
			++s;
			--len;
		}

		// hh
		if (len < 2)
			goto make_time;
		hh_ = date_hour(s);
		if (hh_ < 0)
			goto make_time;
		tm.tm_hour += sgn * hh_;
		s += 2;
		len -= 2;

		if (len == 0) {
			goto make_time;
		}

		// :mm
		if (len < 3 or *s != ':')
			goto make_time;
		++s;
		--len;

		mm_ = date_minute(s);
		if (mm_ < 0)
			goto make_time;
		tm.tm_min += sgn * mm_;
		s += 2;
		len -= 2;

		if (len == 0)
			goto make_time;
		if (*s == 'Z') {
			zulu = true;
			++s;
			--len;
		}
#endif // 0

	make_time:
		if (len == 0) {
			time_t ignore;
			ignore = zulu ? _mkgmtime(&tm) : mktime(&tm);
		}

		return tm;
	}

#ifdef _DEBUG
#include <cassert>
	inline int test()
	{
		{
			for (int i = 0; i < 10; ++i) {
				assert(i == digit('0' + i));
			}
			assert(-1 == digit('A'));
		}
		{
			assert(1234 == poly("1234", 4));
		}
		{
			struct tm tm = parse("2021-01-02");
			assert(is_valid(&tm));
			assert(tm.tm_year == 2021 - 1900);
			assert(tm.tm_mon == 1 - 1);
			assert(tm.tm_mday == 2);
			assert(tm.tm_hour == 0);
			assert(tm.tm_min == 0);
			assert(tm.tm_sec == 0);
		}
		{
			struct tm tm = parse("2021-01-02T01:02:03");
			assert(is_valid(&tm));
			assert(tm.tm_year == 2021 - 1900);
			assert(tm.tm_mon == 1 - 1);
			assert(tm.tm_mday == 2);
			assert(tm.tm_hour == 1);
			assert(tm.tm_min == 2);
			assert(tm.tm_sec == 3);
		}

		return 0;
	}
#endif // _DEBUG
}