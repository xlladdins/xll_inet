// xll_yahoo.cpp - Yahoo! finance urls
#include "xll/xll/xll.h"

#define YAHOO_URL "https://query1.finance.yahoo.com/v7/finance/download/"

using namespace xll;

using xcstr = traits<XLOPERX>::xcstr;

inline OPER xll_mktime(const OPER& date)
{
	struct tm tm;

	tm.tm_year = static_cast<int>(Excel(xlfYear, date).as_num() - 1900);
	tm.tm_mon = static_cast<int>(Excel(xlfMonth, date).as_num() - 1);
	tm.tm_mday = static_cast<int>(Excel(xlfDay, date).as_num());
	tm.tm_hour = static_cast<int>(Excel(xlfHour, date).as_num());
	tm.tm_min = static_cast<int>(Excel(xlfMinute, date).as_num());
	tm.tm_sec = static_cast<int>(Excel(xlfSecond, date).as_num());
	tm.tm_isdst = -1;

	time_t t = mktime(&tm);

	return t == (time_t)-1 ? ErrValue : OPER(static_cast<double>(t));
}

inline OPER xll_localtime(const OPER& time)
{
	OPER d = ErrValue;

	const time_t t = static_cast<time_t>(time.val.num);
	struct tm tm;
	errno_t err = localtime_s(&tm, &t);
	if (err == 0) {
		d = Excel(xlfDate, OPER(tm.tm_year + 1900), OPER(tm.tm_mon + 1), OPER(tm.tm_mday));
		d = d.as_num() + Excel(xlfTime, OPER(tm.tm_hour), OPER(tm.tm_min), OPER(tm.tm_sec)).as_num();
	}

	return d;
}

AddIn xai_mktime_(
	Function(XLL_LPOPER, "xll_mktime_", "MKTIME")
	.Arguments({
		Arg(XLL_DOUBLE, "date", "is an Excel date."),
		})
	.Category("XLL")
	.FunctionHelp("Convert the local time to a calendar value.")
	.HelpTopic("https://docs.microsoft.com/en-us/cpp/c-runtime-library/reference/mktime-mktime32-mktime64")
	.Documentation(R"()")
);
LPOPER WINAPI xll_mktime_(double date)
{
#pragma XLLEXPORT
	static OPER t;
	
	t = xll_mktime(OPER(date));

	return &t;
}


AddIn xai_localtime_(
	Function(XLL_LPOPER, "xll_localtime_", "LOCALTIME")
	.Arguments({
		Arg(XLL_DOUBLE, "time", "is a time_t."),
		})
	.Category("XLL")
	.FunctionHelp("Converts a time_t time value to an Excel date and corrects for the local time zone.")
	.HelpTopic("https://docs.microsoft.com/en-us/cpp/c-runtime-library/reference/localtime-s-localtime32-s-localtime64-s")
	.Documentation(R"()")
);
LPOPER WINAPI xll_localtime_(double time)
{
#pragma XLLEXPORT
	static OPER d;
	
	d = xll_localtime(OPER(time));

	return &d;
}

AddIn xai_yahoo_finance(
	Function(XLL_LPOPER, "xll_yahoo_finance", "YAHOO.FINANCE")
	.Arguments({
		Arg(XLL_LPOPER, "symbol", "is the list of symbols to query"),
		Arg(XLL_DOUBLE, "start", "is the start date."),
		Arg(XLL_DOUBLE, "end", "is the end date."),
		Arg(XLL_CSTRING, "_interval", "is the sampling interval. Default is \"1d\"."),
		})
	.Category("YAHOO")
	.FunctionHelp("Return url for querying Yahoo! financial data.")
	.HelpTopic("https://finance.yahoo.com/")
	.Documentation(R"xyzyx(
Return URL for Yahoo! finance.
)xyzyx")
);
LPOPER WINAPI xll_yahoo_finance(LPOPER psymbol, double start, double end, xcstr interval)
{
#pragma XLLEXPORT
	static OPER o;

	// number as text, no decimal points, no commas
	auto fixed = [](const OPER& o) { return Excel(xlfFixed, o, OPER(0), OPER(true)); };

	try {
		o = YAHOO_URL;
		o.append((*psymbol)[0]); // one symbol for now
		o.append("?period1=").append(fixed(xll_mktime(start)));
		o.append("&period2=").append(fixed(xll_mktime(end ? OPER(end) : Excel(xlfNow))));
		o.append("&interval=").append(*interval ? interval : _T("1d"));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = ErrValue;
	}

	return &o;
}

#ifdef _DEBUG

Auto<Open> xao_yahoo_finance_test([]() {

	try {
		OPER symbol("SPY");
		OPER start(Excel(xlfDate, OPER(2021), OPER(1), OPER(1)));
		OPER end(Excel(xlfDate, OPER(2021), OPER(2), OPER(1)));
		OPER url = *xll_yahoo_finance(&symbol, start.as_num(), end.as_num(), _T(""));
		ensure(url == "https://query1.finance.yahoo.com/v7/finance/download/SPY?period1=1609477200&period2=1612155600&interval=1d");
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}

	return TRUE;

	});

#endif // _DEBUG