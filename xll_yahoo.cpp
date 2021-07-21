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

	return OPER(static_cast<double>(mktime(&tm)));
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
	auto fixed = [](const OPER& o) { return Excel(xlfFixed, o, OPER(0), OPER(true));  };

	o = YAHOO_URL;
	o.append((*psymbol)[0]); // one symbol for now
	o.append("?period1=").append(fixed(xll_mktime(start)));
	o.append("&period2=").append(fixed(xll_mktime(end ? OPER(end) : Excel(xlfNow))));
	o.append("&interval=").append(*interval ? interval : _T("1d"));

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