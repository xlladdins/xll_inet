// eod_historical_data.cpp - https://eodhistoricaldata.com/
#include "xll/xll/xll.h"

using namespace xll;

// Base URL for API
static const OPER EOD_URL("https://eodhistoricaldata.com/api/eod/");

// Public API key for MCD.US
static OPER EOD_TOKEN("OeAFFmMliFG5orCUuwAKQ8l4WWFQ67YX");

AddIn xai_eod_token(
	Function(XLL_LPOPER, "xll_eod_token", "EOD.TOKEN")
	.Arguments({
		Arg(XLL_CSTRING, "token", "is the EOD Historical Data API token"),
	})
	.FunctionHelp("Set Historical Data API token.")
	.Category("EOD")
	.Documentation(R"(
If <code>token</code> is missing return current EOD Historical Data API token.
)")
);
LPOPER WINAPI xll_eod_token(LPCTSTR token)
{
#pragma XLLEXPORT
	if (token and *token) {
		EOD_TOKEN = token;
	}

	return &EOD_TOKEN;
}


AddIn xai_eod_historical(
	Function(XLL_LPOPER, "xll_eod_historical", "EOD.HISTORICAL")
	.Arguments({
		Arg(XLL_CSTRING, "symbol", "is the EOD intsrument symbol.", "MCD.US"),
		Arg(XLL_DOUBLE, "from", "is the from date", "TODAY()"),
		Arg(XLL_DOUBLE, "_to", "is the optional to date. Default is TODAY()."),
		Arg(XLL_CSTRING, "_period", "is the 'd' for daily, 'w' for weekly, or 'm' for monthly. Default is daily."),
		Arg(XLL_BOOL, "_descending", "is an optional boolean indicating dates are sorted from latest to oldest.")
		})
	.FunctionHelp("Set Historical Data API historical.")
	.Category("EOD")
	.Documentation(R"(
Return URL for end-of-day data for <code>symbol</code> between <code>from</code> and <code>_to</code> date.
)")
);
LPOPER WINAPI xll_eod_historical(LPCTSTR symbol, double from, double to, LPCTSTR period, BOOL desc)
{
#pragma XLLEXPORT
	static OPER result;

	try {
		result = EOD_URL;
		result &= symbol;
		result &= "?api_token=";
		result &= EOD_TOKEN;
		result &= "&from=";
		result &= Excel(xlfText, OPER(from), OPER("yyyy-mm-dd"));
		if (to) {
			result &= "&to=";
			result &= Excel(xlfText, OPER(to), OPER("yyyy-mm-dd"));
		}
		if (period and *period) {
			ensure(*period == 'd' or *period == 'w' or *period == 'm');
			result &= "&period=";
			result &= OPER(period, 1);
		}
		result &= "&order=";
		result &= desc ? "d" : "a";
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		result = ErrNA;
	}

	return &result;
}
