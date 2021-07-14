// xll_inet.h - https://docs.microsoft.com/en-us/windows/win32/wininet/about-wininet
#pragma once
#include "xll/xll/xll.h"
#include "xll/xll/win.h"
#include "xll/xll/win_view.h"
#include <wininet.h>

#pragma comment(lib, "Wininet.lib")

#ifndef CATEGORY
#define CATEGORY "Inet"
#endif

#define USER_AGENT "Mozilla/5.0 (Windows NT) Gecko/20100101 Firefox/89.0"

namespace Inet {

	using HInet = Win::Hnd<HINTERNET, InternetCloseHandle>;

	inline HInet hInet = InternetOpen(_T("Xll_" CATEGORY), INTERNET_OPEN_TYPE_DIRECT, NULL, NULL, 0);

} // namespace Inet
