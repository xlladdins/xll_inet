// xll_inet.h - https://docs.microsoft.com/en-us/windows/win32/wininet/about-wininet
#pragma once
#include "xll/xll/xll.h"
#include "xll/xll/win.h"
#include <wininet.h>
#include <memoryapi.h>

#pragma comment(lib, "Wininet.lib")

#ifndef CATEGORY
#define CATEGORY "Inet"
#endif

namespace Inet {

	using HInet = Win::Hnd<HINTERNET, InternetCloseHandle>;

	inline HInet hInet = InternetOpen(_T("Xll_" CATEGORY), INTERNET_OPEN_TYPE_DIRECT, NULL, NULL, 0);

} // namespace Inet
