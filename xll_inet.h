// xll_inet.h - https://docs.microsoft.com/en-us/windows/win32/wininet/about-wininet
#pragma once
#include <cstdint>
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

	// memory mapped file class
	class MapFile : public xll::view<char> {
		Win::Handle h;
	public:
		MapFile(DWORD len = 1<<20)
			: h(CreateFileMapping(INVALID_HANDLE_VALUE, 0, PAGE_READWRITE, 0, len, nullptr))
		{
			ptr((char*)MapViewOfFile(h, FILE_MAP_ALL_ACCESS, 0, 0, len));
		}
		MapFile(const MapFile&) = delete;
		MapFile& operator=(const MapFile&) = delete;
		~MapFile()
		{
			UnmapViewOfFile(ptr());
		}
	};

	// Network Time Protocol !!! split off
	struct NTP {
		struct Timestamp {
			// 1970/1/1 - 1900/1/1 in seconds
			static constexpr uint64_t epoch = 2208988800ul;
			uint64_t secs : 32, frac : 32;
			// NTP to Unix time_t
			Timestamp(time_t t = 0)
			{
				secs = t + epoch;
				frac = 0;
			}
			time_t time() const
			{
				return secs - epoch;
			}
			// jitter() - set random frac
		};
		// leap indicator 
		enum LI {
			NO_WARNING,
			LAST_MINUTE_59,
			LAST_MINURE_61,
			ALARM,
		};
		// version number 
		enum VN {
			IPv4 = 3, // only
			IPV7 = 4, // either IPv4 or IPv6
		};
		enum MODE {
			RESERVED,
			SYMMETRIC_ACTIVE,
			SYMMETRIC_PASSIVE,
			CLIENT,
			SERVER,
			BROADCAST,
		};
		enum STRATUM {
			UNSPECIFIED,
			PRIMARY,
			// 2-15 secondary reference
			// 16-255 reserved
		};
		int32_t 
			li        : 2, 
			vn        : 3,
			mode      : 3,
			stratum   : 8,
			poll      : 8, // poll interval in log2 seconds
			precision : 8; // precision of the local clock in log2 seconds
		int32_t delay;      // root delay
		int32_t dispersion; // root dispersion
		char identifier[4]; // reference identifier
		Timestamp reference;
		Timestamp originate;
		Timestamp receive;
		Timestamp transmit;
		NTP(int32_t li = LI::NO_WARNING, int32_t vn = VN::IPv4)
			: li(li), vn(vn), mode(0), stratum(0), poll(0), precision(0),
			delay(0), dispersion(0), identifier{ 0,0,0,0 }
		{ }
	};

} // namespace Inet
