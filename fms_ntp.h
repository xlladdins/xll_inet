// fms_ntp.h - Network Time Protocol
#pragma once
#include <cstdint>
#include <ctime>

namespace fms {

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
			IPv6 = 4, // either IPv4 or IPv6
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
			li : 2,
			vn : 3,
			mode : 3,
			stratum : 8,
			poll : 8, // poll interval in log2 seconds
			precision : 8; // precision of the local clock in log2 seconds
		int32_t delay;      // root delay
		int32_t dispersion; // root dispersion
		char identifier[4]; // reference identifier
		Timestamp reference;
		Timestamp originate;
		Timestamp receive;
		Timestamp transmit;
		NTP(int32_t li = LI::NO_WARNING, int32_t vn = VN::IPv6)
			: li(li), vn(vn), mode(0), stratum(0), poll(0), precision(0),
			delay(0), dispersion(0), identifier{ 0,0,0,0 }
		{ }
	};

} // namespace fms
