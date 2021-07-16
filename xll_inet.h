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

#define INET_ICU_TOPIC "https://docs.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-internetcanonicalizeurla"
#define INET_ICU(X) \
X(ICU_BROWSER_MODE, "Does not encode or decode characters after \"#\" or \"?\", and does not remove trailing white space after \"?\". If this value is not specified, the entire URL is encoded and trailing white space is removed.") \
X(ICU_DECODE, "Converts all %XX sequences to characters, including escape sequences, before the URL is parsed.") \
X(ICU_ENCODE_PERCENT, "Encodes any percent signs encountered. By default, percent signs are not encoded. This value is available in Microsoft Internet Explorer 5 and later.") \
X(ICU_ENCODE_SPACES_ONLY, "Encodes spaces only.") \
X(ICU_NO_ENCODE, "Does not convert unsafe characters to escape sequences.") \
X(ICU_NO_META, "Does not remove meta sequences (such as \".\" and \"..\") from the URL. ") \

#define INET_INTERNET_FLAG_TOPIC "https://docs.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-internetopenurla"
#define INET_INTERNET_FLAG(X) \
X(INTERNET_FLAG_EXISTING_CONNECT, "Attempts to use an existing InternetConnect object if one exists with the same attributes required to make the request.") \
X(INTERNET_FLAG_HYPERLINK, "Forces a reload if there was no Expires timeand no LastModified time returned from the server when determining whether to reload the item from the network.") \
X(INTERNET_FLAG_IGNORE_CERT_CN_INVALID, "Disables checking of SSL/PCT-based certificates that are returned from the server against the host namegiven in the request.WinINet functions use a simple check against certificates by comparing for matching host namesand simple wildcarding rules.") \
X(INTERNET_FLAG_IGNORE_CERT_DATE_INVALID, "Disables checking of SSL/PCT - based certificates for proper validity dates.") \
X(INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTP, "Disables detection of this special type of redirect.When this flag is used, WinINet transparently allows redirects from HTTPS to HTTP URLs.") \
X(INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTPS, "Disables the detection of this special type of redirect.When this flag is used, WinINet transparently allows redirects from HTTP to HTTPS URLs.") \
X(INTERNET_FLAG_KEEP_CONNECTION, "Uses keep - alive semantics, if available, for the connection.This flag is required for Microsoft Network(MSN), NTLM, and other types of authentication.") \
X(INTERNET_FLAG_NEED_FILE, "Causes a temporary file to be created if the file cannot be cached.") \
X(INTERNET_FLAG_NO_AUTH, "Does not attempt authentication automatically.") \
X(INTERNET_FLAG_NO_AUTO_REDIRECT, "Does not automatically handle redirection in HttpSendRequest.") \
X(INTERNET_FLAG_NO_CACHE_WRITE, "Does not add the returned entity to the cache.") \
X(INTERNET_FLAG_NO_COOKIES, "Does not automatically add cookie headers to requests, and does not automatically add returned cookies to the cookie database.") \
X(INTERNET_FLAG_NO_UI, "Disables the cookie dialog box.") \
X(INTERNET_FLAG_PASSIVE, "Uses passive FTP semantics.InternetOpenUrl uses this flag for FTP filesand directories.") \
X(INTERNET_FLAG_PRAGMA_NOCACHE, "Forces the request to be resolved by the origin server, even if a cached copy exists on the proxy.") \
X(INTERNET_FLAG_RAW_DATA, "Returns the data as a WIN32_FIND_DATA structure when retrieving FTP directory information.If this flag is not specified or if the call was made through a CERN proxy, InternetOpenUrl returns the HTML version of the directory.") \
X(INTERNET_FLAG_RELOAD, "Forces a download of the requested file, object, or directory listing from the origin server, not from the cache.") \
X(INTERNET_FLAG_RESYNCHRONIZE, "Reloads HTTP resources if the resource has been modified since the last time it was downloaded.All FTP resources are reloaded.") \
X(INTERNET_FLAG_SECURE, "Uses secure transaction semantics.This translates to using Secure Sockets Layer/Private Communications Technology(SSL /PCT) and is only meaningful in HTTP requests.") \


namespace Inet {

	using HInet = Win::Hnd<HINTERNET, InternetCloseHandle>;

	inline HInet hInet = InternetOpen(_T("Xll_" CATEGORY), INTERNET_OPEN_TYPE_DIRECT, NULL, NULL, 0);

} // namespace Inet
