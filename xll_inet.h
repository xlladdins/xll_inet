// xll_inet.h - https://docs.microsoft.com/en-us/windows/win32/wininet/about-wininet
#pragma once
#include "xll/xll/xll.h"
#include "xll/xll/win.h"
#include <wininet.h>

#pragma comment(lib, "Wininet.lib")

#ifndef CATEGORY
#define CATEGORY "Inet"
#endif

#ifndef INET_URL
#define INET_URL "https://xlladdins.github.io/xll_inet/"
#endif

#define USER_AGENT "Mozilla/5.0 (Windows NT) Gecko/20100101 Firefox/89.0"

// Define name with description. Define XLL_CATEGORY and XLL_TOPIC before using
#define XLL_CONST_DEFAULT(name, desc) XLL_CONST(LONG, ##name, ##name, desc, XLL_CATEGORY, XLL_TOPIC)

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
X(INTERNET_FLAG_IGNORE_CERT_CN_INVALID, "Disables checking of SSL/PCT-based certificates that are returned from the server against the host namegiven in the request.") \
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
X(INTERNET_FLAG_RAW_DATA, "Returns the data as a WIN32_FIND_DATA structure when retrieving FTP directory information.") \
X(INTERNET_FLAG_RELOAD, "Forces a download of the requested file, object, or directory listing from the origin server, not from the cache.") \
X(INTERNET_FLAG_RESYNCHRONIZE, "Reloads HTTP resources if the resource has been modified since the last time it was downloaded.All FTP resources are reloaded.") \
X(INTERNET_FLAG_SECURE, "Uses secure transaction semantics.This translates to using Secure Sockets Layer/Private Communications Technology(SSL/PCT) and is only meaningful in HTTP requests.") \

#define INET_SCHEME_TOPIC "https://docs.microsoft.com/en-us/windows/win32/api/wininet/ne-wininet-internet_scheme"
#define INET_SCHEME(X) \
X(INTERNET_SCHEME_PARTIAL, "Partial URL.") \
X(INTERNET_SCHEME_UNKNOWN, "Unknown URL scheme.") \
X(INTERNET_SCHEME_DEFAULT, "Default URL scheme.") \
X(INTERNET_SCHEME_FTP, "FTP URL scheme (ftp:).") \
X(INTERNET_SCHEME_GOPHER, "Gopher URL scheme (gopher:).") \
X(INTERNET_SCHEME_HTTP, "HTTP URL scheme (http:).") \
X(INTERNET_SCHEME_HTTPS, "HTTPS URL scheme (https:).") \
X(INTERNET_SCHEME_FILE, "File URL scheme (file:).") \
X(INTERNET_SCHEME_NEWS, "News URL scheme (news:).") \
X(INTERNET_SCHEME_MAILTO, "Mail URL scheme (mailto:).") \
X(INTERNET_SCHEME_SOCKS, "Socks URL scheme (socks:).") \
X(INTERNET_SCHEME_JAVASCRIPT, "JScript URL scheme (javascript:).") \
X(INTERNET_SCHEME_VBSCRIPT, "VBScript URL scheme (vbscript:).") \
X(INTERNET_SCHEME_RES, "Resource URL scheme (res:).") \
X(INTERNET_SCHEME_FIRST, "Lowest known scheme value.") \
X(INTERNET_SCHEME_LAST, "Highest known scheme value.") \

#define INET_HTTP_QUERY_TOPIC "https://docs.microsoft.com/en-us/windows/win32/wininet/query-info-flags"
#define INET_HTTP_QUERY(X) \
X(HTTP_QUERY_ACCEPT, "Retrieves the acceptable media types for the response.") \
X(HTTP_QUERY_ACCEPT_CHARSET, "Retrieves the acceptable character sets for the response.") \
X(HTTP_QUERY_ACCEPT_ENCODING, "Retrieves the acceptable content-coding values for the response.") \
X(HTTP_QUERY_ACCEPT_LANGUAGE, "Retrieves the acceptable natural languages for the response.") \
X(HTTP_QUERY_ACCEPT_RANGES, "Retrieves the types of range requests that are accepted for a resource.") \
X(HTTP_QUERY_AGE, "Retrieves the Age response-header field, which contains the sender's estimate of the amount of time since the response was generated at the origin server.") \
X(HTTP_QUERY_ALLOW, "Receives the HTTP verbs supported by the server.") \
X(HTTP_QUERY_AUTHORIZATION, "Retrieves the authorization credentials used for a request.") \
X(HTTP_QUERY_CACHE_CONTROL, "Retrieves the cache control directives.") \
X(HTTP_QUERY_CONNECTION, "Retrieves any options that are specified for a particular connection and must not be communicated by proxies over further connections.") \
X(HTTP_QUERY_CONTENT_BASE, "Retrieves the base URI (Uniform Resource Identifier) for resolving relative URLs within the entity.") \
X(HTTP_QUERY_CONTENT_DESCRIPTION, "Obsolete. Maintained for legacy application compatibility only.") \
X(HTTP_QUERY_CONTENT_DISPOSITION, "Obsolete. Maintained for legacy application compatibility only.") \
X(HTTP_QUERY_CONTENT_ENCODING, "Retrieves any additional content codings that have been applied to the entire resource.") \
X(HTTP_QUERY_CONTENT_ID, "Retrieves the content identification.") \
X(HTTP_QUERY_CONTENT_LANGUAGE, "Retrieves the language that the content is in.") \
X(HTTP_QUERY_CONTENT_LENGTH, "Retrieves the size of the resource, in bytes.") \
X(HTTP_QUERY_CONTENT_LOCATION, "Retrieves the resource location for the entity enclosed in the message.") \
X(HTTP_QUERY_CONTENT_MD5, "Retrieves an MD5 digest of the entity-body for the purpose of providing an end-to-end message integrity check (MIC) for the entity-body. For more information, see RFC1864, The Content-MD5 Header Field, at http://ftp.isi.edu/in-notes/rfc1864.txt.") \
X(HTTP_QUERY_CONTENT_RANGE, "Retrieves the location in the full entity-body where the partial entity-body should be inserted and the total size of the full entity-body.") \
X(HTTP_QUERY_CONTENT_TRANSFER_ENCODING, "Receives the additional content coding that has been applied to the resource.") \
X(HTTP_QUERY_CONTENT_TYPE, "Receives the content type of the resource (such as text/html).") \
X(HTTP_QUERY_COOKIE, "Retrieves any cookies associated with the request.") \
X(HTTP_QUERY_COST, "No longer supported.") \
X(HTTP_QUERY_CUSTOM, "Causes HttpQueryInfo to search for the header name specified in lpvBuffer and store the header data in lpvBuffer.") \
X(HTTP_QUERY_DATE, "Receives the date and time at which the message was originated.") \
X(HTTP_QUERY_DERIVED_FROM, "No longer supported.") \
X(HTTP_QUERY_ECHO_HEADERS, "Not currently implemented.") \
X(HTTP_QUERY_ECHO_HEADERS_CRLF, "Not currently implemented.") \
X(HTTP_QUERY_ECHO_REPLY, "Not currently implemented.") \
X(HTTP_QUERY_ECHO_REQUEST, "Not currently implemented.") \
X(HTTP_QUERY_ETAG, "Retrieves the entity tag for the associated entity.") \
X(HTTP_QUERY_EXPECT, "Retrieves the Expect header, which indicates whether the client application should expect 100 series responses.") \
X(HTTP_QUERY_EXPIRES, "Receives the date and time after which the resource should be considered outdated.") \
X(HTTP_QUERY_FORWARDED, "Obsolete. Maintained for legacy application compatibility only.") \
X(HTTP_QUERY_FROM, "Retrieves the email address for the human user who controls the requesting user agent if the From header is given.") \
X(HTTP_QUERY_HOST, "Retrieves the Internet host and port number of the resource being requested.") \
X(HTTP_QUERY_IF_MATCH, "Retrieves the contents of the If-Match request-header field.") \
X(HTTP_QUERY_IF_MODIFIED_SINCE, "Retrieves the contents of the If-Modified-Since header.") \
X(HTTP_QUERY_IF_NONE_MATCH, "Retrieves the contents of the If-None-Match request-header field.") \
X(HTTP_QUERY_IF_RANGE, "Retrieves the contents of the If-Range request-header field. This header enables the client application to verify that the entity related to a partial copy of the entity in the client application cache has not been updated.") \
X(HTTP_QUERY_IF_UNMODIFIED_SINCE, "Retrieves the contents of the If-Unmodified-Since request-header field.") \
X(HTTP_QUERY_LAST_MODIFIED, "Receives the date and time at which the server believes the resource was last modified.") \
X(HTTP_QUERY_LINK, "Obsolete. Maintained for legacy application compatibility only.") \
X(HTTP_QUERY_LOCATION, "Retrieves the absolute Uniform Resource Identifier (URI) used in a Location response-header.") \
X(HTTP_QUERY_MAX, "Not a query flag. Indicates the maximum value of an HTTP_QUERY_* value.") \
X(HTTP_QUERY_MAX_FORWARDS, "Retrieves the number of proxies or gateways that can forward the request to the next inbound server.") \
X(HTTP_QUERY_MESSAGE_ID, "No longer supported.") \
X(HTTP_QUERY_MIME_VERSION, "Receives the version of the MIME protocol that was used to construct the message.") \
X(HTTP_QUERY_ORIG_URI, "Obsolete. Maintained for legacy application compatibility only.") \
X(HTTP_QUERY_PRAGMA, "Receives the implementation-specific directives that might apply to any recipient along the request/response chain.") \
X(HTTP_QUERY_PROXY_AUTHENTICATE, "Retrieves the authentication scheme and realm returned by the proxy.") \
X(HTTP_QUERY_PROXY_AUTHORIZATION, "Retrieves the header that is used to identify the user to a proxy that requires authentication. This header can only be retrieved before the request is sent to the server.") \
X(HTTP_QUERY_PROXY_CONNECTION, "Retrieves the Proxy-Connection header.") \
X(HTTP_QUERY_PUBLIC, "Receives methods available at this server.") \
X(HTTP_QUERY_RANGE, "Retrieves the byte range of an entity.") \
X(HTTP_QUERY_RAW_HEADERS, "Receives all the headers returned by the server. Each header is terminated by \"\\0\". An additional \"\\0\" terminates the list of headers.") \
X(HTTP_QUERY_RAW_HEADERS_CRLF, "Receives all the headers returned by the server. Each header is separated by a carriage return/line feed (CR/LF) sequence.") \
X(HTTP_QUERY_REFERER, "Receives the Uniform Resource Identifier (URI) of the resource where the requested URI was obtained.") \
X(HTTP_QUERY_REFRESH, "Obsolete. Maintained for legacy application compatibility only.") \
X(HTTP_QUERY_REQUEST_METHOD, "Receives the HTTP verb that is being used in the request, typically GET or POST.") \
X(HTTP_QUERY_RETRY_AFTER, "Retrieves the amount of time the service is expected to be unavailable.") \
X(HTTP_QUERY_SERVER, "Retrieves data about the software used by the origin server to handle the request.") \
X(HTTP_QUERY_SET_COOKIE, "Receives the value of the cookie set for the request.") \
X(HTTP_QUERY_STATUS_CODE, "Receives the status code returned by the server. For more information and a list of possible values, see HTTP Status Codes.") \
X(HTTP_QUERY_STATUS_TEXT, "Receives any additional text returned by the server on the response line.") \
X(HTTP_QUERY_TITLE, "Obsolete. Maintained for legacy application compatibility only.") \
X(HTTP_QUERY_TRANSFER_ENCODING, "Retrieves the type of transformation that has been applied to the message body so it can be safely transferred between the sender and recipient.") \
X(HTTP_QUERY_UNLESS_MODIFIED_SINCE, "Retrieves the Unless-Modified-Since header.") \
X(HTTP_QUERY_UPGRADE, "Retrieves the additional communication protocols that are supported by the server.") \
X(HTTP_QUERY_URI, "Receives some or all of the Uniform Resource Identifiers (URIs) by which the Request-URI resource can be identified.") \
X(HTTP_QUERY_USER_AGENT, "Retrieves data about the user agent that made the request.") \
X(HTTP_QUERY_VARY, "Retrieves the header that indicates that the entity was selected from a number of available representations of the response using server-driven negotiation.") \
X(HTTP_QUERY_VERSION, "Receives the last response code returned by the server.") \
X(HTTP_QUERY_VIA, "Retrieves the intermediate protocols and recipients between the user agent and the server on requests, and between the origin server and the client on responses.") \
X(HTTP_QUERY_WARNING, "Retrieves additional data about the status of a response that might not be reflected by the response status code.") \
X(HTTP_QUERY_WWW_AUTHENTICATE, "Retrieves the authentication scheme and realm returned by the server.") \
X(HTTP_QUERY_X_CONTENT_TYPE_OPTIONS, "Retrieves the X-Content-Type-Options header value.") \
X(HTTP_QUERY_P3P, "Retrieves the P3P header value.") \
X(HTTP_QUERY_X_P2P_PEERDIST, "Retrieves the X-P2P-PeerDist header value.") \
X(HTTP_QUERY_X_UA_COMPATIBLE, "Retrieves the X-UA-Compatible header value.") \
X(HTTP_QUERY_DEFAULT_STYLE, "Retrieves the Default-Style header value.") \
X(HTTP_QUERY_X_FRAME_OPTIONS, "Retrieves the X-Frame-Options header value.") \
X(HTTP_QUERY_X_XSS_PROTECTION, "Retrieves the X-XSS-Protection header value.") \
X(HTTP_QUERY_FLAG_COALESCE, "Not implemented.") \
X(HTTP_QUERY_FLAG_NUMBER, "Returns the data as a 32-bit number for headers whose value is a number, such as the status code.") \
X(HTTP_QUERY_FLAG_SYSTEMTIME, "Returns the header value as a SYSTEMTIME structure, which does not require the application to parse the data. Use for headers whose value is a date/time string, such as \"Last - Modified - Time\".") \

//!!! value > 2^53 X(HTTP_QUERY_FLAG_REQUEST_HEADERS, "Queries request headers only.") \

namespace Inet {

	using HInet = Win::Hnd<HINTERNET, InternetCloseHandle>;

	inline HInet hInet = InternetOpen(_T("Xll_" CATEGORY), INTERNET_OPEN_TYPE_DIRECT, NULL, NULL, 0);

} // namespace Inet
