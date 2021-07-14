// xll_inet.cpp - WinInet wrappers
#include "xll_inet.h"
#include "xll_parse.h"

using namespace xll;

#define INTERNET_FLAG(X) \
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

#define TOPIC "https://docs.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-internetopenurla"
#define X(a, b) XLL_CONST(LONG, ##a, ##a, b, CATEGORY, TOPIC)
INTERNET_FLAG(X)
#undef INTERNET_FLAGS_TOPIC
#undef X

//#define XLL_STR(s) XLOPER{ .val = { .w = _countof(s)}, .xltype = xltypeStr}

// convert two columns range into "key: value\r\n ..."
inline OPER headers(const OPER& o)
{
    static OPER co(": ");
    static OPER rn("\r\n");

    OPER h;

    if (o.is_missing() or o.is_nil()) {
        h = "";
    }
    else if (o.is_str()) {
        h = o;
    }
    else if (o.is_multi()){
        ensure(o.rows() == 2);
        for (unsigned i = 0; i < o.rows(); ++i) {
            h &= o(0, i) & co & o(1, i) & rn;
        }
        h &= rn;
    }
    else {
        ensure(!"xll::headers: must be missing, nil, string, or two row range");
        h = ErrNA;
    }

    return h;
}

#ifdef _DEBUG
Auto<OpenAfter> xaoa_inet_doc([]() {
    return Documentation("INET", R"(
Functions for retrieving and parsing URLs.
)");
});
#endif // _DEBUG

AddIn xai_inet_read_file(
    Function(XLL_HANDLEX, "xll_inet_read_file", "\\URL.VIEW")
    .Arguments({
        Arg(XLL_CSTRING, "url", "is a URL to read."),
        Arg(XLL_LPOPER, "_headers", "are optional headers to send to the HTTP server."),
        Arg(XLL_LONG, "_flags", "are optional flags from INTERNET_FLAGS_*. Default is 0.")
        })
    .Uncalced()
    .Category(CATEGORY)
    .FunctionHelp("Return a handle to the string returned by url.")
    .Documentation(R"xyzyx(
Read all url data into memory using
<a href="https://docs.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-internetopenurla">InternetOpenUrl</a>
and
<a href="https://docs.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-internetreadfile">InternetReadFile</a>.
<p>
Headers are specified as a two row array of keys in the first row and values in the second.
</p>
)xyzyx")
);
HANDLEX WINAPI xll_inet_read_file(LPCTSTR url, LPOPER pheaders, LONG flags)
{
#pragma XLLEXPORT
    HANDLEX h = INVALID_HANDLEX;

    try {
        handle<fms::view<char>> h_(new win::mem_view);
        
        DWORD_PTR context = NULL;
        OPER head = OPER("User-Agent: " USER_AGENT "\r\n");
        head.append(headers(*pheaders));
        Inet::HInet hurl(InternetOpenUrl(Inet::hInet, url, head.val.str + 1, head.val.str[0], flags, context));
        ensure(hurl || !"\\INET.READ: failed to open URL")
 
        DWORD len = 4096; // ??? page size
        *h_->buf = 0; // buffer index when split
        char* buf = h_->buf;
        while (InternetReadFile(hurl, buf, len, &len) and len != 0) {
            h_->len += len;
            buf += len;
        }

        h = h_.get();
    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());
    }

    return h;
}

template<class T>
fms::view<T> drop_take_n(fms::view<T> v, LONG off, LONG len)
{
    v.drop(off);

    if (len == 0) {
        len = (LONG)v.len;
    }
    v.take(len);
    
    return v;
}

template<class T>
fms::view<T> drop_take_c(fms::view<T> v, LONG off, int c)
{
    //v.drop(off);

    if (off >= 0) {
        LONG len = 0;
        while (len < (LONG)v.len and v.buf[len] != c) {
            ++len;
        }
        v.take(len);
    }
    else if (off < 0) {
        long len = -1;
        while (-len < (LONG)v.len and v.buf[v.len + len] != c) {
            --len;
        }
        v.take(len);
    }

    return v;
}

AddIn xai_view(
    Function(XLL_LPOPER, "xll_view", "VIEW")
    .Arguments({
        Arg(XLL_HANDLEX, "handle", "is a handle returned by \\URL.VIEW."),
        Arg(XLL_LONG, "_offset", "is the view offset. Default is 0."),
        Arg(XLL_LPOPER, "_count", "is the number of characters to return. Default is all.")
        })
    .FunctionHelp("Return substring of view.")
    .Category(CATEGORY)
    .Documentation(R"xyzyx(
Drop <code>_offset</code> and take <code>_count</code> charaters from a view.
If <code>_count</code> is a string then the view is truncated at the
first character of this string if offset is positive. If offset is negative
the all characters up to the last occurence are taken.
)xyzyx")
);
LPOPER WINAPI xll_view(HANDLEX h, LONG off, LPOPER plen)
{
#pragma XLLEXPORT
    static OPER result;

    try {
        handle<fms::view<char>> v(h);
        ensure(v || !"VIEW: unrecognized handle");


        if (v->len > 3 and v->buf[0] == 0xEF) {
            ensure(v->buf[1] == 0xBB and v->buf[2] == 0xBF || !"VIEW: does not start with UTF-8 BOM");
            v->drop(3);
        }

        if (off > (LONG)v->len or off < -(LONG)v->len) {
            result = ErrNA;
        }
        else if (plen->is_num()) {
            LONG len = (LONG)plen->val.num;
            *v = drop_take_n(*v, off, len);
        }
        else if (plen->is_str() and plen->val.str[0] > 0) {
            int c = plen->val.str[1];
            *v = drop_take_c(*v, off, c);
        }
        // else noop

        result = OPER(v->buf, v->len);
    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());

        result = ErrNA;
    }

    return &result;
}

#ifdef _DEBUG

Auto<OpenAfter> xaoa_view_test([]() {
    try {
        fms::view<const char> v("abcde");

        {
            ensure(drop_take_n(v, 0, 0).equal(v));
        }
        {
            ensure(drop_take_n(v, 0, 5).equal(v));
        }
        {
            ensure(drop_take_n(v, 0, 6).equal(v));
        }
        {
            ensure(drop_take_n(v, 0, -7).equal(v));
        }
        {
            fms::view<const char> w("abc");
            ensure(drop_take_n(v, 0, 3).equal(w));
            ensure(drop_take_c(v, 0, 'd').equal(w));
        }
        {
            fms::view<const char> w("cde");
            ensure(drop_take_n(v, 2, 3).equal(w));
            ensure(drop_take_n(v, 2, 0).equal(w));
            ensure(drop_take_c(v, -1, 'b').equal(w));
        }

    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());

        return FALSE;
    }

    return TRUE;
});

#endif // _DEBUG

AddIn xai_view_len(
    Function(XLL_LPOPER, "xll_view_len", "VIEW.LEN")
    .Arguments({
        Arg(XLL_HANDLEX, "handle", "is a handle to a view of memory."),
        })
    .FunctionHelp("Return the number of characters in a view.")
    .Category(CATEGORY)
    .Documentation(R"xyzyx(
The function <code>\INET.READ</code> returns a view.
)xyzyx")
);
LPOPER WINAPI xll_view_len(HANDLEX h)
{
#pragma XLLEXPORT
    static OPER result;

    try {
        handle<fms::view<char>> h_(h);
        ensure(h_ || !"VIEW.LEN: unrecognized handle");

        result = h_->len;
    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());

        result = ErrNA;
    }

    return &result;
}

AddIn xai_view_drop(
    Function(XLL_HANDLEX, "xll_view_drop", "VIEW.DROP")
    .Arguments({
        Arg(XLL_HANDLEX, "handle", "is a handle returned by \\URL.VIEW."),
        Arg(XLL_LONG, "count", "number or characters to drop from beginning (count > 0) or end (count < 0) of view"),
        })
    .FunctionHelp("Return substring of view.")
    .Category(CATEGORY)
    .Documentation(R"xyzyx(
)xyzyx")
);
HANDLEX WINAPI xll_view_drop(HANDLEX h, LONG count)
{
#pragma XLLEXPORT
    static OPER result;

    try {
        handle<fms::view<char>> h_(h);
        ensure(h_ || !"VIEW: unrecognized handle");

        count = std::clamp(count, -(LONG)h_->len, (LONG)h_->len);
        if (count > 0) {
            h_->buf += count;
            h_->len -= count;
        }
        else if (count < 0) {
            h_->len += count;
        }
    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());

        h = 0;
    }

    return h;
}

#if 0
// up to 254 instances of split buffers
static unsigned int buf_index = 0;
static XLOPER buf_split[255] = { ErrNA4 } ;
static std::vector<XLOPER> buf_lparray[255];

AddIn xai_inet_split(
    Function(XLL_LPXLOPER4, "xll_inet_split", "INET.SPLIT")
    .Arguments({
        Arg(XLL_HANDLEX, "handle", "is a handle returned by \\INET.READ."),
        })
    .FunctionHelp("Return substring of split.")
    .Category(CATEGORY)
    .Documentation(R"xyzyx(
Split lines returned by <code>\INET.READ</code> at new line characters.
Note that this modifies the memory pointed to by handle. 
Lines longer than 255 characters containing no spaces have their last character clobbered.
)xyzyx")
);
LPXLOPER WINAPI xll_inet_split(HANDLEX h)
{
#pragma XLLEXPORT
    unsigned int b0 = 0;

    try {
        handle<fms::view<char>> h_(h);
        ensure(h_ || !"INET.SPLIT: unrecognized handle");

        char* b = h_->buf;

        b0 = b[0]; // index if already split

        if (b0 == 0) {
            b0 = ++buf_index;
            if (b0 >= 256) {
                XLL_ERROR("INET.SPLIT: at most 254 buffers allowed");

                return &buf_split[0];
            }
            buf_split[b0].xltype = xltypeMulti;
            // skip UTF-8 BOM
            if (b[1] == 0xEF) {
                ++b;
                ensure(b[1] == 0xBB || !"INET.SPLIT: data not UTF-8 encoded");
                ++b;
                ensure(b[1] == 0xBF || !"INET.SPLIT: data not UTF-8 encoded");
                ++b;
            }

            *b = 1; // for now
            while (*b) {
                char* sp = 0;
                unsigned int c = 1;
                while (c <= 255 and b[c] != 0 and b[c] != '\n') {
                    if (b[c] == ' ') {
                        sp = b + c; // last space character
                    }
                    ++c;
                }
                if (c == 256) {
                    *b = static_cast<unsigned char>(sp - b - 1);
                    XLOPER row = { .val = {.str = b }, .xltype = xltypeStr };
                    buf_lparray[b0].push_back(row);
                    if (sp == 0) {
                        XLL_WARNING("INET.SPLIT: clobbering one character in a long line");
                        b += c;
                    }
                    else {
                        b = sp;
                    }
                }
                else {
                    *b = static_cast<unsigned char>(c - 1);
                    XLOPER row = { .val = {.str = b }, .xltype = xltypeStr };
                    buf_lparray[b0].push_back(row);
                    b += c;
                }
            }
            buf_split[b0].val.array.rows = (WORD)buf_lparray[b0].size();
            buf_split[b0].val.array.columns = 1;
            buf_split[b0].val.array.lparray = buf_lparray[b0].data();
        }
    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());

        b0 = 0; // ErrNA
    }

    return &buf_split[b0];
}
#endif // 0
#if 0
AddIn xai_inet_open(
    Function(XLL_HANDLE, L"?xll_inet_open", L"INET.OPEN")
    .Arg(XLL_CSTRING, L"agent", L"is the http user agent.")
    .Uncalced()
    .Category(CATEGORY)
    .FunctionHelp(L"Return an internet connection handle.")
);
HANDLEX WINAPI xll_inet_open(xcstr agent)
{
#pragma XLLEXPORT
    handlex h;

    try {
        if (!*agent)
            agent = INET_AGENT;
        handle<Internet::Open> h_(new Internet::Open(agent));
        h = h_.get();
    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());
    }

    return h;
}


AddIn xai_inet_connect(
    Function(XLL_HANDLE, L"?xll_inet_connect", L"INET.CONNECT")
    .Arg(XLL_HANDLE, L"open", L"is a handle returned by INET.OPEN.")
    .Arg(XLL_CSTRING, L"server", L"is the name of the server to connect to.")
    .Arg(XLL_WORD, L"port", L"is the server port. Default is 80.")
    .Uncalced()
    .Category(CATEGORY)
    .FunctionHelp(L"Return an internet connection handle.")
);
HANDLEX WINAPI xll_inet_connect(HANDLEX open, xcstr server, WORD port)
{
#pragma XLLEXPORT
    handlex h;

    try {
        handle<Internet::Open> open_(open);
        ensure(open_);
        handle<Internet::Connect> conn_(new Internet::Connect(*open_, server, port));
        h = conn_.get();
    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());
    }

    return h;
}

AddIn xai_inet_http_open_request(
    Function(XLL_HANDLE, L"?xll_inet_http_open_request", L"INET.HTTP.OPEN.REQUEST")
    .Arg(XLL_HANDLE, L"connect", L"is a handle returned by INET.CONNECT.")
    .Arg(XLL_CSTRING, L"url", L"is the url for an object.")
    .Arg(XLL_CSTRING, L"verb", L"is the http opertion to perform. Default is GET.")
    .Uncalced()
    .Category(CATEGORY)
    .FunctionHelp(L"Return an internet http_open_requestion handle.")
);
HANDLEX WINAPI xll_inet_http_open_request(HANDLEX conn, xcstr url, xcstr verb)
{
#pragma XLLEXPORT
    handlex h;
    
    try {
        if (!*verb)
            verb = L"GET";

        handle<Internet::Connect> conn_(conn);
        ensure(conn_);
        handle<Internet::HttpOpenRequest> req_(new Internet::HttpOpenRequest(*conn_, verb, url));
        h = req_.get();
    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());
    }
    
    return h;
}

AddIn xai_inet_http_send_request(
    Function(XLL_HANDLE, L"?xll_inet_http_send_request", L"INET.HTTP.SEND.REQUEST")
    .Arg(XLL_HANDLE, L"request", L"is a handle returned by INET.HTTP.OPEN.REQUEST.")
    .Arg(XLL_PSTRING, L"headers", L"optional headers to send.")
    .Arg(XLL_PSTRING, L"optional", L"optional additional data for PUT or POST operation.")
    .Category(CATEGORY)
    .FunctionHelp(L"Return an internet http_send_requestion handle.")
);
HANDLEX WINAPI xll_inet_http_send_request(HANDLEX request, xcstr headers, xcstr optional)
{
#pragma XLLEXPORT
    try {
        handle<Internet::HttpOpenRequest> request_(request);
        ensure(request_);
        ensure (request_->HttpSendRequest(headers+1, headers[0], (LPVOID)(optional + 1), optional[0]));
    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());
    }

    return request;
}

AddIn xai_inet_http_query_info(
    Function(XLL_LPOPER, L"?xll_inet_http_query_info", L"INET.HTTP.QUERY.INFO")
    .Arg(XLL_HANDLE, L"request", L"is a handle returned by INET.HTTP.OPEN.REQUEST.")
    .Arg(XLL_WORD, L"info", L"is the information level to query using a HTTP_QUERY_* enumeration.")
    .Category(CATEGORY)
    .FunctionHelp(L"Return an internet http_query_infoion handle.")
);
LPOPER WINAPI xll_inet_http_query_info(HANDLEX request, WORD info)
{
#pragma XLLEXPORT
    static OPER result;

    try {
        handle<Internet::HttpOpenRequest> request_(request);
        auto query = request_->HttpQueryInfo(info);
        if (query.has_value())
            result = query.value().c_str();
        else
            result = OPER(xlerr::NA);
    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());
    }

    return &result;
}

#ifdef _DEBUG

test test_xll_inet([] {
    {
        Internet::Open open(INET_AGENT);
        Internet::Connect conn(open, L"google.com", INTERNET_DEFAULT_HTTPS_PORT);
        Internet::HttpOpenRequest req(conn, L"/index.html", L"GET");
        req.HttpSendRequest(L"", 0, 0, 0);
        auto q = req.HttpQueryInfo(HTTP_QUERY_RAW_HEADERS);
        ensure (q.has_value());
        // q.value() is padded with 0's at the end
        ensure (0 == wcscmp(q.value().c_str(), L"HTTP/1.0 200 OK"));
    }
    {/*
        Internet::Open open(INET_AGENT);
        //HINTERNET hurl;
        //hurl = ::InternetOpenUrl(open, L"https://google.com/index.html", NULL, 0, 0, 0);
        Internet::OpenUrl url(open, L"https://google.com/index.html");
        auto read = url.ReadFile();
        ensure (read.has_value());
    */}
});

#endif // _DEBUG
#endif // 0