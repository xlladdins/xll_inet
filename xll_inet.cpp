// xll_inet.cpp - WinInet wrappers
#include "xll_inet.h"

using namespace xll;

// convert two columns range into "key: value\r\n ..."
inline OPER headers(const OPER& o)
{
    OPER h;

    if (o.is_missing()) {
        h = "";
    }
    else if (o.is_str()) {
        h = o;
    }
    else if (o.is_multi()){
        ensure(o.columns() == 2);
        OPER sc(": ");
        OPER rn("\r\n");
        for (unsigned i = 0; i < o.rows(); ++i) {
            h &= o(i, 0) & sc & o(i, 1) & rn;
        }
        h &= rn;
    }
    else {
        ensure(!"xll::headers: must be mising, string, or two column range");
        h = ErrNA;
    }

    return h;
}

AddIn xai_inet_read_file(
    Function(XLL_HANDLEX, "xll_inet_read_file", "\\INET.READ")
    .Arguments({
        Arg(XLL_CSTRING, "url", "is a URL to read."),
        Arg(XLL_LPOPER, "_headers", "are optional headers to send to the HTTP server."),
        })
    .Uncalced()
    .Category(CATEGORY)
    .FunctionHelp("Return a handle to the string returned by url.")
    .Documentation(R"xyzyx(
Read all url data into memory using
<a href="https://docs.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-internetopenurla">InternetOpenUrl</a>
and
<a href="https://docs.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-internetreadfile">InternetReadFile</a>.
)xyzyx")
);
HANDLEX WINAPI xll_inet_read_file(LPCTSTR url, LPOPER pheaders)
{
#pragma XLLEXPORT
    HANDLEX h = INVALID_HANDLEX;

    try {
        handle<utf8::view<char>> h_(new utf8::mem_view);
        
        DWORD flags = 0;
        DWORD_PTR context = NULL;
        OPER head = headers(*pheaders);
        Inet::HInet hurl(InternetOpenUrl(Inet::hInet, url, head.val.str + 1, head.val.str[0], flags, context));
 
        DWORD len = 4096; // ??? page size
        char* buf = *h_;
        while (InternetReadFile(hurl, buf, len, &len) and len != 0) {
            h_->size(h_->size() + len);
            buf += len;
        }

        h = h_.get();
    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());
    }

    return h;
}

AddIn xai_inet_file(
    Function(XLL_LPOPER, "xll_inet_file", "STR.GET")
    .Arguments({
        Arg(XLL_HANDLEX, "handle", "is a handle to a string."),
        Arg(XLL_LONG, "_offset", "is the view offset. Default is 0."),
        Arg(XLL_LONG, "_length", "is the number of characters to return. Default is all.")
        })
    .FunctionHelp("Return substring of file.")
    .Category(CATEGORY)
    .Documentation(R"xyzyx(
Get a string returned by <code>\INET.READ</code>.
)xyzyx")
);
LPOPER WINAPI xll_inet_file(HANDLEX h, LONG off, LONG len)
{
#pragma XLLEXPORT
    static OPER result;

    try {
        handle<utf8::view<char>> h_(h);

        if (len == 0 or len > static_cast<LONG>(h_->size()) - off) {
            len = h_->size() - off;
        }

        result = OPER(*h_ + off, len);
    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());

        result = ErrNA;
    }

    return &result;
}

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