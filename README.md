# xll_inet

[WinInet](https://docs.microsoft.com/en-us/windows/win32/wininet/portal) for Excel, or what
[`WEBSERVICE`](https://support.microsoft.com/en-us/office/webservice-function-0546a35a-ecc6-4739-aed7-c0b7ce1562c4)
wants to grow up to be.

The problem with `WEBSERVICE` is that it returns a string. Strings in Excel are
limited to 32767 = 2<sup>15</sup> - 1 characters and most web pages are larger
than that. Much larger. The function `\\INET.FILE` returns a handle 
to all the characters in the web page. It uses 
[`InternetReadFile`](https://docs.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-internetreadfile)
and memory mapped files to buffer data to memory.
The returned handle can be used to access or transform this data.

## Remarks

http://worldtimeapi.org/pages/schema
