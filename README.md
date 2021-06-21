# xll_inet

[WinInet](https://docs.microsoft.com/en-us/windows/win32/wininet/portal) for Excel, or what
[`WEBSERVICE`](https://support.microsoft.com/en-us/office/webservice-function-0546a35a-ecc6-4739-aed7-c0b7ce1562c4)
wants be when it grows up.

A limitation of `WEBSERVICE` is that it returns a string. Strings in Excel are
limited to 32767 = 2<sup>15</sup> - 1 characters. Most web pages are larger
than that. Much larger. The function `\INET.READ(url)` returns a handle 
to all the characters returned from the URL. It uses 
[`InternetOpenUrl](https://docs.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-internetopenurla)
, [`InternetReadFile`](https://docs.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-internetreadfile)
, and memory mapped files to buffer data to memory.
The returned handle is a view of the characters. 
Use `INET.STR(view, offset, count)` to access this data.

## XML

This library uses [libxml2](http://xmlsoft.org/downloads.html) for XML/HTML parsing and XPath.
To build you should install `libxml2` using [vcpkg](https://vcpkg.io/en/).
This is what [`FILTERXML`](https://support.microsoft.com/en-us/office/filterxml-function-4df72efc-11ec-4951-86f5-c1374812f5b7)
wants to be when it grows up.

`\XML.DOCUMENT` parses the data returned by `\INET.READ`. Use `XML.DOCUMENT.ROOT`
to get the XML root node. Given any XML node, `XML.NODE.NODES(node)` returns pointers
`node` and all following nodes, if any. Use `XML.NODE.CHILDREN(node)` to get pointers
to all children of `node`. There are `XML.NODE.*` functions for the node type, name,
full document path, and content.

## XPath

`\XPATH.QUERY(doc, query)` executes a XPath query on a document.
Use `XPATH.QUERY.NODES` to get pointers to all nodes returned by the query, if any.
A simple way to get a full picture of the result of a URL query is to
call `doc = \INET.READ(url)`, `doc = \XML.DOCUMENT(doc)`, `query = XPATH.QUERY(doc, "\\*")`,
then call `XML.NODE.*` functions to get types, names, paths, and content.

## Remarks

http://worldtimeapi.org/pages/schema
