# xll_inet

[WinInet](https://docs.microsoft.com/en-us/windows/win32/wininet/portal) for Excel, or what
[`WEBSERVICE`](https://support.microsoft.com/en-us/office/webservice-function-0546a35a-ecc6-4739-aed7-c0b7ce1562c4)
wants be when it grows up.

A limitation of `WEBSERVICE` is that it returns a string. Strings in Excel are
limited to 32767 = 2<sup>15</sup> - 1 characters. Most web pages are larger
than that. Much larger. The function `\INET.READ.VIEW(url)` returns a handle 
to all the characters returned from the URL. It uses 
[`InternetOpenUrl](https://docs.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-internetopenurla)
, [`InternetReadFile`](https://docs.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-internetreadfile)
, and memory mapped files to buffer data to memory.
The returned handle is a view of the characters and can be used to 
access (`VIEW.STR(view, offset, count)`) this data.

## Usage

```
view = \INET.READ.VIEW(url)
str = VIEW.STR(view, offset, length)
doc = XML.DOCUMENT(view) // load and parse
type = XML.NODE.TYPE(node) // document, element, pc_data, ...
name = XML.NODE.NAME(node)
attr = XML.NODE.ATTRIBUTES(node)
children = XML.NODE.CHILDREN(node)
```

## XPath

An XML _document_ is the _root_ of a directed acyclic tree of _node_s. 
Each node has _siblings_ and zero or more _children_. 
Siblings have the same _parent_ node and are the children of that parent. 
A node can be an _element_, _pcdata_, _cdata_, _comment_, _pi_, or _declaration_.

Element nodes have a _name_, _attributes_ and child nodes.
Each attribute has a _name_ and _value_.

```
node := document | element | pcdata | cdata | comment | pi | declaration
```

## Remarks

http://worldtimeapi.org/pages/schema
