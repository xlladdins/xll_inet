# xll_inet

[WinInet](https://docs.microsoft.com/en-us/windows/win32/wininet/portal) for Excel, or what
[`WEBSERVICE`](https://support.microsoft.com/en-us/office/webservice-function-0546a35a-ecc6-4739-aed7-c0b7ce1562c4)
wants be when it grows up.

A limitation of `WEBSERVICE` is that it returns a string. Strings in Excel are
limited to 32767 = 2<sup>15</sup> - 1 characters. Most web pages are larger
than that. Much larger. The function `\URL.VIEW(url)` returns a handle 
to all the characters returned from the URL. It uses 
[`InternetOpenUrl`](https://docs.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-internetopenurla)
, [`InternetReadFile`](https://docs.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-internetreadfile)
, and a memory mapped file to buffer data to memory.


## HTML/XML

This library uses [libxml2](http://xmlsoft.org/downloads.html) for HTML/XML parsing and XPath.
Install `libxml2` using [vcpkg](https://vcpkg.io/en/).
This is what [`FILTERXML`](https://support.microsoft.com/en-us/office/filterxml-function-4df72efc-11ec-4951-86f5-c1374812f5b7)
wants to be when it grows up.

The function `\URL.VIEW(url)` reads `url` and returns a view of the characters returned.
Use `VIEW(view, offset, length)` to return the characters in `view`. The default value
of `offset` is 0 and if `length` is not specified then all charaters from the offset
are returned. The first dozen or so characters let you identify what type of
document was returned.

Call `\HTML.DOCUMENT(view)` to get a handle to an HTML document containing
the parsed `view`. 
The function `\XML.DOCUMENT(view)` is also provided for parsing XML documents.
A XML _document_ is an ordered tree of _nodes_.
Every XML document has a _root_ node.
All nodes except the root node have a unique _parent_ node.
Nodes having the same parent are _siblings_ and are the _children_ of the parent.
The node ordering is called _document order_.

Nodes have a _type_, _name_, _path_, and _content_.
The most common node types are _element_, _attribute_, and _text_.
Element nodes have the form `<name attribute*>content</name>` where
attribute has the form `key="value"` and `*` indicates zero or
more occurences. Elements with no content can be
abbreviated as `<name attribute* />`
The path is the node name and list of
parents names to the root.

Use `XML.DOCUMENT.ROOT(doc)` to get the root node for either type of document.
In addition to `XML.NODE.TYPE`, `XML.NODE.NAME`, `XML.NODE.PATH`, and `XML.NODE.CONTENT`,
the function `XML.NODE.ATTRIBUTES(node)` will return a two row range of
all node attributes with keys in the first row and values in the second.
The functions `XML.NODE.NEXT(node, type)` and `XML.NODE.PREV(node, type)`
find the first sibling node following or preceding `node`
of the given `type`. If `type` is missing the next node is returned.
The function `XML.NODE.CHILD(node, type)` finds the first child of node
with `type` and `XML.NODE.PARENT(node)` returns the unique parent of `node`. 
Node types are defined in the `XML_*_NODE()` enumeration.
These functions can be used to (tediously) traverse a document node-by-node.
The functions `XML.NODE.SIBLINGS(node)` and `XML.NODE.CHILDREN(node)`
return handles to all sibling and children nodes that have type `XML_NODE_ELEMENT()`.
Element nodes are usually what you want when traversing a document.

## XPath

`XPATH.QUERY(doc, query)` returns all nodes of `doc` matching `query`.
A simple way to get a full picture of the result of a URL query is to
call `\URL.VIEW(url)`, `\XML.DOCUMENT(view)`, `XPATH.QUERY(document, "//*")`,
then call `XML.NODE.*` functions to get types, names, paths, and content.

## CSV

Comma separated strings are parsed into 2-dimensional ranges by 
`CSV.PARSE(view, rs, fs, esc, offset, count)` where `rs` is the record
separator, `fs` is the field separator, `esc` is the escape character,
`offset` is the number of initial lines to skip, and `count` is the
number of rows to return. The default record separator is comma (','),
field separator is newline ('\n'), and escape character is ('\').
The default offset is 0 and if count is missing then all lines are
returned. All range elements are returned as strings.

Use `CSV.CONVERT(range, types, index)` to convert columns specified
by (0-based) `index` into corresponding `types` from the `TYPE_*` enumeration.

## JSON

JSON strings are parsed using `JSON.PARSE` into values. Objects are
parsed into two row arrays where the first row contains the keys and
the second row contains their corresponding values. Arrays are parsed
into one row ranges. Use `JSON.TYPE(value)` to detect if the value
is an object, array, string, number, boolean, or null.

The function `JSON.INDEX(value, index)` retrieves values from a JSON `value`.
If `value` is an object and `index` is a string this is equivalent to `HLOOKUP(index, object, 2, FALSE)`.
If `value` is an array and `index` is a number this is equivalent to `INDEX(array, index + 1)`.
The index can be an array of keys and will lookup values recursively.
The index can also be specifed in dotted [jq](https://stedolan.github.io/jq/) style.

## Remarks

http://worldtimeapi.org/pages/schema
http://worldtimeapi.org/api/timezone/Europe/London.txt
https://www.wikidata.org/wiki/Wikidata:Data_access
Each item or property has a persistent URI that you obtain by appending its ID (such as Q42 or P31) to the Wikidata concept namespace: http://www.wikidata.org/entity/
Linked data clients would receive Wikidata's data about the entity in a different format such as JSON or RDF, depending on the HTTP Accept: header of their request. 
For cases in which it is inconvenient to use content negotiation (e.g. to view non-HTML content in a web browser), you can also access data about an entity in a specific format by extending the data URL with an extension suffix to indicate the content format that you want, such as .json, .rdf, .ttl, .nt or .jsonld
https://www.w3.org/TR/rdf11-concepts/

https://www.wikidata.org/w/api.php?action=wbgetentities&sites=enwiki&titles=Berlin&props=descriptions&languages=en&format=json

`ALL(f, x, args...)` returns an array `{f(x, args), f(f(x, args), args), ...}`.