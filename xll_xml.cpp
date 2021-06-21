// xll_xml.cpp - pugixml wrapper
#include <sstream>
#include "libxml2.h"
#include "xll/xll/xll.h"

#define CATEGORY "XML"

using namespace xml;
using namespace xll;

extern "C" void xmlErrorHandler(void* xml_error, xmlErrorPtr error)
{
	*(xmlErrorPtr)xml_error = *error;
	std::ostringstream err(__FUNCTION__ ": ");
	err << error->message << "\n";
	err << "file: " << error->file << "\n";
	err << "line: " << error->line << "\n";
	if (error->str1) {
		err << "info: " << error->str1 << "\n";
		if (error->str2) {
			err << "info: " << error->str2 << "\n";
			if (error->str3) {
				err << "info: " << error->str3 << "\n";
			}
		}
	}

	switch (error->level) {
	case XML_ERR_WARNING:
		XLL_WARNING(err.str().c_str());
		break;
	case XML_ERR_ERROR:
	case XML_ERR_FATAL:
		XLL_ERROR(err.str().c_str());
		break;
	default:
		XLL_ERROR(__FUNCTION__ ": unknown error");
	}
}

inline xmlNodePtr node_pointer(HANDLEX node)
{
	return to_pointer<xmlNode>(node);
}

inline HANDLEX node_handle(xmlNodePtr node)
{
	return to_handle<xmlNode>(node);
}

#define XML_PARSE_TOPIC "http://xmlsoft.org/html/libxml-parser.html"
#define XML_PARSE_ENUM(X) \
	X(XML_PARSE_RECOVER, "recover on errors") \
	X(XML_PARSE_NOENT, "substitute entities") \
	X(XML_PARSE_DTDLOAD, "load the external subset") \
	X(XML_PARSE_DTDATTR, "default DTD attributes") \
	X(XML_PARSE_DTDVALID, "validate with the DTD") \
	X(XML_PARSE_NOERROR, "suppress error reports") \
	X(XML_PARSE_NOWARNING, "suppress warning reports") \
	X(XML_PARSE_PEDANTIC, "pedantic error reporting") \
	X(XML_PARSE_NOBLANKS, "remove blank nodes") \
	X(XML_PARSE_SAX1, "use the SAX1 interface internally") \
	X(XML_PARSE_XINCLUDE, "Implement XInclude substitution ") \
	X(XML_PARSE_NONET, "Forbid network access") \
	X(XML_PARSE_NODICT, "Do not reuse the context dictionary") \
	X(XML_PARSE_NSCLEAN, "remove redundant namespaces declarations") \
	X(XML_PARSE_NOCDATA, "merge CDATA as text nodes") \
	X(XML_PARSE_NOXINCNODE, "do not generate XINCLUDE STARTnodes") \
	X(XML_PARSE_COMPACT, "compact small text no") \
	X(XML_PARSE_OLD10, "parse using XML-1.0 before update 5") \
	X(XML_PARSE_NOBASEFIX, "do not fixup XINCLUDE xml:base uris") \
	X(XML_PARSE_HUGE, "relax any hardcoded limit from the parser") \
	X(XML_PARSE_OLDSAX, "parse using SAX2 interface before 2.7.0") \
	X(XML_PARSE_IGNORE_ENC, "ignore internal document encoding hint") \
	X(XML_PARSE_BIG_LINES, "Store big lines numbers in text PSVI field") \

#define XML_PARSE_DATA(a, b) XLL_CONST(LONG, a, (LONG)a, b, CATEGORY, XML_PARSE_TOPIC);

XML_PARSE_ENUM(XML_PARSE_DATA)

#undef XML_PARSE_TOPIC
#undef XML_PARSE_DATA

AddIn xai_xml_document(
	Function(XLL_HANDLEX, "xll_xml_document", "\\XML.DOCUMENT")
	.Arguments({
		Arg(XLL_HANDLEX, "data", "is a handle to a string"),
		Arg(XLL_CSTRING4, "_url", "is an optional URL."),
		Arg(XLL_CSTRING4, "_encoding", "is an optional encoding."),
		Arg(XLL_SHORT, "_options", "are optional options."),
		})
	.Uncalced()
	.Category(CATEGORY)
	.FunctionHelp("Return handle to a XML document.")
	.Documentation(R"(
Load and parse <code>data</code> for a XML document.
)")
);
HANDLEX WINAPI xll_xml_document(HANDLEX str, const char* url, const char* encoding, SHORT options)
{
#pragma XLLEXPORT
	HANDLEX h = INVALID_HANDLEX;

	try {
		handle<view<char>> str_(str);
		//options |= XML_PARSE_HUGE;
		if (!*url) url = nullptr;
		if (!*encoding) encoding = nullptr;
		handle<xml::document> h_(new xml::document(str_->buf, str_->len, url, encoding, options));

		h = h_.get();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return h;
}

AddIn xai_xml_document_root(
	Function(XLL_LPOPER, "xll_xml_document_root", "XML.DOCUMENT.ROOT")
	.Arguments({
		Arg(XLL_HANDLEX, "doc", "is a handle returned by \\XLL.DOCUMENT."),
		})
		.Category(CATEGORY)
	.FunctionHelp("Return pointers to the root node of a XML document.")
	.Documentation(R"(
)")
);
LPOPER WINAPI xll_xml_document_root(HANDLEX doc)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		o = ErrNA;
		handle<xml::document> doc_(doc);
		ensure(doc_);
		
		xmlNodePtr root = xmlDocGetRootElement(*doc_);
		if (root) {
			o = OPER(node_handle(root));
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}

AddIn xai_xml_node_children(
	Function(XLL_LPOPER, "xll_xml_node_children", "XML.NODE.CHILDREN")
	.Arguments({
		Arg(XLL_HANDLEX, "node", "is a handle to a XML node."),
		})
		.Category(CATEGORY)
	.FunctionHelp("Return children of a node.")
	.Documentation(R"(
Array of pointers to all children nodes.
)")
);
LPOPER WINAPI xll_xml_node_children(HANDLEX node)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		xmlNode* pnode = node_pointer(node);

		o = OPER{};
		for (const auto child : node::children(pnode)) {
			o.push_back(OPER(node_handle(child)));
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}

AddIn xai_xml_node_content(
	Function(XLL_LPOPER, "xll_xml_node_content", "XML.NODE.CONTENT")
	.Arguments({
		Arg(XLL_HANDLEX, "node", "is a pointer to a XML node."),
		})
		.Category(CATEGORY)
	.FunctionHelp("Return content of a node.")
	.Documentation(R"(
Content of the node or <code>#N/A</code> if none.
)")
);
LPOPER WINAPI xll_xml_node_content(HANDLEX node)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		o = ErrNA;
		if (node) {
			o = node::content(node_pointer(node));
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}

AddIn xai_xml_node_name(
	Function(XLL_LPOPER, "xll_xml_node_name", "XML.NODE.NAME")
	.Arguments({
		Arg(XLL_HANDLEX, "node", "is a pointer to a XML node."),
		})
		.Category(CATEGORY)
	.FunctionHelp("Return name of a node.")
	.Documentation(R"(
Name of the node.
)")
);
LPOPER WINAPI xll_xml_node_name(HANDLEX node)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		o = ErrNA;
		o = (char*)node_pointer(node)->name;
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}

AddIn xai_xml_node_next(
	Function(XLL_LPOPER, "xll_xml_node_next", "XML.NODE.NODES")
	.Arguments({
		Arg(XLL_HANDLEX, "node", "is a handle to a XML node."),
		})
		.Category(CATEGORY)
	.FunctionHelp("Return next nodes of a XML node.")
	.Documentation(R"(
All nodes following <code>node</code>.
)")
);
LPOPER WINAPI xll_xml_node_next(HANDLEX node)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		xmlNode* pnode = node_pointer(node);
		if (pnode) {
			for (const auto nodes : node::iterator(pnode)) {
				o.push_back(OPER(node_handle(nodes)));
			}
		}
		else {
			o = ErrNA;
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}

AddIn xai_xml_node_path(
	Function(XLL_LPOPER, "xll_xml_node_path", "XML.NODE.PATH")
	.Arguments({
		Arg(XLL_HANDLEX, "node", "is a handle to a XML node."),
		})
		.Category(CATEGORY)
	.FunctionHelp("Return path of a node.")
	.Documentation(R"(
Full path to node in XML document.
)")
);
LPOPER WINAPI xll_xml_node_path(HANDLEX node)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		o = ErrNA;
		o = (char*)xmlGetNodePath(node_pointer(node));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}


AddIn xai_xml_node_type(
	Function(XLL_LPOPER, "xll_xml_node_type", "XML.NODE.TYPE")
	.Arguments({
		Arg(XLL_HANDLEX, "element", "is a ponter to an XML element."),
		})
		.Category(CATEGORY)
	.FunctionHelp("Get element type.")
);
LPOPER WINAPI xll_xml_node_type(HANDLEX node)
{
#pragma XLLEXPORT
	static OPER ElementType[] = {
		OPER("UNKNOWN"),
		OPER("ELEMENT"),
		OPER("ATTRIBUTE"),
		OPER("TEXT"),
		OPER("CDATA_SECTION"),
		OPER("ENTITY_REF"),
		OPER("ENTITY"),
		OPER("PI"),
		OPER("COMMENT"),
		OPER("DOCUMENT"),
		OPER("DOCUMENT_TYPE"),
		OPER("DOCUMENT_FRAG"),
		OPER("NOTATION"),
		OPER("HTML_DOCUMENT"),
		OPER("DTD"),
		OPER("ELEMENT_DECL"),
		OPER("ATTRIBUTE_DECL"),
		OPER("ENTITY_DECL"),
		OPER("NAMESPACE_DECL"),
		OPER("XINCLUDE_START"),
		OPER("XINCLUDE_END"),
	};
	xmlNode* pnode = node_pointer(node);
	xmlElementType type = pnode ? pnode->type : static_cast<xmlElementType>(0);

	return &ElementType[type];
}

AddIn xai_xpath_query(
	Function(XLL_HANDLEX, "xll_xpath_query", "\\XPATH.QUERY")
	.Arguments({
		Arg(XLL_HANDLEX, "doc", "is a handle to a XML document."),
		Arg(XLL_CSTRING4, "query", "is a XPath query."),
		})
	.Uncalced()
	.FunctionHelp("Return handle to a XPath query")
	.Category(CATEGORY)
	.Documentation(R"()")
);
HANDLEX WINAPI xll_xpath_query(HANDLEX doc, const char* query)
{
#pragma XLLEXPORT
	HANDLEX h = INVALID_HANDLEX;

	try {
		handle<xml::document> doc_(doc);
		ensure(doc_);
		handle<xpath::query> query_(new xpath::query(*doc_, query));

		h = query_.get();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return h;
}

AddIn xai_xpath_query_nodes(
	Function(XLL_LPOPER, "xll_xpath_query_nodes", "XPATH.QUERY.NODES")
	.Arguments({
		Arg(XLL_HANDLEX, "query", "is a handle returned by \\XPATH.QUERY."),
		})
	.FunctionHelp("Return all nodes matched by query.")
	.Category(CATEGORY)
	.Documentation(R"()")
);
LPOPER WINAPI xll_xpath_query_nodes(HANDLEX query)
{
#pragma XLLEXPORT
	static OPER nodes;

	try {
		handle<xpath::query> query_(query);
		ensure(query_);

		nodes = OPER{};
		for (const auto node : *query_) {
			nodes.push_back(OPER(node_handle(node)));
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &nodes;
}

/*
XML.NODE.ATTRIBUTE.TYPE
	XML_ATTRIBUTE_CDATA = 1,
	XML_ATTRIBUTE_ID,
	XML_ATTRIBUTE_IDREF	,
	XML_ATTRIBUTE_IDREFS,
	XML_ATTRIBUTE_ENTITY,
	XML_ATTRIBUTE_ENTITIES,
	XML_ATTRIBUTE_NMTOKEN,
	XML_ATTRIBUTE_NMTOKENS,
	XML_ATTRIBUTE_ENUMERATION,
	XML_ATTRIBUTE_NOTATION
*/

#if 0

#ifdef _DEBUG

int test_xml_document()
{
	try {
		OPER o = Excel(xlUDF, OPER("\\INET.READ"), OPER("https://google.com"));
		OPER doc = Excel(xlUDF, OPER("\\XML.DOCUMENT"), o);
		ensure(doc.is_num());
		xml_node node = to_node(doc.val.num);
		ensure(node.type() == xml_node_type::node_document);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}

	return TRUE;
}
Auto<OpenAfter> xaoa_test_xml_document(test_xml_document);

#endif // _DEBUG

AddIn xai_xml_node_type(
	Function(XLL_LPOPER, "xll_xml_node_type", "XML.NODE.TYPE")
	.Arguments({
		Arg(XLL_HANDLEX, "node", "is a handle to a XML node."),
		})
		.Category(CATEGORY)
	.FunctionHelp("Return type of a node.")
	.Documentation(R"()")
);
#define XML_NODE_SWITCH(a, b) case xml_node_type::node_##a: { o = OPER(#a); break; }
LPOPER WINAPI xll_xml_node_type(HANDLEX h)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		xml_node node = to_node(h);
		switch (node.type()) {
			XML_NODE_TYPE(XML_NODE_SWITCH)
		default:
			o = ErrNA;
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}
#undef XML_NODE_SWITCH

AddIn xai_xml_node_name(
	Function(XLL_LPOPER, "xll_xml_node_name", "XML.NODE.NAME")
	.Arguments({
		Arg(XLL_HANDLEX, "node", "is a handle to a XML node."),
		})
		.Category(CATEGORY)
	.FunctionHelp("Return name of a node.")
	.Documentation(R"()")
);
LPOPER WINAPI xll_xml_node_name(HANDLEX h)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		o = to_node(h).name();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}

AddIn xai_xml_node_value(
	Function(XLL_LPOPER, "xll_xml_node_value", "XML.NODE.VALUE")
	.Arguments({
		Arg(XLL_HANDLEX, "node", "is a handle to a XML node."),
		})
		.Category(CATEGORY)
	.FunctionHelp("Return value of a node.")
	.Documentation(R"()")
);
LPOPER WINAPI xll_xml_node_value(HANDLEX h)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		o = to_node(h).value();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}

AddIn xai_xml_node_text(
	Function(XLL_LPOPER, "xll_xml_node_text", "XML.NODE.TEXT")
	.Arguments({
		Arg(XLL_HANDLEX, "node", "is a handle to a XML node."),
		})
		.Category(CATEGORY)
	.FunctionHelp("Return text of a node.")
	.Documentation(R"(
Get text from <code>cdata</code> or <code>pcdata</code> node.
)")
);
LPOPER WINAPI xll_xml_node_text(HANDLEX h)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		xml_node node = to_node(h);
		ensure(node.type() == xml_node_type::node_cdata or node.type() == xml_node_type::node_pcdata);
		o = node.text().get();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}

AddIn xai_xml_node_path(
	Function(XLL_LPOPER, "xll_xml_node_path", "XML.NODE.PATH")
	.Arguments({
		Arg(XLL_HANDLEX, "node", "is a handle to a XML node."),
		})
		.Category(CATEGORY)
	.FunctionHelp("Return path of a node.")
	.Documentation(R"(
Full path to node in XML document.
)")
);
LPOPER WINAPI xll_xml_node_path(HANDLEX h)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		o = to_node(h).path().c_str();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}

AddIn xai_xml_node_attributes(
	Function(XLL_LPOPER, "xll_xml_node_attributes", "XML.NODE.ATTRIBUTES")
	.Arguments({
		Arg(XLL_HANDLEX, "node", "is a handle to a XML node."),
		})
	.Category(CATEGORY)
	.FunctionHelp("Return two column range of all attributes of a node.")
	.Documentation(R"(
Return range of attribute keys in first column with corresponding
attribute values in the second column.
)")
);
LPOPER WINAPI xll_xml_node_attributes(HANDLEX h)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		o = OPER{};
		for (const auto& attr : to_node(h).attributes()) {
			o.push_back(OPER({ OPER(attr.name()), OPER(attr.value()) }));
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}

AddIn xai_xml_node_children(
	Function(XLL_LPOPER, "xll_xml_node_children", "XML.NODE.CHILDREN")
	.Arguments({
		Arg(XLL_HANDLEX, "node", "is a handle to a XML node."),
		})
		.Category(CATEGORY)
	.FunctionHelp("Return pointers to all children of a node.")
	.Documentation(R"()")
);
LPOPER WINAPI xll_xml_node_children(HANDLEX h)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		o = OPER{};
		for (const auto& node : to_node(h)) {
			o.push_back(OPER(from_node(node)));
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}

AddIn xai_xpath_node_set(
	Function(XLL_HANDLEX, "xll_xpath_node_set", "\\XPATH.NODE_SET")
	.Arguments({
		Arg(XLL_HANDLEX, "doc", "is a handle to a XML document."),
		Arg(XLL_CSTRING4, "query", "is an XPath query."),
		})
	.Uncalced()
	.Category("XML")
	.FunctionHelp("Return handle to an XPath node set matching query.")
	.Documentation(R"(
Return a handle to all nodes in <code>doc</code> matching <code>query</code>.
)")
);
HANDLEX WINAPI xll_xpath_node_set(HANDLEX doc, const char* xpath)
{
#pragma XLLEXPORT
	HANDLEX h = INVALID_HANDLEX;

	try {
		handle<xml_document> doc_(doc);
		if (doc_) {
			handle<xpath_node_set> set_(new xpath_node_set(doc_->select_nodes(xpath)));
			h = set_.get();
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return h;
}

AddIn xai_xpath_node_set_size(
	Function(XLL_LONG, "xll_xpath_node_set_size", "XPATH.NODE_SET.SIZE")
	.Arguments({
		Arg(XLL_HANDLEX, "nodes", "is a handle returned by \\XPATH.NODE_SET."),
		})
	.Category("XML")
	.FunctionHelp("Return the number of nodes in a node set.")
	.Documentation(R"()")
);
LONG WINAPI xll_xpath_node_set_size(HANDLEX nodes)
{
#pragma XLLEXPORT
	LONG n = 0;

	try {
		handle<xpath_node_set> nodes_(nodes);

		n = (LONG)nodes_->size();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return n;
}

AddIn xai_xpath_node_set_nodes(
	Function(XLL_LPOPER, "xll_xpath_node_set_nodes", "XPATH.NODE_SET.XPATH_NODES")
	.Arguments({
		Arg(XLL_HANDLEX, "nodes", "is a handle returned by \\XPATH.NODE_SET."),
		})
	.Category(CATEGORY)
	.FunctionHelp("Return pointers to all xpath_nodes in an XPath node set.")
	.Documentation(R"(
)")
);
LPOPER WINAPI xll_xpath_node_set_nodes(HANDLEX nodes)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		handle<xpath_node_set> nodes_(nodes);
		ensure(nodes_);

		o = OPER{};
		for (const xpath_node& node : *nodes_) {
			o.push_back(OPER(from_node(node.node())));
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}
AddIn xai_xpath_node_node(
	Function(XLL_HANDLEX, "xll_xpath_node_node", "XPATH_NODE.NODE")
	.Arguments({
		Arg(XLL_HANDLEX, "node", "is a handle returned by \\XPATH.NODE_SET.XPATH_NODES"),
		})
		.Category(CATEGORY)
	.FunctionHelp("Return node of XPath node.")
	.Documentation(R"(
)")
);
HANDLEX WINAPI xll_xpath_node_node(HANDLEX node)
{
#pragma XLLEXPORT
	HANDLEX h = INVALID_HANDLEX;

	try {
		handle<xpath_node> node_(node);
		if (node_) {
			h = from_node(node_->node());
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return h;
}

#ifdef _DEBUG

int test_xml()
{
	try {
		OPER o = Excel(xlUDF, OPER("\\INET.READ"), OPER("https://xlladdins.com"));
		OPER doc = Excel(xlUDF, OPER("\\XML.DOCUMENT"), o);
		OPER nodes = Excel(xlUDF, OPER("\\XPATH.NODE_SET"), doc, OPER("//*"));
		handle<xpath_node_set> nodes_(nodes.val.num);
		size_t n;
		n = nodes_->size();
		OPER N = Excel(xlUDF, OPER("XPATH.NODE_SET.SIZE"), nodes);
		ensure(N.val.num == n);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}

	return TRUE;
}
Auto<OpenAfter> xaoa_test_xml(test_xml);

#endif // _DEBUG
#endif // 0