// xll_xml.cpp - pugixml wrapper
#include "libxml2.h"
#include "xll/xll/xll.h"

#define CATEGORY "XML"

using namespace xml;
using namespace xll;

AddIn xai_xml_document(
	Function(XLL_HANDLEX, "xll_xml_document", "\\XML.DOCUMENT")
	.Arguments({
		Arg(XLL_HANDLEX, "data", "is a handle to a string"),
		})
	.Uncalced()
	.Category(CATEGORY)
	.FunctionHelp("Return a handle to the root node of a XML document.")
	.Documentation(R"(
Load and parse data for a XML document.
)")
);
HANDLEX WINAPI xll_xml_document(HANDLEX str)
{
#pragma XLLEXPORT
	HANDLEX h = INVALID_HANDLEX;

	try {
		handle<view<char>> str_(str);
		handle<xml::document> h_(new xml::document(str_->buf, str_->len));

		h = h_.get();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return h;
}

AddIn xai_xml_document_nodes(
	Function(XLL_LPOPER, "xll_xml_document_nodes", "XML.DOCUMENT.NODES")
	.Arguments({
		Arg(XLL_HANDLEX, "doc", "is a handle returned by \\XLL.DOCUMENT."),
		})
		.Category(CATEGORY)
	.FunctionHelp("Return pointers to all nodes in a XML document.")
	.Documentation(R"(
)")
);
LPOPER WINAPI xll_xml_document_nodes(HANDLEX doc)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		handle<xml::document> doc_(doc);
		ensure(doc_);
		
		xmlNodePtr root = xmlDocGetRootElement(*doc_);
		o = OPER{};
		for (xmlNodePtr node = root; node; node = node->next) {
			o.push_back(OPER(to_handle<xmlNode>(node)));
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
		o = (char*)to_pointer<xmlNode>(node)->name;
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
		o = (char*)xmlGetNodePath(to_pointer<xmlNode>(node));
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
Pointers to children nodes.
)")
);
LPOPER WINAPI xll_xml_node_children(HANDLEX node)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		xmlNode* pnode = to_pointer<xmlNode>(node);

		o = OPER{};
		for (xmlNodePtr child = pnode->children; child; child = child->next) {
			o.push_back(OPER(to_handle<xmlNode>(child)));
		}

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
	};
	xmlNode* pnode = to_pointer<xmlNode>(node);
	xmlElementType type = pnode->type;

	return &ElementType[type];
}

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