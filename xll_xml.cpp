// xll_xml.cpp - pugixml wrapper
#include "pugixml-1.11/src/pugixml.hpp"
#include "xll/xll/xll.h"

#define CATEGORY "XML"

using namespace pugi;
using namespace xll;

#define XML_NODE_TYPE(X) \
		X(null,			"Empty (null) node handle") \
		X(document,		"A document tree's absolute root") \
		X(element,		"Element tag, i.e. '<node/>'") \
		X(pcdata,		"Plain character data, i.e. 'text'") \
		X(cdata,		"Character data, i.e. '<![CDATA[text]]>'") \
		X(comment,		"Comment tag, i.e. '<!-- text -->'") \
		X(pi,			"Processing instruction, i.e. '<?name?>'") \
		X(declaration,	"Document declaration, i.e. '<?xml version=\"1.0\"?>'") \
		X(doctype,		"Document type declaration, i.e. '<!DOCTYPE doc>'") \

// use internal object pointer to node struct
inline xml_node to_node(HANDLEX h)
{
	return xml_node(to_pointer<xml_node_struct>(h));
}
inline HANDLEX from_node(const xml_node& node)
{
	return to_handle<xml_node_struct>(node.internal_object());
}

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
		handle<xml_document> h_(new xml_document{});
		// Makes a copy of buf. Consider using load_buffer_inplace.
		xml_parse_result result = h_->load_buffer(str_->buf, str_->len);
		if (!result) {
			std::string err = "\\XML.DOCUMENT: parse error at: \"";
			err += std::string(str_->buf + result.offset, 32);
			err += "...\"";
			XLL_WARNING(err.c_str());
		}

		h = from_node(*h_);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return h;
}

#ifdef _DEBUG

int test_xml_document()
{
	try {
		OPER o = Excel(xlUDF, OPER("\\INET.READ"), OPER("https://google.com"));
		OPER doc = Excel(xlUDF, OPER("\\XML.DOCUMENT"), o);
		ensure(doc.is_num());
		xml_node node(to_pointer<xml_node_struct>(doc.val.num));
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
	.Documentation(R"()")
);
LPOPER WINAPI xll_xml_node_text(HANDLEX h)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		o = to_node(h).text();
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
	.Documentation(R"()")
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
		Arg(XLL_HANDLEX, "doc", "is a handle to an xpath_node_set."),
		Arg(XLL_CSTRING4, "query", "is an XPath query."),
		})
	.Uncalced()
	.Category("XML")
	.FunctionHelp("Return handle to an XPath node set matching query.")
	.Documentation(R"()")
);
HANDLEX WINAPI xll_xpath_node_set(HANDLEX doc, const char* xpath)
{
#pragma XLLEXPORT
	HANDLEX h = INVALID_HANDLEX;

	try {
		handle<xml_document> doc_(doc);
		ensure(doc_);
		xml_attribute a;
		handle<xpath_node_set> set_(new xpath_node_set(doc_->select_nodes(xpath)));

		h = set_.get();
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
	Function(XLL_LPOPER, "xll_xpath_node_set_nodes", "XPATH.NODE_SET.NODES")
	.Arguments({
		Arg(XLL_HANDLEX, "nodes", "is a handle returned by \\XPATH.NODE_SET."),
		})
		.Category("XML")
	.FunctionHelp("Return pointers to all xpath_nodes returned by \\XPATH.NODE_SET.")
	.Documentation(R"()")
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
			o.push_back(OPER(to_handle<const xpath_node>(&node)));
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}

/*
doc = \XML.DOCUMENT(data)
set = \XPATH.NODE_SET(doc, query)
nodes = XPATH.NODE_SET.NODES(set) {xpath_node,...}
xml = XML.XPATH_NODE.NODE(node)
xml = XML.XPATH_NODE.NODE(nodes, i)
*/

AddIn xai_xpath_node_node(
	Function(XLL_HANDLEX, "xll_xpath_node_node", "XML.XPATH_NODE.NODE")
	.Arguments({
		Arg(XLL_HANDLEX, "nodes", "is node or a handle returned by \\XPATH.NODE_SET.NODES."),
		Arg(XLL_WORD, "_index", "is a 1-based optional index."),
		})
	.Category("XML")
	.FunctionHelp("Return a pointer to an xml_node.")
	.Documentation(R"()")
);
HANDLEX WINAPI xll_xpath_node_set_node(HANDLEX nodes, WORD i)
{
#pragma XLLEXPORT
	HANDLEX node = INVALID_HANDLEX;

	try {
		const xpath_node* pnode;
		if (i == 0) {
			pnode = to_pointer<const xpath_node>(nodes);
			node = nodes;
		}
		else {
			handle<xpath_node_set> nodes_(nodes);
			ensure(nodes_);
			ensure(i <= nodes_->size());
			node = 0; // to_handle<const xml_node>(&(*nodes_)[i - 1].node());
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return node;
}

#if 0
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