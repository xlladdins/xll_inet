// xll_xml.cpp - pugixml wrapper
#include "pugixml-1.11/src/pugixml.hpp"
#include "xll/xll/xll.h"

using namespace pugi;
using namespace xll;

AddIn xai_xml_document(
	Function(XLL_HANDLEX, "xll_xml_document", "\\XML.DOCUMENT")
	.Arguments({
		Arg(XLL_HANDLEX, "data", "is a handle returned by \\INET.GET."),
		})
	.Uncalced()
	.Category("XML")
	.FunctionHelp("Return a handle to an XML document.")
	.Documentation(R"()")
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
			std::string err = "\\XML.DOCUMENT: parse error: >";
			err += std::string(str_->buf + result.offset, 20);
			err += "<...";
			XLL_WARNING(err.c_str());
		}

		h = h_.get();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return h;
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
nodes = XPATH.NODE_SET.NODES(set)
xml = XPATH_NODE.NODE(node)
xml = XPATH_NODE.NODE(nodes, i)
*/

AddIn xai_xpath_node_node(
	Function(XLL_HANDLEX, "xll_xpath_node_node", "XPATH_NODE.NODE")
	.Arguments({
		Arg(XLL_HANDLEX, "node", "is node or a handle returned by \\XPATH.NODE_SET.NODES."),
		Arg(XLL_WORD, "_index", "is a 1-based optional index."),
		})
		.Category("XML")
	.FunctionHelp("Return a pointer to an xml_node.")
	.Documentation(R"()")
);
HANDLEX WINAPI xll_xpath_node_set_nodes(HANDLEX node, WORD i)
{
#pragma XLLEXPORT
	HANDLEX node = INVALID_HANDLEX;

	try {
		if (i == 0) {
			const xml_node* pnode = to_pointer<const xml_node>(node);
			node = to_handle<const xml_node>(pnode);
		}
		else {
			handle<xpath_node_set> nodes_(node);
			ensure(nodes_);
			ensure(i <= nodes_->size());
			node = to_handle<const xml_node>(&(*nodes_)[i - 1].node());
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return node;
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