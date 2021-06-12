// xll_xml.cpp - pugixml wrapper
#include "pugixml-1.11/src/pugixml.hpp"
#include "xll/xll/xll.h"

using namespace pugi;
using namespace xll;

AddIn xai_xml_document(
	Function(XLL_HANDLEX, "xll_xml_document", "\\XML.DOCUMENT")
	.Arguments({
		Arg(XLL_HANDLEX, "str", "is a handle to a string"),
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
			XLL_WARNING(result.description());
		}

		h = h_.get();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return h;
}

AddIn xai_xml_xpath(
	Function(XLL_HANDLEX, "xll_xml_xpath", "\\XML.XPATH")
	.Arguments({
		Arg(XLL_HANDLEX, "doc", "is a handle to an XML document."),
		Arg(XLL_CSTRING4, "xpath", "is an XPath query"),
		})
	.Uncalced()
	.Category("XML")
	.FunctionHelp("Return handle to node set matching query.")
	.Documentation(R"()")
);
HANDLEX WINAPI xll_xml_xpath(HANDLEX doc, const char* xpath)
{
#pragma XLLEXPORT
	HANDLEX h = INVALID_HANDLEX;

	try {
		handle<xml_document> doc_(doc);
		handle<xpath_node_set> set_(new xpath_node_set(doc_->select_nodes(xpath)));

		h = set_.get();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return h;
}
