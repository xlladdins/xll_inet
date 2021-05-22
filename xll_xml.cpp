// xll_xml.cpp - XML functions
#include "pugixml-1.11/src/pugixml.hpp"
#include "xll/xll/xll.h"

using namespace pugi;
using namespace xll;

AddIn xai_xll_document(
	Function(XLL_HANDLEX, "xll_xml_document", "\\XML.DOCUMENT")
	.Arguments({
		Arg(XLL_HANDLEX, "handle", "is a handle to the document view.")
		})
	.Uncalced()
	.FunctionHelp("Return a handle to an XML document.")
	.Category("XML")
	.Documentation(R"()")
);
HANDLEX WINAPI xll_xml_document(HANDLEX h)
{
#pragma XLLEXPORT
	HANDLEX doc = INVALID_HANDLEX;

	try {
		handle<xml_document> doc_(new xml_document);
		handle<utf8::view<char>> h_(h); //!!! replace by generic buffer

		xml_parse_result result = doc_->load_buffer_inplace(*h_, h_->size());
		doc = doc_.get();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return doc;
}
