#include "libxml2.h"
#include "xll/xll/xll.h"

#define CATEGORY "XPATH"

using namespace xll;
using namespace xpath;

AddIn xai_xpath_query(
	Function(XLL_LPOPER, "xll_xpath_query", "XPATH.QUERY")
	.Arguments({
		Arg(XLL_HANDLEX, "xml", "is a handle to a XML document."),
		Arg(XLL_CSTRING4, "query", "is a XPath query."),
		})
	.FunctionHelp("Return nodes in XML document matched by the XPath query")
	.Category(CATEGORY)
	.Documentation(R"(
Return all nodes of the <code>xml</code> document matching <code>query</code>.
)")
);
LPOPER WINAPI xll_xpath_query(HANDLEX doc, const char* query)
{
#pragma XLLEXPORT
	static OPER result;

	try {
		handle<xml::document> doc_(doc);
		ensure(doc_);

		xpath::query nodes(doc_->ptr(), query);
		
		result = OPER{};
		for (const auto node : nodes) {
			result.push_back(OPER(safe_handle<xmlNode>(node)));
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &result;
}

