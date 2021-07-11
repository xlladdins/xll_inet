// xll_xml.cpp -libxml2 wrapper
#include <sstream>
#include "libxml2.h"
#include "xll/xll/xll.h"

#define CATEGORY "XML"

using namespace xll;
using namespace xml;

// extract libxml2 error information
extern "C" void xmlErrorHandler(void*, xmlErrorPtr perror)
{
	if (!perror) {
		XLL_ERROR(__FUNCTION__ ": no error information");

		return;
	}
	std::ostringstream err(__FUNCTION__ ": ");
	if (perror->message) err << perror->message << "\n";
	if (perror->file) err << "file: " << perror->file << "\n";
	if (perror->line) err << "line: " << perror->line << "\n";
	if (perror->str1) {
		err << "info: " << perror->str1 << "\n";
		if (perror->str2) {
			err << "info2: " << perror->str2 << "\n";
			if (perror->str3) {
				err << "info3: " << perror->str3 << "\n";
			}
		}
	}

	switch (perror->level) {
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

#define XML_NODE_TYPE_TOPIC "http://xmlsoft.org/html/libxml-parser.html#xmlParserOption"
#define XML_NODE_TYPE_DATA(a, b) XLL_CONST(LONG, a, (LONG)a, #b, CATEGORY, XML_NODE_TYPE_TOPIC);

XML_NODE_TYPE_ENUM(XML_NODE_TYPE_DATA)

#undef XML_NODE_TYPE_DATA
#undef XML_NODE_TYPE_TOPIC

#define XML_PARSE_TOPIC "http://xmlsoft.org/html/libxml-parser.html#xmlParserOption"
#define XML_PARSE_DATA(a, b) XLL_CONST(LONG, a, (LONG)a, b, CATEGORY, XML_PARSE_TOPIC);

XML_PARSER_OPTION_ENUM(XML_PARSE_DATA)

#undef XML_PARSE_TOPIC
#undef XML_PARSE_DATA

AddIn xai_xml_document(
	Function(XLL_HANDLEX, "xll_xml_document", "\\XML.DOCUMENT")
	.Arguments({
		Arg(XLL_HANDLEX, "view", "is a handle to a memory view"),
		Arg(XLL_CSTRING4, "_url", "is an optional URL."),
		Arg(XLL_CSTRING4, "_encoding", "is an optional encoding."),
		Arg(XLL_SHORT, "_options", "are optional options from the XML_PARSE_* enumeration."),
		})
	.Uncalced()
	.Category(CATEGORY)
	.FunctionHelp("Return handle to a XML document.")
	.Documentation(R"(
Load and parse <code>view</code> for a XML document.
)")
);
HANDLEX WINAPI xll_xml_document(HANDLEX str, const char* url, const char* encoding, SHORT options)
{
#pragma XLLEXPORT
	HANDLEX h = INVALID_HANDLEX;

	try {
		handle<fms::view<char>> str_(str);
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

#define HTML_PARSER_OPTION_ENUM(X) \
	X(HTML_PARSE_RECOVER, "Relaxed parsing") \
	X(HTML_PARSE_NODEFDTD, "do not default a doctype if not found") \
	X(HTML_PARSE_NOERROR, "suppress error reports") \
	X(HTML_PARSE_NOWARNING, "suppress warning reports") \
	X(HTML_PARSE_PEDANTIC, "pedantic error reporting") \
	X(HTML_PARSE_NOBLANKS, "remove blank nodes") \
	X(HTML_PARSE_NONET, "Forbid network access") \
	X(HTML_PARSE_NOIMPLIED, "Do not add implied html/body... elements") \
	X(HTML_PARSE_COMPACT, "compact small text nodes") \
	X(HTML_PARSE_IGNORE_ENC, "ignore internal document encoding hint") \

#define HTML_PARSE_TOPIC "http://xmlsoft.org/html/libxml-HTMLparser.html#htmlParserOption"
#define HTML_PARSE_DATA(a, b) XLL_CONST(LONG, a, (LONG)a, b, CATEGORY, HTML_PARSE_TOPIC);

HTML_PARSER_OPTION_ENUM(HTML_PARSE_DATA)

#undef HTML_PARSE_TOPIC
#undef HTML_PARSE_DATA

AddIn xai_html_document(
	Function(XLL_HANDLEX, "xll_html_document", "\\HTML.DOCUMENT")
	.Arguments({
		Arg(XLL_HANDLEX, "view", "is a handle to a memory view"),
		Arg(XLL_CSTRING4, "_url", "is an optional URL."),
		Arg(XLL_CSTRING4, "_encoding", "is an optional encoding."),
		Arg(XLL_SHORT, "_options", "are optional options from the XML_PARSE_* enumeration."),
		})
		.Uncalced()
	.Category(CATEGORY)
	.FunctionHelp("Return handle to a HTML/XML document.")
	.Documentation(R"(
Load and parse <code>view</code> for a HTML/XML document. All <code>XML.*</code>
and <code>XPATH.*</code> functions can be used with the handle returned by this function.
)")
);
HANDLEX WINAPI xll_html_document(HANDLEX str, const char* url, const char* encoding, SHORT options)
{
#pragma XLLEXPORT
	HANDLEX h = INVALID_HANDLEX;

	try {
		handle<fms::view<char>> str_(str);
		//options |= XML_PARSE_HUGE;
		if (!*url) url = nullptr;
		if (!*encoding) encoding = nullptr;
		handle<xml::document> h_(new html::document(str_->buf, str_->len, url, encoding, options));

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
		Arg(XLL_HANDLEX, "doc", "is a handle returned by \\XML.DOCUMENT."),
		})
		.Category(CATEGORY)
	.FunctionHelp("Return pointers to the root node of a XML document.")
	.Documentation(R"(
Every well-formed XML document has a root node. This is it.
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
		
		auto root = xmlDocGetRootElement(doc_->ptr());
		if (root) {
			o = OPER(safe_handle<xmlNode>(root));
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}

AddIn xai_xml_node_attribues(
	Function(XLL_LPOPER, "xll_xml_node_attributes", "XML.NODE.ATTRIBUTES")
	.Arguments({
		Arg(XLL_HANDLEX, "node", "is a pointer to a XML node."),
		})
		.Category(CATEGORY)
	.FunctionHelp("Return attributes of an element node.")
	.Documentation(R"(
Content of the node or <code>#N/A</code> if none.
)")
);
LPOPER WINAPI xll_xml_node_attributes(HANDLEX pnode)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		o = ErrNA;

		xmlNodePtr node = safe_pointer<xmlNode>(pnode);
		ensure(node);
		ensure(node->type == XML_ELEMENT_NODE);

		if (node) {
			o = OPER{};
			OPER v;
			xmlAttrPtr cur = node->properties;
			while (cur and cur->name and cur->children) {
				o.push_back(OPER((const char*)cur->name));
				v.push_back(OPER(xml::node(cur->children).content()));
				cur = cur->next;
			}
			o.resize(1, o.size());
			v.resize(1, v.size());
			o.push_back(v/*, OPER::Side::Bottom*/);
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
LPOPER WINAPI xll_xml_node_content(HANDLEX pnode)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		auto node = xml::node(safe_pointer<xmlNode>(pnode));

		if (node) {
			o = node.content();
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

AddIn xai_xml_node_list(
	Function(XLL_LPOPER, "xll_xml_node_list", "XML.NODE.LIST")
	.Arguments({
		Arg(XLL_HANDLEX, "node", "is a pointer to a XML node."),
		})
		.Category(CATEGORY)
	.FunctionHelp("Return list string of a node.")
	.Documentation(R"(
List string of the node or <code>#N/A</code> if none.
)")
);
LPOPER WINAPI xll_xml_node_list(HANDLEX pnode)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		auto node = xml::node(safe_pointer<xmlNode>(pnode));

		if (node) {
			o = node.list();
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
LPOPER WINAPI xll_xml_node_name(HANDLEX pnode)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		auto node = xml::node(safe_pointer<xmlNode>(pnode));

		if (pnode) {
			o = node.name();
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

AddIn xai_xml_node_next(
	Function(XLL_LPOPER, "xll_xml_node_next", "XML.NODE.NEXT")
	.Arguments({
		Arg(XLL_HANDLEX, "node", "is a handle to a XML node."),
		Arg(XLL_LPOPER, "_type", "is an optional node type to find. Default is any."),
		})
	.Category(CATEGORY)
	.FunctionHelp("Return the next XML node of a given type.")
	.Documentation(R"(
If <code>_type</code> is missing the next node is returned.
The function <code>XML.NODE.NEXT.ELEMENT(node)</code> is
equivalent to <code>XML.NODE.NEXT(node, XML_ELEMENT_NODE())</code>
This is the usual way to walk the document tree.
)")
);
LPOPER WINAPI xll_xml_node_next(HANDLEX pnode, LPOPER ptype)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		o = ErrNA;

		xmlElementType type = static_cast<xmlElementType>(0);
		if (ptype->is_num()) {
			type = static_cast<xmlElementType>(ptype->val.num);
		}

		auto node = xml::node(safe_pointer<xmlNode>(pnode));

		if (node) {
			++node;
			while (node and type and node.type() != type) {
				++node;
			}
		}

		if (node) {
			o = safe_handle<xmlNode>(node.ptr());
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}

AddIn xai_xml_node_next_element(
	Function(XLL_LPOPER, "xll_xml_node_next_element", "XML.NODE.NEXT.ELEMENT")
	.Arguments({
		Arg(XLL_HANDLEX, "node", "is a handle to a XML node."),
		})
	.Category(CATEGORY)
	.FunctionHelp("Return the next XML node element.")
	.Documentation(R"(
This function is
equivalent to <code>XML.NODE.NEXT(node, XML_ELEMENT_NODE())</code>.
)")
);
LPOPER WINAPI xll_xml_node_next_element(HANDLEX pnode)
{
#pragma XLLEXPORT
	OPER element(XML_ELEMENT_NODE);

	return xll_xml_node_next(pnode, &element);
}

AddIn xai_xml_node_prev(
	Function(XLL_LPOPER, "xll_xml_node_prev", "XML.NODE.PREV")
	.Arguments({
		Arg(XLL_HANDLEX, "node", "is a handle to a XML node."),
	})
	.Category(CATEGORY)
	.FunctionHelp("Return preceding sibling of a XML node.")
	.Documentation(R"(
Return handles to at most <code>_count</code> preceding siblings of <code>node</code> 
starting from <code>_offset</code>.
)")
);
LPOPER WINAPI xll_xml_node_prev(HANDLEX pnode)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		auto node = xml::node(safe_pointer<xmlNode>(pnode));

		o = ErrNA;
		--node;
		if (node) {
			o = safe_handle<xmlNode>(node.ptr());
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
		Arg(XLL_HANDLEX, "node", "is a handle to the parent node."),
		})
		.Category(CATEGORY)
	.FunctionHelp("Return children of a node.")
	.Documentation(R"(
Return handles to all children of <code>node</code> that are element nodes.
)")
);
LPOPER WINAPI xll_xml_node_children(HANDLEX pnode)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		o = ErrNA;

		xmlNodePtr node = safe_pointer<xmlNode>(pnode);
		ensure(node);

		o = OPER{};
		for (auto child = node->children; child; child = child->next) {
			if (child->type == XML_ELEMENT_NODE) {
				o.push_back(OPER(safe_handle<xmlNode>(child)));
			}
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}

AddIn xai_xml_node_siblings(
	Function(XLL_LPOPER, "xll_xml_node_siblings", "XML.NODE.SIBLINGS")
	.Arguments({
		Arg(XLL_HANDLEX, "node", "is a handle to a XML node."),
		//		Arg(XLL_LONG, "_axis", "is an optional axis direction. Default is 0."),
		})
	.Category(CATEGORY)
	.FunctionHelp("Return following sibling of a XML node.")
	.Documentation(R"(
//Return handles to at most <code>_count</code> following siblings of <code>node</code> 
//starting from <code>_offset</code>.
)")
);
LPOPER WINAPI xll_xml_node_siblings(HANDLEX pnode)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		o = ErrNA;

		xmlNodePtr node = safe_pointer<xmlNode>(pnode);
		ensure(node);

		o = OPER{};
		for (auto cur = node->parent->children; cur; cur = cur->next) {
			if (cur->type == XML_ELEMENT_NODE) {
				o.push_back(OPER(safe_handle<xmlNode>(cur)));
			}
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}
/*
AddIn xai_xml_node_attributes(
	Function(XLL_LPOPER, "xll_xml_node_attributes", "XML.NODE.ATTRIBUTES")
	.Arguments({
		Arg(XLL_HANDLEX, "node", "is a handle to at node."),
		})
		.Category(CATEGORY)
	.FunctionHelp("Return attributes of a node.")
	.Documentation(R"(
Return handles to all attributes of <code>node</code>.
)")
);
LPOPER WINAPI xll_xml_node_attributes(HANDLEX node)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		xmlNode* pnode = safe_pointer<xmlNode>(node);

		if (pnode) {
			auto child = pnode->children;
			o = *xll_xml_node_next(safe_handle<xmlNode>(child), 0, 0);
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
*/

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
LPOPER WINAPI xll_xml_node_path(HANDLEX pnode)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		auto node = xml::node(safe_pointer<xmlNode>(pnode));

		if (node) {
			o = node.path();
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


AddIn xai_xml_node_type(
	Function(XLL_LPOPER, "xll_xml_node_type", "XML.NODE.TYPE")
	.Arguments({
		Arg(XLL_HANDLEX, "element", "is a ponter to an XML element."),
		})
	.Category(CATEGORY)
	.FunctionHelp("Get element type.")
	.Documentation(R"(
Return the node type: element, attribute, text, CDATA, entity reference, entity
processing instruction, comment, document, document type, document fragment,
notation, HTML document, document type definition, ...
)")
);
LPOPER WINAPI xll_xml_node_type(HANDLEX node)
{
#pragma XLLEXPORT
	static OPER o;
	
	o = ErrNA;
	auto pnode = safe_pointer<xmlNode>(node);
	if (pnode) {
		if (false)
			;
#define XLL_NODE_TYPE_VALUE(a, b) else if (pnode->type == a) { o = #b; }
		XML_NODE_TYPE_ENUM(XLL_NODE_TYPE_VALUE)
#undef XLL_NODE_TYPE_VALUE
		else {
			o = "UNKNOWN";
		}
	}

	return &o;
}


