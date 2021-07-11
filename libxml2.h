// libxml2.h - libxml2 wrappers
// http://xmlsoft.org/
#pragma once
#include <compare>
#include <iterator>
#include <libxml/parser.h>
#include <libxml/tree.h>
#include <libxml/HTMLparser.h>
#include <libxml/HTMLtree.h>
#include <libxml/xpath.h>

#define XML_NODE_TYPE_ENUM(X) \
	X(XML_ELEMENT_NODE, ELEMENT) \
	X(XML_ATTRIBUTE_NODE, ATTRIBUTE) \
	X(XML_TEXT_NODE, TEXT) \
	X(XML_CDATA_SECTION_NODE, CDATA) \
	X(XML_ENTITY_REF_NODE, ENTITY_REF) \
	X(XML_ENTITY_NODE, ENTITY) \
	X(XML_PI_NODE, PI) \
	X(XML_COMMENT_NODE, COMMENT) \
	X(XML_DOCUMENT_NODE, DOCUMENT) \
	X(XML_DOCUMENT_TYPE_NODE, DOCUMENT_TYPE) \
	X(XML_DOCUMENT_FRAG_NODE, DOCUMENT_FRAG) \
	X(XML_NOTATION_NODE, NOTATION) \
	X(XML_HTML_DOCUMENT_NODE, HTML_DOCUMENT) \
	X(XML_DTD_NODE, DTD) \
	X(XML_ELEMENT_DECL, ELEMENT_DECL) \
	X(XML_ATTRIBUTE_DECL, ATTRIBUTE_DECL) \
	X(XML_ENTITY_DECL, ENTITY_DECL) \
	X(XML_NAMESPACE_DECL, NAMESPACE_DECL) \
	X(XML_XINCLUDE_START, XINCLUDE_START) \
	X(XML_XINCLUDE_END, XINCLUDE_END) \

#define XML_PARSER_OPTION_ENUM(X) \
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

// Send errors to Excel alert pop-ups
extern "C" void xmlErrorHandler(void* userData, xmlErrorPtr error);

namespace xml {

	class document {
	protected:
		xmlDocPtr pdoc;
	public:
		document(xmlDocPtr pdoc = nullptr)
			: pdoc(pdoc)
		{ }
		document(const char* buf, int len, const char* url = nullptr, const char* encoding = nullptr, int options = 0)
		{
			LIBXML_TEST_VERSION;
			xmlLineNumbersDefault(1);
			xmlSetStructuredErrorFunc(nullptr, xmlErrorHandler);
			pdoc = xmlReadMemory(buf, len, url, encoding, options);
		}
		document(const document&) = delete;
		document& operator=(const document&) = delete;
		virtual ~document()
		{
			xmlFreeDoc(pdoc);
			xmlCleanupParser();
		}

		xmlDocPtr ptr()
		{
			return pdoc;
		}
		const xmlDocPtr ptr() const
		{
			return pdoc;
		}
	};

	class node {
		xmlNodePtr pnode;
	public:
		using iterator_category = std::forward_iterator_tag;
		using value_type = xmlNodePtr;

		node(xmlNodePtr pnode = nullptr)
			: pnode(pnode)
		{ }

		explicit operator bool() const
		{
			return pnode != nullptr;
		}
		auto operator<=>(const node&) const = default;
		
		xmlNodePtr operator*()
		{
			return pnode;
		}
		xmlNodePtr ptr() const
		{
			return pnode;
		}

		auto begin() const
		{
			return *this;
		}
		auto end() const
		{
			return node(nullptr);
		}

		node& operator++()
		{
			if (pnode)
				pnode = pnode->next;

			return *this;
		}
		node& operator--()
		{
			if (pnode)
				pnode = pnode->prev;

			return *this;
		}

		const char* name() const
		{
			return (const char*)pnode->name;
		}
		// temporary strings
		class string {
			xmlChar* str;
		public:
			string(xmlChar* str)
				: str(str)
			{ }
			string(const string&) = delete;
			string& operator=(const string&) = delete;
			~string()
			{
				xmlFree(str);
			}
			const char* ptr() const
			{
				return (const char*)str;
			}
			operator const char* () const
			{
				return (const char*)str;
			}
		};
		const node::string content() const
		{
			return node::string(xmlNodeGetContent(pnode));
		}
		const node::string path() const
		{
			return node::string(xmlGetNodePath(pnode));
		}
		const node::string list(int inLine = 1) const
		{
			return node::string(xmlNodeListGetString(pnode->doc, pnode, inLine));
		}
	};

} // namespace xml

namespace html {

	struct document : public xml::document {
		document(const char* buf, int len, const char* url = nullptr, const char* encoding = nullptr, int options = 0)
		{
			LIBXML_TEST_VERSION;
			xmlLineNumbersDefault(1);
			xmlSetStructuredErrorFunc(nullptr, xmlErrorHandler);
			document::pdoc = htmlReadMemory(buf, len, url, encoding, options);
		}
	};

} // html

namespace xpath {

	class context {
		xmlXPathContextPtr pcontext;
	public:
		context(xmlDocPtr pdoc)
			: pcontext(xmlXPathNewContext(pdoc))
		{ }
		context(const context&) = delete;
		context& operator=(const context&) = delete;
		~context()
		{
			xmlXPathFreeContext(pcontext);
		}

		xmlXPathContextPtr ptr()
		{
			return pcontext;
		}
		const xmlXPathContextPtr ptr() const
		{
			return pcontext;
		}
	};

	class object {
		xmlXPathObjectPtr pobject;
	public:
		object(xmlChar* query, xmlXPathContextPtr context)
			: pobject(xmlXPathEvalExpression(query, context))
		{
			if (pobject and xmlXPathNodeSetIsEmpty(pobject->nodesetval)) {
				xmlXPathFreeObject(pobject);
				pobject = nullptr;
			}
		}
		object(const object&) = delete;
		object& operator=(const object&) = delete;
		~object()
		{
			if (pobject)
				xmlXPathFreeObject(pobject);
		}

		xmlXPathObjectPtr ptr()
		{
			return pobject;
		}
		const xmlXPathObjectPtr ptr() const
		{
			return pobject;
		}

		class iterator {
			xmlNodeSetPtr pnodeset;
			int index;
		public:
			using iterator_category = std::forward_iterator_tag;
			using value_type = xmlNodePtr;

			iterator(xmlNodeSetPtr pnodeset, int index)
				: pnodeset(pnodeset), index(index)
			{ }
			iterator(xmlXPathObjectPtr pobject)
				: pnodeset(pobject ? pobject->nodesetval : nullptr), index(0)
			{ }
			iterator(const iterator&) = default;
			iterator& operator=(const iterator&) = default;
			~iterator()
			{ }

			auto operator<=>(const iterator&) const = default;

			iterator begin() const
			{
				return *this;
			}
			iterator end() const
			{
				return iterator(pnodeset, pnodeset->nodeMax);
			}
			value_type operator*() const
			{
				return pnodeset->nodeTab[index];
			}
			iterator& operator++()
			{
				if (pnodeset and index < pnodeset->nodeMax) {
					++index;
				}

				return *this;
			}
			iterator& operator++(int)
			{
				auto tmp(*this);

				operator++();

				return tmp;
			}
		};
	};

	class query {
		object::iterator i;
	public:
		query(xmlDocPtr doc, const char* path)
			: i(object((xmlChar*)path, context(doc).ptr()).ptr())
		{ }
		query(const query&) = delete;
		query& operator=(const query&) = delete;
		~query()
		{ }

		// iterate over node set returned by query
		auto begin() const
		{
			return i.begin();
		}
		auto end() const
		{
			return i.end();
		}
	};

} // namespace xpath
