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

	namespace node {
		class content {
			xmlChar* pcontent;
		public:
			content(xmlNodePtr pnode)
				: pcontent(xmlNodeGetContent(pnode))
			{ }
			content(const content&) = delete;
			content& operator=(const content&) = delete;
			~content()
			{
				xmlFree(pcontent);
			}
			const char* ptr() const
			{
				return (const char*)pcontent;
			}
		};
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
