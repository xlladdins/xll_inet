// libxml2.h - XML wrappers
#pragma once
#include <compare>
#include <iterator>
#include <libxml/parser.h>
#include <libxml/tree.h>
#include <libxml/xpath.h>


extern "C" void xmlErrorHandler(void* userData, xmlErrorPtr error);

namespace xml {

	class document {
		xmlDocPtr pdoc;
	public:
		document()
			: pdoc(nullptr)
		{ }
		document(const char* buf, int len, const char* url = nullptr, const char* encoding = nullptr, int options = 0)
		{
			xmlError error;
			memset(&error, 0, sizeof(error));

			LIBXML_TEST_VERSION;
			xmlLineNumbersDefault(1);
			xmlSetStructuredErrorFunc(&error, xmlErrorHandler);
			pdoc = xmlReadMemory(buf, len, url, encoding, options);

			if (error.code != 0) {
				xmlFreeDoc(pdoc);
				pdoc = nullptr;
			}
		}
		document(const document&) = delete;
		document& operator=(const document&) = delete;
		~document()
		{
			xmlFreeDoc(pdoc);
			xmlCleanupParser();
		}

		operator xmlDocPtr() const
		{
			return pdoc;
		}
	};

	class node {
		xmlNodePtr pnode;
	public:
		node(xmlNodePtr pnode = nullptr)
			: pnode(pnode)
		{ }
		operator xmlNodePtr() const
		{
			return pnode;
		}

		class iterator {
			xmlNodePtr pnode;
		public:
			using iterator_category = std::forward_iterator_tag;
			using value_type = node;

			iterator(xmlNodePtr pnode = nullptr)
				: pnode(pnode)
			{ }

			auto operator<=>(const iterator&) const = default;

			iterator begin() const
			{
				return iterator(pnode);
			}
			iterator end() const
			{
				return iterator{};
			}
			value_type operator*() const
			{
				return node(pnode);
			}
			iterator& operator++()
			{
				if (pnode) {
					pnode = pnode->next;
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

		class children {
			xmlNodePtr pnode;
		public:
			using iterator_category = std::forward_iterator_tag;
			using value_type = node;

			children()
				: pnode(nullptr)
			{ }
			children(xmlNodePtr pnode)
				: pnode(xmlFirstElementChild(pnode))
			{ }
			auto operator<=>(const children&) const = default;
			children begin() const
			{
				return children(pnode);
			}
			children end() const
			{
				return children{};
			}
			value_type operator*() const
			{
				return node(pnode);
			}
			children& operator++()
			{
				if (pnode) {
					pnode = xmlNextElementSibling(pnode);
				}

				return *this;
			}
			children& operator++(int)
			{
				auto tmp(*this);

				operator++();

				return tmp;
			}
		};

		class content {
			xmlChar* pcontent;
		public:
			content()
				: pcontent(nullptr)
			{ }
			content(const node& pnode)
				: pcontent(xmlNodeGetContent(pnode))
			{ }
			content(const content&) = delete;
			content& operator=(const content&) = delete;
			~content()
			{
				xmlFree(pcontent);
			}
			operator char* () const
			{
				return (char*)pcontent;
			}
		};
	};

} // namespace xml

namespace xpath {

	class context {
		xmlXPathContextPtr pcontext;
	public:
		context()
			: pcontext(nullptr)
		{ }
		context(const xml::document& doc)
			: pcontext(xmlXPathNewContext(doc))
		{ }
		context(const context&) = delete;
		context& operator=(const context&) = delete;
		~context()
		{
			xmlXPathFreeContext(pcontext);
		}
		operator xmlXPathContextPtr() const
		{
			return pcontext;
		}
		/*
		xmlDocPtr doc() const
		{
			return pcontext->doc;
		}
		*/

		class object {
			xmlXPathObjectPtr pobject;
		public:
			object()
				: pobject(nullptr)
			{ }
			object(xmlChar* query, const context& context)
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
				xmlXPathFreeObject(pobject);
			}
			operator xmlXPathObjectPtr() const
			{
				return pobject;
			}

			class iterator {
				xmlNodeSetPtr pnodeset;
				xmlNodePtr pnode;
				int index;
			public:
				using iterator_category = std::forward_iterator_tag;
				using value_type = xml::node;

				iterator(xmlNodeSetPtr pnodeset = nullptr, xmlNodePtr pnode = nullptr, int index = -1)
					: pnodeset(pnodeset), pnode(pnode), index(index)
				{ }
				iterator(const context::object& object)
					: pnodeset(((xmlXPathObjectPtr)object)->nodesetval)
				{
					if (pnodeset and !xmlXPathNodeSetIsEmpty(pnodeset)) {
						pnode = pnodeset->nodeTab[0];
						index = 0;
					}
					else {
						pnode = nullptr;
						index = -1;
					}
				}
				iterator(const iterator&) = default;
				iterator& operator=(const iterator&) = default;
				~iterator()
				{ }

				auto operator<=>(const iterator&) const = default;

				iterator begin() const
				{
					return pnodeset ? iterator(pnodeset, pnodeset->nodeTab[0], 0) : iterator{};
				}
				iterator end() const
				{
					return iterator{};
				}
				value_type operator*() const
				{
					return xml::node(pnode);
				}
				iterator& operator++()
				{
					if (pnodeset) {
						++index;
						if (index == pnodeset->nodeNr) {
							pnodeset = nullptr;
							pnode = nullptr;
							index = -1;
						}
						else {
							pnode = pnodeset->nodeTab[index];
						}
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
	};

	class query {
		context ctxt;
		context::object obj;
	public:
		query(const xml::document& doc, const char* query)
			: ctxt(doc), obj((xmlChar*)query, ctxt)
		{ }
		query(const query&) = delete;
		query& operator=(const query&) = delete;
		~query()
		{ }

		// iterate over node set returned by query
		auto begin() const
		{
			return context::object::iterator(obj).begin();
		}
		auto end() const
		{
			return context::object::iterator(obj).end();
		}
	};

} // namespace xml
