// libxml2.h - XML wrappers
#pragma once
#include <libxml/parser.h>
#include <libxml/tree.h>

namespace xml {

	class document {
		xmlDocPtr pdoc;
	public:
		document()
			: pdoc(nullptr)
		{ }
		document(const char* buf, int len)
			: pdoc(xmlReadMemory(buf, len, NULL, NULL, 0))
		{ }
		document(const document&) = delete;
		document& operator=(const document&) = delete;
		~document()
		{ 
			xmlFreeDoc(pdoc);
		}
	};

} // namespace xml
