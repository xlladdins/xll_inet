// libxml2.h - XML wrappers
#pragma once
#include <libxml/parser.h>
#include <libxml/tree.h>

extern "C" void xmlErrorHandler(void* userData, xmlErrorPtr error);

namespace xml {

	
	class document {
		xmlDocPtr pdoc;
	public:
		document()
			: pdoc(nullptr)
		{ }
		document(const char* buf, int len)
		{
			LIBXML_TEST_VERSION;
			xmlSetStructuredErrorFunc(nullptr, xmlErrorHandler);
			pdoc = xmlReadMemory(buf, len, NULL, NULL, 0);
		}
		document(const document&) = delete;
		document& operator=(const document&) = delete;
		~document()
		{ 
			xmlFreeDoc(pdoc);
			xmlCleanupParser();
		}

		operator xmlDoc*()
		{
			return pdoc;
		}
		operator const xmlDoc*() const
		{
			return pdoc;
		}
	};

} // namespace xml
