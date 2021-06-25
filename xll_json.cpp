#include "xll_parse_json.h"

using namespace xll;

#ifdef _DEBUG

Auto<OpenAfter> aoa_test_parse_json([]() { return parse::json::test<XLOPERX>(); });

#endif // _DEBUG