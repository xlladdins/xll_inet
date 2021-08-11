// xll_parse.h - string parsing
#pragma once
#include "xll/xll/xll.h"
#include "fms_parse/fms_parse.h"

// phony xltypes
#define xltypeDate (xltypeNum|xlbitXLFree)
#define xltypeTime (xltypeNum|xlbitDLLFree)

#define XLTYPE(X) \
	X(xltypeNum, _NUM, "is a 64-bit floating point number") \
	X(xltypeStr, _STR, "is a string") \
	X(xltypeBool, _BOOL, "is a boolean") \
	X(xltypeRef, _REF, "is a multiple reference") \
	X(xltypeErr, _ERR, "is an error") \
	X(xltypeMulti, _MULTI, "is a two dimensional range") \
	X(xltypeMissing, _MISSING, "is missing function argument") \
	X(xltypeNil, _NIL, "is a null type") \
	X(xltypeSRef, _SREF, "is a single reference") \
	X(xltypeInt, _INT, "is an integer") \
	X(xltypeDate, _DATE, "is a datetime") \
	X(xltypeTime, _TIME, "is a time") \

namespace xll::parse {

	template<class X>
	inline void convert(XOPER<X>& o, int type)
	{
		if (type == xltypeNum || type == xltypeBool || type == xltypeInt) {
			o = Excel(xlfEvaluate, o);
			if (type == xltypeBool and !o.is_bool()) {
				o = !!o;
				ensure(o.is_bool());
			}
			else if (type == xltypeInt) {
				o.val.w = static_cast<traits<X>::xint>(o.val.num);
				o.xltype = xltypeInt;
			}
		}
		else if (type == xltypeDate) {
			o = Excel(xlfDatevalue, o);
			o.xltype = xltypeDate;
		}
		else if (type == xltypeTime) {
			o = Excel(xlfTimevalue, o);
			o.xltype = xltypeTime;
		}
	}
	template<class X>
	inline void convert(XOPER<X>& o, const XOPER<X>& type)
	{
		if (type.is_num() and type.as_num() != 0) {
			convert(o, static_cast<int>(type.as_num()));
		}
	}

} // xll
