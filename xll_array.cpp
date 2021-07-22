// xll_array.cpp - array functions
#include "xll/xll/xll.h"

using namespace xll;

AddIn xai_array_(
	Function(XLL_HANDLEX, "xll_array_", "\\ARRAY")
	.Arguments({
		Arg(XLL_FPX, "array", "is the array to set.", "{1,2;3,4}")
		})
	.Uncalced()
	.FunctionHelp("Return a handle to an array.")
	.Category("XLL")
	.Documentation(R"(
Create a handle to a two dimensional array of floating point numbers.
)")
);
HANDLEX WINAPI xll_array_(_FPX* px)
{
#pragma XLLEXPORT
	handle<FPX> h(new FPX(*px));

	return h.get();
}

AddIn xai_array_get(
	Function(XLL_FPX, "xll_array_get", "ARRAY")
	.Arguments({
		Arg(XLL_HANDLEX, "handle", "is a handle returned by \\ARRAY().")
		})
	.FunctionHelp("Return the array held by a handle.")
	.Category("XLL")
	.Documentation(R"(
Return a two dimensional array of cells.
)")
);
_FPX* WINAPI xll_array_get(HANDLEX h)
{
#pragma XLLEXPORT
	handle<FPX> h_(h);

	if (!h_) {
		XLL_ERROR(__FUNCTION__ ": unknown handle");

		return nullptr;
	}

	return h_->get();
}

AddIn xai_array_index(
	Function(XLL_FPX, "xll_array_index", "ARRAY.INDEX")
	.Arguments({
		Arg(XLL_FPX, "array", "is a array or handle to a array."),
		Arg(XLL_FPX, "rows", "are an array of rows to return."),
		Arg(XLL_FPX, "columns", "are an array of columns to return."),
		})
	.FunctionHelp("Return rows and columns of array.")
	.Category("XLL")
	.Documentation(R"(
This works like <code>INDEX</code> for arrays.
)")
);
_FPX* WINAPI xll_array_index(_FPX* parray, LPOPER prows, LPOPER pcolumns)
{
#pragma XLLEXPORT
	static FPX a;
	
	try {
		unsigned r = prows->size();
		unsigned c = pcolumns->size();

		if (size(*parray) == 1) {
			handle<FPX> h_(parray->array[0]);
			if (h_) {
				parray = h_->get();
			}
		}

		if (prows->is_missing()) {
			r = parray->rows;
		}
		if (pcolumns->is_missing()) {
			c = parray->columns;
		}

		a.resize(r, c);

		for (unsigned i = 0; i < r; ++i) {
			unsigned ri = prows->is_missing() ? i : static_cast<unsigned>((*prows)[i].val.num);
			for (unsigned j = 0; j < c; ++j) {
				unsigned cj = pcolumns->is_missing() ? j : static_cast<unsigned>((*pcolumns)[j].val.num);
				a(i, j) = index(*parray, ri, cj);
			}
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return a.get();
}

#if 0

AddIn xai_array_take(
	Function(XLL_LPOPER, "xll_array_take", "array.TAKE")
	.Arguments({
		Arg(XLL_LPOPER, "array", "is a array or handle to a array."),
		Arg(XLL_LONG, "count", "take items from front (count > 0) or back (count < 0)."),
		})
	.FunctionHelp("Return taken array items.")
	.Category("XLL")
	.Documentation(R"(
If array has one row/column then take row/columns elements,
otherwise take array rows.
)")
);
LPOPER WINAPI xll_array_take(LPOPER parray, LONG count)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		o = ErrNA;

		if (parray->is_num()) {
			handle<OPER> h_(parray->val.num);
			if (h_) {
				parray = h_.ptr();
			}
		}

		o = parray->take(count);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}

AddIn xai_array_drop(
	Function(XLL_LPOPER, "xll_array_drop", "array.DROP")
	.Arguments({
		Arg(XLL_LPOPER, "array", "is a array or handle to a array."),
		Arg(XLL_LONG, "count", "drop items from front (count > 0) or back (count < 0)."),
		})
		.FunctionHelp("Return array with dropped items removed.")
	.Category("XLL")
	.Documentation(R"(
If array has one row/column then drop row/columns elements,
otherwise drop array rows.
)")
);
LPOPER WINAPI xll_array_drop(LPOPER parray, LONG count)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		o = ErrNA;

		if (parray->is_num()) {
			handle<OPER> h_(parray->val.num);
			if (h_) {
				parray = h_.ptr();
			}
		}

		o = parray->drop(count);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}

AddIn xai_array_cumprod(
	Function(XLL_LPOPER, "xll_array_cumprod", "array.CUMPROD")
	.Arguments({
		Arg(XLL_LPOPER, "array", "is a array."),
		})
	.Category("XLL")
	.FunctionHelp("Cumulative product of array columns")
	.Documentation(R"()")
);
LPOPER WINAPI xll_array_cumprod(LPOPER parray)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		if (parray->is_num()) {
			handle<OPER> h_(parray->val.num);
			if (h_) {
				parray = h_.ptr();
			}
		}
		const OPER& array = *parray;
		o.resize(array.rows(), array.columns());
		for (unsigned j = 0; j < o.columns(); ++j) {
			o(0, j) = array(0, j);
		}
		for (unsigned i = 1; i < o.rows(); ++i) {
			for (unsigned j = 0; j < o.columns(); ++j) {
				o(i, j) = o(i - 1, j).as_num() * array(i, j).as_num();
			}
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = ErrValue;
	}

	return &o;
}
#endif