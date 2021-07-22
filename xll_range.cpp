// xll_range.cpp - range functions
#include "xll/xll/xll.h"

using namespace xll;

AddIn xai_range_(
	Function(XLL_HANDLEX, "xll_range_", "\\RANGE")
	.Arguments({
		Arg(XLL_LPOPER, "range", "is the range to set.", "{1,2;3,4}")
		})
	.Uncalced()
	.FunctionHelp("Return a handle to a range.")
	.Category("XLL")
	.Documentation(R"(
Create a handle to a two dimensional range of cells.
)")
);
HANDLEX WINAPI xll_range_(LPOPER px)
{
#pragma XLLEXPORT
	handle<OPER> h(new OPER(*px));

	return h.get();
}

AddIn xai_range_get(
	Function(XLL_LPOPER, "xll_range_get", "RANGE")
	.Arguments({
		Arg(XLL_HANDLEX, "handle", "is a handle returned by \\RANGE().", "\\RANGE.SET({0,1;2,3})")
		})
	.FunctionHelp("Return the range held by a handle.")
	.Category("XLL")
	.Documentation(R"(
Return a two dimensional range of cells.
)")
);
LPOPER WINAPI xll_range_get(HANDLEX h)
{
#pragma XLLEXPORT
	handle<OPER> h_(h);

	if (!h_) {
		XLL_ERROR(__FUNCTION__ ": unknown handle");

		return nullptr;
	}

	return h_.ptr();
}

AddIn xai_range_index(
	Function(XLL_LPOPER, "xll_range_index", "RANGE.INDEX")
	.Arguments({
		Arg(XLL_LPOPER, "range", "is a range or handle to a range."),
		Arg(XLL_LPOPER, "rows", "are an array of rows to return."),
		Arg(XLL_LPOPER, "columns", "are an array of columns to return."),
		})
	.FunctionHelp("Return rows and columns of range.")
	.Category("XLL")
	.Documentation(R"(
This is a replacement for <code>INDEX</code> that actually works.
)")
);
LPOPER WINAPI xll_range_index(LPOPER prange, LPOPER prows, LPOPER pcolumns)
{
#pragma XLLEXPORT
	static OPER o;
	
	try {
		o = ErrValue;

		if (prange->is_num()) {
			handle<OPER> h_(prange->as_num());
			if (h_) {
				prange = h_.ptr();
			}
		}

		auto r = prows->size();
		if (prows->is_missing()) {
			r = prange->rows();
		}
		auto c = pcolumns->size();
		if (pcolumns->is_missing()) {
			c = prange->columns();
		}

		o.resize(r, c);
		OPER& range = *prange;

		for (unsigned i = 0; i < r; ++i) {
			unsigned ri = prows->is_missing() ? i : static_cast<unsigned>((*prows)[i].val.num);
			for (unsigned j = 0; j < c; ++j) {
				unsigned cj = pcolumns->is_missing() ? j : static_cast<unsigned>((*pcolumns)[j].val.num);
				o(i, j) = range(ri, cj);
			}
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}

AddIn xai_range_take(
	Function(XLL_LPOPER, "xll_range_take", "RANGE.TAKE")
	.Arguments({
		Arg(XLL_LPOPER, "range", "is a range or handle to a range."),
		Arg(XLL_LONG, "count", "take items from front (count > 0) or back (count < 0)."),
		})
	.FunctionHelp("Return taken range items.")
	.Category("XLL")
	.Documentation(R"(
If range has one row/column then take row/columns elements,
otherwise take range rows.
)")
);
LPOPER WINAPI xll_range_take(LPOPER prange, LONG count)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		o = ErrNA;

		if (prange->is_num()) {
			handle<OPER> h_(prange->val.num);
			if (h_) {
				prange = h_.ptr();
			}
		}

		if (prange->is_multi()) {
			o = prange->take(count);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}

AddIn xai_range_drop(
	Function(XLL_LPOPER, "xll_range_drop", "RANGE.DROP")
	.Arguments({
		Arg(XLL_LPOPER, "range", "is a range or handle to a range."),
		Arg(XLL_LONG, "count", "drop items from front (count > 0) or back (count < 0)."),
		})
		.FunctionHelp("Return range with dropped items removed.")
	.Category("XLL")
	.Documentation(R"(
If range has one row/column then drop row/columns elements,
otherwise drop range rows.
)")
);
LPOPER WINAPI xll_range_drop(LPOPER prange, LONG count)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		o = ErrNA;

		if (prange->is_num()) {
			handle<OPER> h_(prange->val.num);
			if (h_) {
				prange = h_.ptr();
			}
		}

		if (prange->is_multi()) {
			o = prange->drop(count);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}

AddIn xai_range_cumprod(
	Function(XLL_LPOPER, "xll_range_cumprod", "RANGE.CUMPROD")
	.Arguments({
		Arg(XLL_LPOPER, "range", "is a range."),
		})
	.Category("XLL")
	.FunctionHelp("Cumulative product of range columns")
	.Documentation(R"()")
);
LPOPER WINAPI xll_range_cumprod(LPOPER prange)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		if (prange->is_num()) {
			handle<OPER> h_(prange->val.num);
			if (h_) {
				prange = h_.ptr();
			}
		}
		const OPER& range = *prange;
		o.resize(range.rows(), range.columns());
		for (unsigned j = 0; j < o.columns(); ++j) {
			o(0, j) = range(0, j);
		}
		for (unsigned i = 1; i < o.rows(); ++i) {
			for (unsigned j = 0; j < o.columns(); ++j) {
				o(i, j) = o(i - 1, j).as_num() * range(i, j).as_num();
			}
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = ErrValue;
	}

	return &o;
}
