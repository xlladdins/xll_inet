// xll_array.cpp - array functions
#include <numeric>
#include "xll/xll/xll.h"

#ifndef CATEGORY
#define CATEGORY "XLL"
#endif

using namespace xll;

using xcstr = traits<XLOPERX>::xcstr;

AddIn xai_array_(
	Function(XLL_HANDLEX, "xll_array_", "\\ARRAY")
	.Arguments({
		Arg(XLL_FPX, "array", "is the array to set.", "{1,2;3,4}")
		})
	.Uncalced()
	.FunctionHelp("Return a handle to an array.")
	.Category(CATEGORY)
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
	.Category(CATEGORY)
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

AddIn xai_array_size(
	Function(XLL_LONG, "xll_array_size", "ARRAY.SIZE")
	.Arguments({
		Arg(XLL_FPX, "array", "is an array or handle to an array."),
	})
	.FunctionHelp("Return the size of array.")
	.Category(CATEGORY)
	.Documentation("Return the size of an array.")
);
LONG WINAPI xll_array_size(_FPX* pa)
{
#pragma XLLEXPORT

	LONG n = size(*pa);

	if (n == 1) {
		handle<FPX> a_(pa->array[0]);
		if (a_) {
			n = a_->size();
		}
	}

	return n;
}

AddIn xai_array_rows(
	Function(XLL_LONG, "xll_array_rows", "ARRAY.ROWS")
	.Arguments({
		Arg(XLL_FPX, "array", "is an array or handle to an array."),
		})
		.FunctionHelp("Return the number of rows of an array.")
	.Category(CATEGORY)
	.Documentation("Return the number of rows of an array.")
);
LONG WINAPI xll_array_rows(_FPX* pa)
{
#pragma XLLEXPORT

	LONG n = pa->rows;

	if (n == 1) {
		handle<FPX> a_(pa->array[0]);
		if (a_) {
			n = a_->rows();
		}
	}

	return n;
}
AddIn xai_array_columns(
	Function(XLL_LONG, "xll_array_columns", "ARRAY.COLUMNS")
	.Arguments({
		Arg(XLL_FPX, "array", "is an array or handle to an array."),
		})
		.FunctionHelp("Return the number of columns of an array.")
	.Category(CATEGORY)
	.Documentation("Return the number of columns of an array.")
);
LONG WINAPI xll_array_columns(_FPX* pa)
{
#pragma XLLEXPORT

	LONG n = pa->columns;

	if (n == 1) {
		handle<FPX> a_(pa->array[0]);
		if (a_) {
			n = a_->columns();
		}
	}

	return n;
}

AddIn xai_array_resize(
	Function(XLL_FPX, "xll_array_resize", "ARRAY.RESIZE")
	.Arguments({
		Arg(XLL_FPX, "array", "is an array or handle to an array."),
		Arg(XLL_LONG, "rows", "is the number of rows."),
		Arg(XLL_LONG, "columns", "is the number of columns."),
		})
		.FunctionHelp("Resize an array.")
	.Category(CATEGORY)
	.Documentation(R"(
Resize array to <code>rows</code> and <code>columns</code>.
If <code>array</code> is a handle this function resizes the in-memory array and
returns its handle, otherwise the resized array is returned.
)")
);
_FPX* WINAPI xll_array_resize(_FPX* pa, LONG r, LONG c)
{
#pragma XLLEXPORT
	static FPX a;

	try {
		a = *pa;
		if (size(*pa) == 1) {
			handle<FPX> a_(pa->array[0]);
			if (a_) {
				a_->resize(r, c);
			}
		}
		else {
			a.resize(r, c);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	catch (...) {
		XLL_ERROR(__FUNCTION__ ": unknown exception");
	}

	return a.get();
}


AddIn xai_array_index(
	Function(XLL_FPX, "xll_array_index", "ARRAY.INDEX")
	.Arguments({
		Arg(XLL_FPX, "array", "is a array or handle to a array."),
		Arg(XLL_LPOPER, "rows", "are an array of rows to return."),
		Arg(XLL_LPOPER, "columns", "are an array of columns to return."),
		})
	.FunctionHelp("Return rows and columns of array.")
	.Category(CATEGORY)
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

AddIn xai_array_take(
	Function(XLL_FPX, "xll_array_take", "ARRAY.TAKE")
	.Arguments({
		Arg(XLL_FPX, "array", "is a array or handle to a array."),
		Arg(XLL_LONG, "count", "take items from front (count > 0) or back (count < 0)."),
		})
	.FunctionHelp("Return taken array items.")
	.Category(CATEGORY)
	.Documentation(R"(
If array has one row/column then take row/columns elements,
otherwise take array rows.
)")
);
_FPX* WINAPI xll_array_take(_FPX* parray, LONG n)
{
#pragma XLLEXPORT
	static FPX a;

	try {
		a = FPX{};
		a[0] = std::numeric_limits<double>::quiet_NaN();

		if (size(*parray) == 1) {
			handle<FPX> h_(parray->array[0]);
			if (h_) {
				xll::take(h_->get(), n);
				a[0] = parray->array[0];
			}
		}
		else {
			a = *take(parray, n);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return a.get();
}

AddIn xai_array_sequence(
	Function(XLL_FP, "xll_array_sequence", "ARRAY.SEQUENCE")
	.Arguments({
		Arg(XLL_DOUBLE, "start", "is the first value in the sequence.", "0"),
		Arg(XLL_DOUBLE, "stop", "is the last value in the sequence.", "3"),
		Arg(XLL_DOUBLE, "_incr", "is an optional value to increment by. Default is 1.")
		})
	.FunctionHelp("Return a one column array from start to stop with specified optional increment.")
	.Category(CATEGORY)
	.Documentation(R"(
Return a one columns array <code>{start; start + incr; ...; stop}<code>.
If <code>incr</code> is an integer greater than 1, then generate
an array of <code>incr</code> numbers from <code>start</code> to <code>stop</code>.
)")
);
_FPX* WINAPI xll_array_sequence(double start, double stop, double incr)
{
#pragma XLLEXPORT
	static xll::FPX a;

	try {
		if (incr == 0) {
			incr = 1;
			if (start > stop) {
				incr = -1;
			}
		}

		unsigned n;
		if (incr > 1) {
			n = static_cast<unsigned>(incr);
			incr = (stop - start) / (n - 1);
		}
		else {
			n = 1u + static_cast<unsigned>(fabs((stop - start) / incr));
		}

		a.resize(n, 1);
		for (unsigned i = 0; i < n; ++i) {
			a[i] = start + i * incr;
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	catch (...) {
		XLL_ERROR("ARRAY.SEQUENCE: unknown exception");
	}

	return a.get();
}

AddIn xai_array_diff(
	Function(XLL_FPX, "xll_array_diff", "ARRAY.DIFF")
	.Arguments({
		Arg(XLL_FPX, "array", "is an array or handle to an array."),
		Arg(XLL_CSTRING, "_op", "is a binary operation.")
		})
	.FunctionHelp("Return adjacent differences of array.")
	.Category(CATEGORY)
	.Documentation(R"xyzyx(
Call <code>std::adjacent_difference</code> using <code>op</code>.
If <op> is missing use subtraction ("-").
)xyzyx")
);
_FPX* WINAPI xll_array_diff(_FPX* pa, xcstr op)
{
#pragma XLLEXPORT
	unsigned na = size(*pa);

	if (na == 1) {
		handle<FPX> a_(pa->array[0]);
		if (a_) {
			pa = a_->get();
		}
	}

	if (*op == 0 or *op == '-') {
		std::adjacent_difference(begin(*pa), end(*pa), begin(*pa));
	}
	if (*op == '/') {
		std::adjacent_difference(begin(*pa), end(*pa), begin(*pa), std::divides<double>{});
	}

	return pa;
}

AddIn xai_array_scan(
	Function(XLL_FPX, "xll_array_scan", "ARRAY.SCAN")
	.Arguments({
		Arg(XLL_FPX, "array", "is an array or handle to an array."),
		Arg(XLL_CSTRING, "_op", "is a binary operation")
		})
	.FunctionHelp("Return right fold of array using op.")
	.Category(CATEGORY)
	.Documentation(R"xyzyx(
If <op> is missing use addition ("+").
)xyzyx")
);
_FPX* WINAPI xll_array_scan(_FPX* pa, xcstr op)
{
#pragma XLLEXPORT
	unsigned na = size(*pa);

	if (na == 1) {
		handle<FPX> a_(pa->array[0]);
		if (a_) {
			pa = a_->get();
		}
	}

	if (*op == 0 or *op == '+') {
		for (unsigned i = 1; i < size(*pa); ++i) {
			pa->array[i] = pa->array[i - 1] + pa->array[i];
		}
	}
	if (*op == '*') {
		for (unsigned i = 1; i < size(*pa); ++i) {
			pa->array[i] = pa->array[i - 1] * pa->array[i];
		}
	}

	return pa;
}

#ifdef _DEBUG


#endif // _DEBUG
#if 0


AddIn xai_array_drop(
	Function(XLL_LPOPER, "xll_array_drop", "array.DROP")
	.Arguments({
		Arg(XLL_LPOPER, "array", "is a array or handle to a array."),
		Arg(XLL_LONG, "count", "drop items from front (count > 0) or back (count < 0)."),
		})
		.FunctionHelp("Return array with dropped items removed.")
	.Category(CATEGORY)
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
	.Category(CATEGORY)
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