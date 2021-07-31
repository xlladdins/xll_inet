// xll_plot.cpp - plot graph using named ranges
#include "xll/xll/xll.h"

using namespace xll;

AddIn xai_csf3(Macro("xll_csf3", "C.S.F3").Documentation(R"(
Like <code>Ctrl-Shift-F3</code> this macro defines names using
selected text. It looks for a handle in the cell below
the first item in the selected range and defines names
corresponding to the columns in the array or range
associated with the handle.
)"));
int WINAPI xll_csf3()
{
#pragma XLLEXPORT
	try {
		OPER sel = Excel(xlfSelection);
		OPER names = Excel(xlCoerce, sel);

		OPER cell = Excel(xlfOffset, sel, OPER(1), OPER(0), OPER(1), OPER(1));
		cell = Excel(xlfAbsref, OPER(REF(0, 0)), cell);
		OPER handlex = Excel(xlCoerce, cell);
		ensure(handlex.is_num());

		// !!!get.name(cell) and delete name
		// define unique name for handlex
		OPER name(std::tmpnam(nullptr));
		name = name.safe();
		Excel(xlcDefineName, name, cell, OPER(1));

		OPER index;
		{
			handle<FPX> a_(handlex.as_num());
			if (a_) {
				index = OPER("=array.index(") & name & OPER(",,");
			}
		}
		if (!index) {
			handle<OPER> r_(handlex.as_num());
			if (r_) {
				index = OPER("=range.index(") & name & OPER(",,");
			}
		}
		ensure(index);

		// num to text
		auto text = [](unsigned i) {
			return Excel(xlfText, OPER(i), OPER("0"));
		};
		for (unsigned i = 0; i < names.size(); ++i) {
			OPER namei = index & text(i) & OPER(")");
			Excel(xlcDefineName, names[i], namei, OPER(1));
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}

	return TRUE;
}


void plot_xy(const OPER& ref1, const OPER& ref2, const OPER& x, const OPER& y)
{
	ensure(x.size() == y.size());
	OPER xy(x.size(), 2);
	OPER maxy = Excel(xlfMax, y);
	ensure(maxy.is_num());

	for (unsigned i = 0; i < xy.rows(); ++i) {
		xy(i, 0) = x[i];
		xy(i, 1) = maxy.as_num() - y[i].as_num();
	}

	OPER o = Excel(xlfCreateObject
		, OPER(10) // obj_type - open polygon
		, ref1 // ref1
		, OPER(0), OPER(0) // x_offset1, y_offset1, 
		, ref2 // ref2
		, OPER(0), OPER(0) // x_offset2, y_offset2
		, xy // array
		, OPER(true) // fill
	);

}

AddIn xai_plot_timeseries(
	Macro("xll_plot_timeseries", "XPT"/*"XLL.PLOT.TIMESERIES"*/)
);
int WINAPI xll_plot_timeseries()
{
#pragma XLLEXPORT
	try {

		/*
		OPER ic = Excel(xlcWorkbookInsert, OPER(2)); // chart

		// https://xlladdins.github.io/Excel4Macros/format.charttype.html
		// FORMAT.CHARTTYPE(apply_to, group_num, dimension, type_num)
		OPER fc = Excel(xlcFormatCharttype
			, OPER(3) // apply_to - entire chart
			, Missing // group_num
			, OPER(1) // dimension - 1d
			, OPER(8) // type_num - XY (Scatter)
		);

		// CHART.ADD.DATA(ref, rowcol, titles, categories, replace, series)
		OPER cad = Excel(xlcChartAddData
			, ref // ref
			, OPER(1) // rowcol - rows
			, OPER(true) // titles - use first row
			, OPER(false) // categories - ???
			, OPER(false) // replace - ???
			, OPER(1) // series - new series
		);

		// EDIT.SERIES(series_num, name_ref, x_ref, y_ref, z_ref, plot_order)

		
		OPER cw = Excel(xlcChartWizard
			, OPER(false), // long
			, ref // data
			, OPER(7) // gallery_num - XY (scatter)
			, OPER(1) // type_num - first formatting option
			, OPER(2) // plot_by - columns
			, Missing // categories - first data series
			, Missing // ser_titles
			, Missing // legend
			, OPER("title") // title
			, OPER("x title") // x_title
			, OPER("y title") // y_title
			, Missing // z_title
			, Missing // number_cats
			, Missing // number_titles
		);
		*/

		OPER ref = Excel(xlfSelection);
		OPER ref1 = Excel(xlfOffset, ref, OPER(1), OPER(0));
		OPER ref2 = Excel(xlfOffset, ref1, OPER(10), OPER(5));

		OPER names = Excel(xlCoerce, ref);
		OPER x = Excel(xlfGetName, OPER("!") & names[0]);

		for (unsigned i = 1; i < names.size(); ++i) {
			OPER yi = Excel(xlfEvaluate, OPER("!") & names[i]);
			plot_xy(ref1, ref2, x, yi);
		}

		/*
		// invalid number of arguments
		OPER o = Excel(xlfCreateObject
			,OPER(5) // obj_type chart
			,ref1 // ref 1 upper left
			,OPER(0), OPER(0) // x, y offset
			,ref2 // ref2 lower right
			,OPER(0), OPER(0) // x, y offset
			,Missing // xy_series scatter chart
			,OPER(true) // fill
			,OPER(7) // gallery_num xy scatter chart
			,OPER(1) // type_num
			,OPER(true) // plot_visible
			);
		*/
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}
	catch (...) {
		XLL_ERROR(__FUNCTION__ ": unknown exception");

		return FALSE;
	}

	return TRUE;
}