// xll_plot.cpp - plot graph using named ranges
#include "xll/xll/xll.h"

using namespace xll;

AddIn xai_plot_timeseries(
	Macro("xll_plot_timeseries", "XLL.PLOT.TIMESERIES")
);
int WINAPI xll_plot_timeseries()
{
#pragma XLLEXPORT
	try {
		OPER ref1 = Excel(xlfActiveCell);
		OPER ref2 = Excel(xlfOffset, ref1, OPER(10), OPER(5));
		// invalid number of arguments
		OPER o = Excel(xlfCreateObject
			,OPER(5) // obj_type chart
			,ref1 // ref 1 upper left
			,OPER(0), OPER(0) // x, y offset
			,ref2 // ref2 lower right
			,OPER(0), OPER(0) // x, y offset
			,OPER(3) // xy_series scatter chart
			,Missing // fill
			,OPER(7) // gallery_num xy scatter chart
			,OPER(1) // type_num
			,OPER(true) // plot_visible
			);

	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}

	return TRUE;
}