// xll_PnL.cpp - Compute profit and loss.
#include "fms_PnL.h"
#include "xll/xll/xll.h"

using namespace fms;
using namespace xll;

AddIn xai_PnL(
	Function(XLL_FPX, "xll_PnL", "PnL")
	.Arguments({
		Arg(XLL_FPX, "X", "is an array of prices."),
		Arg(XLL_FPX, "C", "is an array of cash flows."),
		Arg(XLL_FPX, "Gamma", "is an array of trade amounts."),
		Arg(XLL_DOUBLE, "_V0", "is an optional initial account value.")
		})
	.Category("XLL")
	.FunctionHelp("Return position, value, and amount associated with the trading strategy.")
);
_FPX* WINAPI xll_PnL(_FPX* pX, _FPX* pC, _FPX* pG, double V0)
{
#pragma XLLEXPORT
	static FPX DVA;

	try {
		unsigned n = size(*pX);
		ensure(n == size(*pC));
		ensure(n == size(*pG));

		const FPX& X = *pX;
		const FPX& C = *pC;
		const FPX& G = *pG;

		DVA.resize(n, 3);
		PnL::init(X[0], G[0], &DVA[0], V0);
		for (unsigned i = 1; i < n; ++i) {
			PnL::next(X[i], C[i], G[i], G[i - 1], &DVA(i, 0));
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return nullptr;
	}

	return DVA.get();
}