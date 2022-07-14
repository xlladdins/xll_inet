// PnL.h - Profit and loss calculation
#pragma once

namespace fms::PnL {

	enum { Dj, Vj, Aj };

	inline const char fms_pnl_doc[] = R"xyzyx(
Given prices \(X_j\), cash flows \(C_j\), and trades \(\Gamma_j\)
define \(\Delta_j = \sum_{i < j} \Gamma_i\) to be the <em>position</em>.
The <em>value</em>, or mark-to-market, of the trading strategy is
\(V_j = (\Delta_j + \Gamma_j) X_j\) and the <em>amount</em> is
\(A_j\) = \Delta_j C_j - \Gamma_j X_j\).
)xyzyx";

	template<class X = double>
	inline void init(X x, X g, X* DVA, X v0 = 0)
	{
		DVA[Dj] = 0;
		DVA[Vj] = v0 + g * x;
		DVA[Aj] = -g * x;
	}

	// price, cash flow, gamma, previous gamma
	template<class X = double>
	inline void next(X x, X c, X g, X _g, X* DVA)
	{
		DVA[Dj] += _g;
		DVA[Vj] = (DVA[Dj] + g) * x;
		DVA[Aj] = DVA[Dj] * c - g * x;
	}

	template<class X = double>
	class blotter {
		double DVA[3];
	public:
		blotter(X x, X g, X v0 = 0)
		{
			init(x, g, DVA, v0);
		}
		blotter& next(X x, X c, X g, X _g)
		{
			next(x, c, g, _g, DVA);

			return *this;
		}
	};

}
