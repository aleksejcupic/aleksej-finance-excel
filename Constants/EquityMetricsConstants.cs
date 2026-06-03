namespace Aleksej.Finance.Excel.Constants;

/// <summary>Names, descriptions, help topics, and argument descriptions for the equity metric (EQ_*) functions.</summary>
internal static class EquityMetricsConstants
{
    public const string Help = Cat.HelpBase;

    // ── EQ_MKTCAP ─────────────────────────────────────────────────────────────
    public const string MktCapName = "EQ_MKTCAP";
    public const string MktCapDesc = "Market capitalisation: shares outstanding × price.";
    public const string Shares = "Shares outstanding";
    public const string Price  = "Current share price";

    // ── EQ_EV ─────────────────────────────────────────────────────────────────
    public const string EvName = "EQ_EV";
    public const string EvDesc = "Enterprise value: EV = Market Cap + Total Debt - Cash.";
    public const string MarketCap = "Market capitalisation";
    public const string TotalDebt = "Total debt";
    public const string Cash      = "Cash and equivalents";

    // ── EQ_PORT_VALUE ─────────────────────────────────────────────────────────
    public const string PortValueName = "EQ_PORT_VALUE";
    public const string PortValueDesc = "Total portfolio market value: sum(positions_i × prices_i). Both ranges must be the same length.";
    public const string Positions = "Shares/units held per position (range)";
    public const string Prices    = "Current price per unit (range)";

    // ── EQ_PE ─────────────────────────────────────────────────────────────────
    public const string PeName = "EQ_PE";
    public const string PeDesc = "Price-to-Earnings ratio: price / EPS. Returns #NUM if EPS ≤ 0 (loss-making).";
    public const string PriceSimple = "Share price";
    public const string Eps         = "Earnings per share";

    // ── EQ_PB ─────────────────────────────────────────────────────────────────
    public const string PbName = "EQ_PB";
    public const string PbDesc = "Price-to-Book ratio: price / book value per share.";
    public const string BookPs = "Book value per share";

    // ── EQ_PS ─────────────────────────────────────────────────────────────────
    public const string PsName = "EQ_PS";
    public const string PsDesc = "Price-to-Sales ratio: market cap / annual revenue.";
    public const string Revenue = "Annual revenue";

    // ── EQ_EVTOEBITDA ─────────────────────────────────────────────────────────
    public const string EvToEbitdaName = "EQ_EVTOEBITDA";
    public const string EvToEbitdaDesc = "EV/EBITDA multiple. A key M&A and LBO valuation multiple.";
    public const string Ev     = "Enterprise value";
    public const string Ebitda = "EBITDA";

    // ── EQ_DIV_YIELD ──────────────────────────────────────────────────────────
    public const string DivYieldName = "EQ_DIV_YIELD";
    public const string DivYieldDesc = "Dividend yield: annual dividend per share / price.";
    public const string AnnualDividend = "Annual dividend per share";

    // ── EQ_UNREAL_PNL ─────────────────────────────────────────────────────────
    public const string UnrealPnlName = "EQ_UNREAL_PNL";
    public const string UnrealPnlDesc = "Unrealised P&L: (currentPrice - avgCostBasis) × shares. Positive = gain.";
    public const string SharesHeld   = "Shares held";
    public const string AvgCostBasis = "Average cost basis per share";
    public const string CurrentPrice = "Current market price";

    // ── EQ_REAL_PNL ───────────────────────────────────────────────────────────
    public const string RealPnlName = "EQ_REAL_PNL";
    public const string RealPnlDesc = "Realised P&L: (salePrice - avgCostBasis) × shares.";
    public const string SharesSold = "Shares sold";
    public const string SalePrice  = "Sale price per share";

    // ── EQ_KELLY ──────────────────────────────────────────────────────────────
    public const string KellyName = "EQ_KELLY";
    public const string KellyDesc = "Continuous Kelly criterion: optimal position size f* = mu / sigma². Use 0.5× (half-Kelly) in practice to reduce variance.";
    public const string Mu    = "Expected annualised return (e.g. 0.10 = 10%)";
    public const string Sigma = "Annualised volatility";

    // ── EQ_HALF_KELLY ─────────────────────────────────────────────────────────
    public const string HalfKellyName = "EQ_HALF_KELLY";
    public const string HalfKellyDesc = "Half-Kelly position size = mu / (2 * sigma²). More conservative than full Kelly, reduces variance significantly.";
    public const string MuHalf = "Expected annualised return";
}
