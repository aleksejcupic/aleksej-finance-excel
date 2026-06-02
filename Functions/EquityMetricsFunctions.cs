using ExcelDna.Integration;
using AleksejCupic.FinancialMath.Portfolio;
using Aleksej.Finance.Excel.Helpers;
using Aleksej.Finance.Excel.Settings;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>Equity valuation multiples, portfolio market value, P&L, and position sizing.</summary>
public static class EquityMetricsFunctions
{
    private static bool Enabled => UserSettings.Load().EnableFeesAttribution;
    private static string Off   => RangeHelper.DisabledMessage("Fees & Attribution");

    [ExcelFunction(Name = "EQ_MKTCAP", Category = "Finance | Equity", IsThreadSafe = true,
        Description = "Market capitalisation: shares outstanding × price.")]
    public static object EqMktCap(
        [ExcelArgument(Name = "shares", Description = "Shares outstanding")]  object shares,
        [ExcelArgument(Name = "price",  Description = "Current share price")] object price)
        => Enabled ? EquityMetrics.MarketCap(RangeHelper.Scalar(shares), RangeHelper.Scalar(price))
                   : (object)Off;

    [ExcelFunction(Name = "EQ_EV", Category = "Finance | Equity", IsThreadSafe = true,
        Description = "Enterprise value: EV = Market Cap + Total Debt - Cash.")]
    public static object EqEv(
        [ExcelArgument(Name = "marketCap", Description = "Market capitalisation")]  object marketCap,
        [ExcelArgument(Name = "totalDebt", Description = "Total debt")]             object totalDebt,
        [ExcelArgument(Name = "cash",      Description = "Cash and equivalents")]   object cash)
        => Enabled ? EquityMetrics.EnterpriseValue(
                         RangeHelper.Scalar(marketCap), RangeHelper.Scalar(totalDebt), RangeHelper.Scalar(cash))
                   : (object)Off;

    [ExcelFunction(Name = "EQ_PORT_VALUE", Category = "Finance | Equity", IsThreadSafe = true,
        Description = "Total portfolio market value: sum(positions_i × prices_i). Both ranges must be the same length.")]
    public static object EqPortValue(
        [ExcelArgument(Name = "positions", Description = "Shares/units held per position (range)")] object positions,
        [ExcelArgument(Name = "prices",    Description = "Current price per unit (range)")]          object prices)
        => Enabled ? EquityMetrics.PortfolioMarketValue(
                         RangeHelper.ToDoubleArray(positions), RangeHelper.ToDoubleArray(prices))
                   : (object)Off;

    [ExcelFunction(Name = "EQ_PE", Category = "Finance | Equity", IsThreadSafe = true,
        Description = "Price-to-Earnings ratio: price / EPS. Returns #NUM if EPS ≤ 0 (loss-making).")]
    public static object EqPe(
        [ExcelArgument(Name = "price", Description = "Share price")] object price,
        [ExcelArgument(Name = "eps",   Description = "Earnings per share")] object eps)
        => Enabled ? EquityMetrics.PriceToEarnings(RangeHelper.Scalar(price), RangeHelper.Scalar(eps))
                   : (object)Off;

    [ExcelFunction(Name = "EQ_PB", Category = "Finance | Equity", IsThreadSafe = true,
        Description = "Price-to-Book ratio: price / book value per share.")]
    public static object EqPb(
        [ExcelArgument(Name = "price",  Description = "Share price")]            object price,
        [ExcelArgument(Name = "bookPS", Description = "Book value per share")]   object bookPs)
        => Enabled ? EquityMetrics.PriceToBook(RangeHelper.Scalar(price), RangeHelper.Scalar(bookPs))
                   : (object)Off;

    [ExcelFunction(Name = "EQ_PS", Category = "Finance | Equity", IsThreadSafe = true,
        Description = "Price-to-Sales ratio: market cap / annual revenue.")]
    public static object EqPs(
        [ExcelArgument(Name = "marketCap", Description = "Market capitalisation")] object marketCap,
        [ExcelArgument(Name = "revenue",   Description = "Annual revenue")]         object revenue)
        => Enabled ? EquityMetrics.PriceToSales(RangeHelper.Scalar(marketCap), RangeHelper.Scalar(revenue))
                   : (object)Off;

    [ExcelFunction(Name = "EQ_EVTOEBITDA", Category = "Finance | Equity", IsThreadSafe = true,
        Description = "EV/EBITDA multiple. A key M&A and LBO valuation multiple.")]
    public static object EqEvToEbitda(
        [ExcelArgument(Name = "ev",     Description = "Enterprise value")] object ev,
        [ExcelArgument(Name = "ebitda", Description = "EBITDA")]           object ebitda)
        => Enabled ? EquityMetrics.EvToEbitda(RangeHelper.Scalar(ev), RangeHelper.Scalar(ebitda))
                   : (object)Off;

    [ExcelFunction(Name = "EQ_DIV_YIELD", Category = "Finance | Equity", IsThreadSafe = true,
        Description = "Dividend yield: annual dividend per share / price.")]
    public static object EqDivYield(
        [ExcelArgument(Name = "annualDividend", Description = "Annual dividend per share")] object annualDividend,
        [ExcelArgument(Name = "price",          Description = "Share price")]               object price)
        => Enabled ? EquityMetrics.DividendYield(RangeHelper.Scalar(annualDividend), RangeHelper.Scalar(price))
                   : (object)Off;

    [ExcelFunction(Name = "EQ_UNREAL_PNL", Category = "Finance | Equity", IsThreadSafe = true,
        Description = "Unrealised P&L: (currentPrice - avgCostBasis) × shares. Positive = gain.")]
    public static object EqUnrealPnl(
        [ExcelArgument(Name = "shares",        Description = "Shares held")]              object shares,
        [ExcelArgument(Name = "avgCostBasis",  Description = "Average cost basis per share")] object avgCostBasis,
        [ExcelArgument(Name = "currentPrice",  Description = "Current market price")]      object currentPrice)
        => Enabled ? EquityMetrics.UnrealizedPnL(
                         RangeHelper.Scalar(shares), RangeHelper.Scalar(avgCostBasis), RangeHelper.Scalar(currentPrice))
                   : (object)Off;

    [ExcelFunction(Name = "EQ_REAL_PNL", Category = "Finance | Equity", IsThreadSafe = true,
        Description = "Realised P&L: (salePrice - avgCostBasis) × shares.")]
    public static object EqRealPnl(
        [ExcelArgument(Name = "shares",       Description = "Shares sold")]               object shares,
        [ExcelArgument(Name = "avgCostBasis", Description = "Average cost basis per share")] object avgCostBasis,
        [ExcelArgument(Name = "salePrice",    Description = "Sale price per share")]       object salePrice)
        => Enabled ? EquityMetrics.RealizedPnL(
                         RangeHelper.Scalar(shares), RangeHelper.Scalar(avgCostBasis), RangeHelper.Scalar(salePrice))
                   : (object)Off;

    [ExcelFunction(Name = "EQ_KELLY", Category = "Finance | Equity", IsThreadSafe = true,
        Description = "Continuous Kelly criterion: optimal position size f* = mu / sigma². Use 0.5× (half-Kelly) in practice to reduce variance.")]
    public static object EqKelly(
        [ExcelArgument(Name = "mu",    Description = "Expected annualised return (e.g. 0.10 = 10%)")] object mu,
        [ExcelArgument(Name = "sigma", Description = "Annualised volatility")]                         object sigma)
        => Enabled ? EquityMetrics.KellyCriterion(RangeHelper.Scalar(mu), RangeHelper.Scalar(sigma))
                   : (object)Off;

    [ExcelFunction(Name = "EQ_HALF_KELLY", Category = "Finance | Equity", IsThreadSafe = true,
        Description = "Half-Kelly position size = mu / (2 * sigma²). More conservative than full Kelly, reduces variance significantly.")]
    public static object EqHalfKelly(
        [ExcelArgument(Name = "mu",    Description = "Expected annualised return")] object mu,
        [ExcelArgument(Name = "sigma", Description = "Annualised volatility")]      object sigma)
        => Enabled ? EquityMetrics.HalfKelly(RangeHelper.Scalar(mu), RangeHelper.Scalar(sigma))
                   : (object)Off;
}
