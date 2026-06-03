using ExcelDna.Integration;
using Aleksej.Finance.Portfolio;
using Aleksej.Finance.Excel.Constants;
using Aleksej.Finance.Excel.Helpers;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>Equity valuation multiples, portfolio market value, P&L, and position sizing.</summary>
public static class EquityMetricsFunctions
{
    [ExcelFunction(Name = EquityMetricsConstants.MktCapName, Category = Cat.Equity, IsThreadSafe = true,
        Description = EquityMetricsConstants.MktCapDesc, HelpTopic = EquityMetricsConstants.Help)]
    public static object EqMktCap(
        [ExcelArgument(Name = "shares", Description = EquityMetricsConstants.Shares)] object shares,
        [ExcelArgument(Name = "price",  Description = EquityMetricsConstants.Price)]  object price)
        => Fn.Run(Category.FeesAttribution, () => EquityMetrics.MarketCap(
               In.Price("shares", shares), In.Price("price", price)));

    [ExcelFunction(Name = EquityMetricsConstants.EvName, Category = Cat.Equity, IsThreadSafe = true,
        Description = EquityMetricsConstants.EvDesc, HelpTopic = EquityMetricsConstants.Help)]
    public static object EqEv(
        [ExcelArgument(Name = "marketCap", Description = EquityMetricsConstants.MarketCap)] object marketCap,
        [ExcelArgument(Name = "totalDebt", Description = EquityMetricsConstants.TotalDebt)] object totalDebt,
        [ExcelArgument(Name = "cash",      Description = EquityMetricsConstants.Cash)]      object cash)
        => Fn.Run(Category.FeesAttribution, () => EquityMetrics.EnterpriseValue(
               In.Price("marketCap", marketCap), In.Num("totalDebt", totalDebt), In.Num("cash", cash)));

    [ExcelFunction(Name = EquityMetricsConstants.PortValueName, Category = Cat.Equity, IsThreadSafe = true,
        Description = EquityMetricsConstants.PortValueDesc, HelpTopic = EquityMetricsConstants.Help)]
    public static object EqPortValue(
        [ExcelArgument(Name = "positions", Description = EquityMetricsConstants.Positions)] object positions,
        [ExcelArgument(Name = "prices",    Description = EquityMetricsConstants.Prices)]    object prices)
        => Fn.Run(Category.FeesAttribution, () => EquityMetrics.PortfolioMarketValue(
               In.Vector("positions", positions), In.Vector("prices", prices)));

    [ExcelFunction(Name = EquityMetricsConstants.PeName, Category = Cat.Equity, IsThreadSafe = true,
        Description = EquityMetricsConstants.PeDesc, HelpTopic = EquityMetricsConstants.Help)]
    public static object EqPe(
        [ExcelArgument(Name = "price", Description = EquityMetricsConstants.PriceSimple)] object price,
        [ExcelArgument(Name = "eps",   Description = EquityMetricsConstants.Eps)]         object eps)
        => Fn.Run(Category.FeesAttribution, () => EquityMetrics.PriceToEarnings(
               In.Price("price", price), In.Num("eps", eps)));

    [ExcelFunction(Name = EquityMetricsConstants.PbName, Category = Cat.Equity, IsThreadSafe = true,
        Description = EquityMetricsConstants.PbDesc, HelpTopic = EquityMetricsConstants.Help)]
    public static object EqPb(
        [ExcelArgument(Name = "price",  Description = EquityMetricsConstants.PriceSimple)] object price,
        [ExcelArgument(Name = "bookPS", Description = EquityMetricsConstants.BookPs)]      object bookPs)
        => Fn.Run(Category.FeesAttribution, () => EquityMetrics.PriceToBook(
               In.Price("price", price), In.Num("bookPS", bookPs)));

    [ExcelFunction(Name = EquityMetricsConstants.PsName, Category = Cat.Equity, IsThreadSafe = true,
        Description = EquityMetricsConstants.PsDesc, HelpTopic = EquityMetricsConstants.Help)]
    public static object EqPs(
        [ExcelArgument(Name = "marketCap", Description = EquityMetricsConstants.MarketCap)] object marketCap,
        [ExcelArgument(Name = "revenue",   Description = EquityMetricsConstants.Revenue)]   object revenue)
        => Fn.Run(Category.FeesAttribution, () => EquityMetrics.PriceToSales(
               In.Price("marketCap", marketCap), In.Num("revenue", revenue)));

    [ExcelFunction(Name = EquityMetricsConstants.EvToEbitdaName, Category = Cat.Equity, IsThreadSafe = true,
        Description = EquityMetricsConstants.EvToEbitdaDesc, HelpTopic = EquityMetricsConstants.Help)]
    public static object EqEvToEbitda(
        [ExcelArgument(Name = "ev",     Description = EquityMetricsConstants.Ev)]     object ev,
        [ExcelArgument(Name = "ebitda", Description = EquityMetricsConstants.Ebitda)] object ebitda)
        => Fn.Run(Category.FeesAttribution, () => EquityMetrics.EvToEbitda(
               In.Num("ev", ev), In.Num("ebitda", ebitda)));

    [ExcelFunction(Name = EquityMetricsConstants.DivYieldName, Category = Cat.Equity, IsThreadSafe = true,
        Description = EquityMetricsConstants.DivYieldDesc, HelpTopic = EquityMetricsConstants.Help)]
    public static object EqDivYield(
        [ExcelArgument(Name = "annualDividend", Description = EquityMetricsConstants.AnnualDividend)] object annualDividend,
        [ExcelArgument(Name = "price",          Description = EquityMetricsConstants.Price)]          object price)
        => Fn.Run(Category.FeesAttribution, () => EquityMetrics.DividendYield(
               In.Num("annualDividend", annualDividend), In.Price("price", price)));

    [ExcelFunction(Name = EquityMetricsConstants.UnrealPnlName, Category = Cat.Equity, IsThreadSafe = true,
        Description = EquityMetricsConstants.UnrealPnlDesc, HelpTopic = EquityMetricsConstants.Help)]
    public static object EqUnrealPnl(
        [ExcelArgument(Name = "shares",        Description = EquityMetricsConstants.SharesHeld)]   object shares,
        [ExcelArgument(Name = "avgCostBasis",  Description = EquityMetricsConstants.AvgCostBasis)] object avgCostBasis,
        [ExcelArgument(Name = "currentPrice",  Description = EquityMetricsConstants.CurrentPrice)] object currentPrice)
        => Fn.Run(Category.FeesAttribution, () => EquityMetrics.UnrealizedPnL(
               In.Num("shares", shares), In.Price("avgCostBasis", avgCostBasis), In.Price("currentPrice", currentPrice)));

    [ExcelFunction(Name = EquityMetricsConstants.RealPnlName, Category = Cat.Equity, IsThreadSafe = true,
        Description = EquityMetricsConstants.RealPnlDesc, HelpTopic = EquityMetricsConstants.Help)]
    public static object EqRealPnl(
        [ExcelArgument(Name = "shares",       Description = EquityMetricsConstants.SharesSold)]   object shares,
        [ExcelArgument(Name = "avgCostBasis", Description = EquityMetricsConstants.AvgCostBasis)] object avgCostBasis,
        [ExcelArgument(Name = "salePrice",    Description = EquityMetricsConstants.SalePrice)]    object salePrice)
        => Fn.Run(Category.FeesAttribution, () => EquityMetrics.RealizedPnL(
               In.Num("shares", shares), In.Price("avgCostBasis", avgCostBasis), In.Price("salePrice", salePrice)));

    [ExcelFunction(Name = EquityMetricsConstants.KellyName, Category = Cat.Equity, IsThreadSafe = true,
        Description = EquityMetricsConstants.KellyDesc, HelpTopic = EquityMetricsConstants.Help)]
    public static object EqKelly(
        [ExcelArgument(Name = "mu",    Description = EquityMetricsConstants.Mu)]    object mu,
        [ExcelArgument(Name = "sigma", Description = EquityMetricsConstants.Sigma)] object sigma)
        => Fn.Run(Category.FeesAttribution, () => EquityMetrics.KellyCriterion(
               In.Num("mu", mu), In.Vol("sigma", sigma)));

    [ExcelFunction(Name = EquityMetricsConstants.HalfKellyName, Category = Cat.Equity, IsThreadSafe = true,
        Description = EquityMetricsConstants.HalfKellyDesc, HelpTopic = EquityMetricsConstants.Help)]
    public static object EqHalfKelly(
        [ExcelArgument(Name = "mu",    Description = EquityMetricsConstants.MuHalf)] object mu,
        [ExcelArgument(Name = "sigma", Description = EquityMetricsConstants.Sigma)]  object sigma)
        => Fn.Run(Category.FeesAttribution, () => EquityMetrics.HalfKelly(
               In.Num("mu", mu), In.Vol("sigma", sigma)));
}
