using ExcelDna.Integration;
using Aleksej.Finance.Risk;
using Aleksej.Finance.Excel.Constants;
using Aleksej.Finance.Excel.Helpers;
using Aleksej.Finance.Excel.Settings;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>Risk metrics, performance ratios, CAPM attribution, and volatility models.</summary>
public static class RiskFunctions
{
    // ── Core risk metrics ─────────────────────────────────────────────────────

    [ExcelFunction(Name = RiskConstants.SharpeName, Category = Cat.Risk, IsThreadSafe = true,
        Description = RiskConstants.SharpeDesc, HelpTopic = RiskConstants.Help)]
    public static object SharpeRatio(
        [ExcelArgument(Name = "returns",     Description = RiskConstants.ReturnsExample)] object returns,
        [ExcelArgument(Name = "rf",          Description = RiskConstants.RfSettings)]     object rf,
        [ExcelArgument(Name = "tradingDays", Description = RiskConstants.TradingDays)]    object tradingDays)
        => Fn.Run(Category.PortfolioRisk, () => RiskMetrics.SharpeRatio(
               In.Vector("returns", returns),
               In.Rate("rf", rf, UserSettings.Current.DefaultRiskFreeRate),
               In.PosInt("tradingDays", tradingDays, UserSettings.Current.DefaultTradingDays)));

    [ExcelFunction(Name = RiskConstants.VarHistName, Category = Cat.Risk, IsThreadSafe = true,
        Description = RiskConstants.VarHistDesc)]
    public static object VarHistorical(
        [ExcelArgument(Name = "returns",    Description = RiskConstants.Returns)]            object returns,
        [ExcelArgument(Name = "confidence", Description = RiskConstants.ConfidenceSettings)] object confidence)
        => Fn.Run(Category.PortfolioRisk, () => RiskMetrics.HistoricalVaR(
               In.Vector("returns", returns),
               In.Prob("confidence", confidence, UserSettings.Current.DefaultConfidence)));

    [ExcelFunction(Name = RiskConstants.CvarName, Category = Cat.Risk, IsThreadSafe = true,
        Description = RiskConstants.CvarDesc)]
    public static object VarCvar(
        [ExcelArgument(Name = "returns",    Description = RiskConstants.Returns)]           object returns,
        [ExcelArgument(Name = "confidence", Description = RiskConstants.ConfidenceDefault)] object confidence)
        => Fn.Run(Category.PortfolioRisk, () => RiskMetrics.ConditionalVaR(
               In.Vector("returns", returns),
               In.Prob("confidence", confidence, UserSettings.Current.DefaultConfidence)));

    [ExcelFunction(Name = RiskConstants.VarParamName, Category = Cat.Risk, IsThreadSafe = true,
        Description = RiskConstants.VarParamDesc)]
    public static object VarParametric(
        [ExcelArgument(Name = "returns",    Description = RiskConstants.Returns)]           object returns,
        [ExcelArgument(Name = "confidence", Description = RiskConstants.ConfidenceDefault)] object confidence)
        => Fn.Run(Category.PortfolioRisk, () => RiskMetrics.ParametricVaR(
               In.Vector("returns", returns),
               In.Prob("confidence", confidence, UserSettings.Current.DefaultConfidence)));

    [ExcelFunction(Name = RiskConstants.AnnReturnName, Category = Cat.Risk, IsThreadSafe = true,
        Description = RiskConstants.AnnReturnDesc)]
    public static object AnnReturn(
        [ExcelArgument(Name = "returns",     Description = RiskConstants.Returns)]              object returns,
        [ExcelArgument(Name = "tradingDays", Description = RiskConstants.TradingDaysDefault)]   object tradingDays)
        => Fn.Run(Category.PortfolioRisk, () => RiskMetrics.AnnualisedReturn(
               In.Vector("returns", returns),
               In.PosInt("tradingDays", tradingDays, UserSettings.Current.DefaultTradingDays)));

    [ExcelFunction(Name = RiskConstants.AnnVolName, Category = Cat.Risk, IsThreadSafe = true,
        Description = RiskConstants.AnnVolDesc)]
    public static object AnnVol(
        [ExcelArgument(Name = "returns",     Description = RiskConstants.Returns)]            object returns,
        [ExcelArgument(Name = "tradingDays", Description = RiskConstants.TradingDaysDefault)] object tradingDays)
        => Fn.Run(Category.PortfolioRisk, () => RiskMetrics.AnnualisedVolatility(
               In.Vector("returns", returns),
               In.PosInt("tradingDays", tradingDays, UserSettings.Current.DefaultTradingDays)));

    [ExcelFunction(Name = RiskConstants.MaxDrawdownName, Category = Cat.Risk, IsThreadSafe = true,
        Description = RiskConstants.MaxDrawdownDesc)]
    public static object MaxDrawdown(
        [ExcelArgument(Name = "returns", Description = RiskConstants.Returns)] object returns)
        => Fn.Run(Category.PortfolioRisk, () => RiskMetrics.MaxDrawdown(In.Vector("returns", returns)));

    // ── Performance ratios ────────────────────────────────────────────────────

    [ExcelFunction(Name = RiskConstants.SortinoName, Category = Cat.Risk, IsThreadSafe = true,
        Description = RiskConstants.SortinoDesc)]
    public static object RiskSortino(
        [ExcelArgument(Name = "returns",     Description = RiskConstants.Returns)]            object returns,
        [ExcelArgument(Name = "rf",          Description = RiskConstants.RfMar)]              object rf,
        [ExcelArgument(Name = "tradingDays", Description = RiskConstants.TradingDaysDefault)] object tradingDays)
        => Fn.Run(Category.PortfolioRisk, () => RiskMetrics.SortinoRatio(
               In.Vector("returns", returns),
               In.Rate("rf", rf, UserSettings.Current.DefaultRiskFreeRate),
               In.PosInt("tradingDays", tradingDays, UserSettings.Current.DefaultTradingDays)));

    [ExcelFunction(Name = RiskConstants.CalmarName, Category = Cat.Risk, IsThreadSafe = true,
        Description = RiskConstants.CalmarDesc)]
    public static object RiskCalmar(
        [ExcelArgument(Name = "returns",     Description = RiskConstants.Returns)]            object returns,
        [ExcelArgument(Name = "tradingDays", Description = RiskConstants.TradingDaysDefault)] object tradingDays)
        => Fn.Run(Category.PortfolioRisk, () => RiskMetrics.CalmarRatio(
               In.Vector("returns", returns),
               In.PosInt("tradingDays", tradingDays, UserSettings.Current.DefaultTradingDays)));

    [ExcelFunction(Name = RiskConstants.BetaName, Category = Cat.Risk, IsThreadSafe = true,
        Description = RiskConstants.BetaDesc)]
    public static object RiskBeta(
        [ExcelArgument(Name = "portfolioReturns", Description = RiskConstants.PortfolioReturns)] object portfolioReturns,
        [ExcelArgument(Name = "benchmarkReturns", Description = RiskConstants.BenchmarkReturns)] object benchmarkReturns)
        => Fn.Run(Category.PortfolioRisk, () => RiskMetrics.Beta(
               In.Vector("portfolioReturns", portfolioReturns),
               In.Vector("benchmarkReturns", benchmarkReturns)));

    [ExcelFunction(Name = RiskConstants.AlphaName, Category = Cat.Risk, IsThreadSafe = true,
        Description = RiskConstants.AlphaDesc)]
    public static object RiskAlpha(
        [ExcelArgument(Name = "portfolioReturns", Description = RiskConstants.PortfolioReturnsShort)] object portfolioReturns,
        [ExcelArgument(Name = "benchmarkReturns", Description = RiskConstants.BenchmarkReturnsShort)] object benchmarkReturns,
        [ExcelArgument(Name = "rf",               Description = RiskConstants.RfDefault)]             object rf,
        [ExcelArgument(Name = "tradingDays",      Description = RiskConstants.TradingDaysPlain)]      object tradingDays)
        => Fn.Run(Category.PortfolioRisk, () => RiskMetrics.JensensAlpha(
               In.Vector("portfolioReturns", portfolioReturns),
               In.Vector("benchmarkReturns", benchmarkReturns),
               In.Rate("rf", rf, UserSettings.Current.DefaultRiskFreeRate),
               In.PosInt("tradingDays", tradingDays, UserSettings.Current.DefaultTradingDays)));

    [ExcelFunction(Name = RiskConstants.TreynorName, Category = Cat.Risk, IsThreadSafe = true,
        Description = RiskConstants.TreynorDesc)]
    public static object RiskTreynor(
        [ExcelArgument(Name = "portfolioReturns", Description = RiskConstants.PortfolioReturnsShort)] object portfolioReturns,
        [ExcelArgument(Name = "benchmarkReturns", Description = RiskConstants.BenchmarkReturnsShort)] object benchmarkReturns,
        [ExcelArgument(Name = "rf",               Description = RiskConstants.RfDefault)]             object rf,
        [ExcelArgument(Name = "tradingDays",      Description = RiskConstants.TradingDaysPlain)]      object tradingDays)
        => Fn.Run(Category.PortfolioRisk, () => RiskMetrics.TreynorRatio(
               In.Vector("portfolioReturns", portfolioReturns),
               In.Vector("benchmarkReturns", benchmarkReturns),
               In.Rate("rf", rf, UserSettings.Current.DefaultRiskFreeRate),
               In.PosInt("tradingDays", tradingDays, UserSettings.Current.DefaultTradingDays)));

    [ExcelFunction(Name = RiskConstants.TrackingErrorName, Category = Cat.Risk, IsThreadSafe = true,
        Description = RiskConstants.TrackingErrorDesc)]
    public static object RiskTrackingError(
        [ExcelArgument(Name = "portfolioReturns", Description = RiskConstants.PortfolioReturnsShort)] object portfolioReturns,
        [ExcelArgument(Name = "benchmarkReturns", Description = RiskConstants.BenchmarkReturnsShort)] object benchmarkReturns,
        [ExcelArgument(Name = "tradingDays",      Description = RiskConstants.TradingDaysDefault)]    object tradingDays)
        => Fn.Run(Category.PortfolioRisk, () => RiskMetrics.TrackingError(
               In.Vector("portfolioReturns", portfolioReturns),
               In.Vector("benchmarkReturns", benchmarkReturns),
               In.PosInt("tradingDays", tradingDays, UserSettings.Current.DefaultTradingDays)));

    [ExcelFunction(Name = RiskConstants.InfoRatioName, Category = Cat.Risk, IsThreadSafe = true,
        Description = RiskConstants.InfoRatioDesc)]
    public static object RiskInfoRatio(
        [ExcelArgument(Name = "portfolioReturns", Description = RiskConstants.PortfolioReturnsShort)] object portfolioReturns,
        [ExcelArgument(Name = "benchmarkReturns", Description = RiskConstants.BenchmarkReturnsShort)] object benchmarkReturns,
        [ExcelArgument(Name = "tradingDays",      Description = RiskConstants.TradingDaysDefault)]    object tradingDays)
        => Fn.Run(Category.PortfolioRisk, () => RiskMetrics.InformationRatio(
               In.Vector("portfolioReturns", portfolioReturns),
               In.Vector("benchmarkReturns", benchmarkReturns),
               In.PosInt("tradingDays", tradingDays, UserSettings.Current.DefaultTradingDays)));

    // ── Volatility models ─────────────────────────────────────────────────────

    [ExcelFunction(Name = RiskConstants.EwmaLatestName, Category = Cat.Risk, IsThreadSafe = true,
        Description = RiskConstants.EwmaLatestDesc)]
    public static object VolEwmaLatest(
        [ExcelArgument(Name = "returns", Description = RiskConstants.Returns)] object returns,
        [ExcelArgument(Name = "lambda",  Description = RiskConstants.Lambda)]  object lambda)
        => Fn.Run(Category.PortfolioRisk, () => VolatilityModels.EwmaVolatilityLatest(
               In.Vector("returns", returns),
               In.Prob("lambda", lambda, UserSettings.Current.DefaultLambda)));

    [ExcelFunction(Name = RiskConstants.GarchLongRunName, Category = Cat.Risk, IsThreadSafe = true,
        Description = RiskConstants.GarchLongRunDesc)]
    public static object VolGarchLongRun(
        [ExcelArgument(Name = "omega", Description = RiskConstants.Omega)] object omega,
        [ExcelArgument(Name = "alpha", Description = RiskConstants.Alpha)] object alpha,
        [ExcelArgument(Name = "beta",  Description = RiskConstants.Beta)]  object beta)
        => Fn.Run(Category.PortfolioRisk, () => VolatilityModels.GarchLongRunVariance(
               In.Num("omega", omega), In.Num("alpha", alpha), In.Num("beta", beta)));

    [ExcelFunction(Name = RiskConstants.GarchForecastName, Category = Cat.Risk, IsThreadSafe = true,
        Description = RiskConstants.GarchForecastDesc)]
    public static object VolGarchForecast(
        [ExcelArgument(Name = "currentVariance", Description = RiskConstants.CurrentVariance)] object currentVariance,
        [ExcelArgument(Name = "omega",           Description = RiskConstants.Omega)]           object omega,
        [ExcelArgument(Name = "alpha",           Description = RiskConstants.Alpha)]           object alpha,
        [ExcelArgument(Name = "beta",            Description = RiskConstants.Beta)]            object beta,
        [ExcelArgument(Name = "nDays",           Description = RiskConstants.NDays)]           object nDays)
        => Fn.Run(Category.PortfolioRisk, () => VolatilityModels.GarchForecast(
               In.Num("currentVariance", currentVariance), In.Num("omega", omega),
               In.Num("alpha", alpha), In.Num("beta", beta), In.PosInt("nDays", nDays)));
}
