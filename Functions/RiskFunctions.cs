using ExcelDna.Integration;
using AleksejCupic.FinancialMath.Risk;
using Aleksej.Finance.Excel.Helpers;
using Aleksej.Finance.Excel.Settings;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>Risk metrics, performance ratios, CAPM attribution, and volatility models.</summary>
public static class RiskFunctions
{
    private static bool Enabled => UserSettings.Load().EnablePortfolioRisk;
    private static string Off   => RangeHelper.DisabledMessage("Portfolio & Risk");
    private static UserSettings Cfg => UserSettings.Load();

    // ── Core risk metrics ─────────────────────────────────────────────────────

    [ExcelFunction(Name = "SHARPE_RATIO", Category = "Finance | Risk", IsThreadSafe = true,
        Description = "Annualised Sharpe ratio: (AnnReturn - rf) / AnnVolatility.",
        HelpTopic = "https://aleksejcupic.github.io/financial-math/risk/risk-metrics")]
    public static object SharpeRatio(
        [ExcelArgument(Name = "returns",     Description = "Range of daily simple returns (e.g. 0.01 = +1%)")] object returns,
        [ExcelArgument(Name = "rf",          Description = "Annual risk-free rate (omit to use Settings default)")] object rf,
        [ExcelArgument(Name = "tradingDays", Description = "Trading days per year (omit to use Settings default)")] object tradingDays)
    {
        if (!Enabled) return Off;
        var cfg = Cfg;
        return RiskMetrics.SharpeRatio(RangeHelper.ToDoubleArray(returns),
            RangeHelper.IsMissing(rf) ? cfg.DefaultRiskFreeRate : RangeHelper.Scalar(rf),
            RangeHelper.IsMissing(tradingDays) ? cfg.DefaultTradingDays : RangeHelper.ScalarInt(tradingDays));
    }

    [ExcelFunction(Name = "VAR_HISTORICAL", Category = "Finance | Risk", IsThreadSafe = true,
        Description = "Historical VaR — the loss not exceeded at the confidence level. Returns a positive number.")]
    public static object VarHistorical(
        [ExcelArgument(Name = "returns",    Description = "Range of daily simple returns")]                           object returns,
        [ExcelArgument(Name = "confidence", Description = "Confidence level (omit for Settings default, e.g. 0.95)")] object confidence)
    {
        if (!Enabled) return Off;
        var cfg = Cfg;
        return RiskMetrics.HistoricalVaR(RangeHelper.ToDoubleArray(returns),
            RangeHelper.IsMissing(confidence) ? cfg.DefaultConfidence : RangeHelper.Scalar(confidence));
    }

    [ExcelFunction(Name = "VAR_CVAR", Category = "Finance | Risk", IsThreadSafe = true,
        Description = "Conditional VaR (Expected Shortfall) — average loss beyond the VaR threshold.")]
    public static object VarCvar(
        [ExcelArgument(Name = "returns",    Description = "Range of daily simple returns")]     object returns,
        [ExcelArgument(Name = "confidence", Description = "Confidence level (default from Settings)")] object confidence)
    {
        if (!Enabled) return Off;
        var cfg = Cfg;
        return RiskMetrics.ConditionalVaR(RangeHelper.ToDoubleArray(returns),
            RangeHelper.IsMissing(confidence) ? cfg.DefaultConfidence : RangeHelper.Scalar(confidence));
    }

    [ExcelFunction(Name = "VAR_PARAMETRIC", Category = "Finance | Risk", IsThreadSafe = true,
        Description = "Parametric VaR assuming normal distribution of returns.")]
    public static object VarParametric(
        [ExcelArgument(Name = "returns",    Description = "Range of daily simple returns")]     object returns,
        [ExcelArgument(Name = "confidence", Description = "Confidence level (default from Settings)")] object confidence)
    {
        if (!Enabled) return Off;
        var cfg = Cfg;
        return RiskMetrics.ParametricVaR(RangeHelper.ToDoubleArray(returns),
            RangeHelper.IsMissing(confidence) ? cfg.DefaultConfidence : RangeHelper.Scalar(confidence));
    }

    [ExcelFunction(Name = "ANN_RETURN", Category = "Finance | Risk", IsThreadSafe = true,
        Description = "Annualised arithmetic mean return from daily returns.")]
    public static object AnnReturn(
        [ExcelArgument(Name = "returns",     Description = "Range of daily simple returns")]              object returns,
        [ExcelArgument(Name = "tradingDays", Description = "Trading days per year (default from Settings)")] object tradingDays)
    {
        if (!Enabled) return Off;
        return RiskMetrics.AnnualisedReturn(RangeHelper.ToDoubleArray(returns),
            RangeHelper.IsMissing(tradingDays) ? Cfg.DefaultTradingDays : RangeHelper.ScalarInt(tradingDays));
    }

    [ExcelFunction(Name = "ANN_VOL", Category = "Finance | Risk", IsThreadSafe = true,
        Description = "Annualised volatility (standard deviation) from daily returns.")]
    public static object AnnVol(
        [ExcelArgument(Name = "returns",     Description = "Range of daily simple returns")]              object returns,
        [ExcelArgument(Name = "tradingDays", Description = "Trading days per year (default from Settings)")] object tradingDays)
    {
        if (!Enabled) return Off;
        return RiskMetrics.AnnualisedVolatility(RangeHelper.ToDoubleArray(returns),
            RangeHelper.IsMissing(tradingDays) ? Cfg.DefaultTradingDays : RangeHelper.ScalarInt(tradingDays));
    }

    [ExcelFunction(Name = "MAX_DRAWDOWN", Category = "Finance | Risk", IsThreadSafe = true,
        Description = "Maximum drawdown — largest peak-to-trough percentage decline.")]
    public static object MaxDrawdown(
        [ExcelArgument(Name = "returns", Description = "Range of daily simple returns")] object returns)
        => Enabled ? RiskMetrics.MaxDrawdown(RangeHelper.ToDoubleArray(returns)) : (object)Off;

    // ── Performance ratios ────────────────────────────────────────────────────

    [ExcelFunction(Name = "RISK_SORTINO", Category = "Finance | Risk", IsThreadSafe = true,
        Description = "Sortino ratio: excess return / downside deviation. Only penalises negative returns.")]
    public static object RiskSortino(
        [ExcelArgument(Name = "returns",     Description = "Range of daily simple returns")]               object returns,
        [ExcelArgument(Name = "rf",          Description = "Annual risk-free rate / MAR (default from Settings)")] object rf,
        [ExcelArgument(Name = "tradingDays", Description = "Trading days per year (default from Settings)")] object tradingDays)
    {
        if (!Enabled) return Off;
        var cfg = Cfg;
        return RiskMetrics.SortinoRatio(RangeHelper.ToDoubleArray(returns),
            RangeHelper.IsMissing(rf) ? cfg.DefaultRiskFreeRate : RangeHelper.Scalar(rf),
            RangeHelper.IsMissing(tradingDays) ? cfg.DefaultTradingDays : RangeHelper.ScalarInt(tradingDays));
    }

    [ExcelFunction(Name = "RISK_CALMAR", Category = "Finance | Risk", IsThreadSafe = true,
        Description = "Calmar ratio: annualised return / max drawdown.")]
    public static object RiskCalmar(
        [ExcelArgument(Name = "returns",     Description = "Range of daily simple returns")]               object returns,
        [ExcelArgument(Name = "tradingDays", Description = "Trading days per year (default from Settings)")] object tradingDays)
        => Enabled ? RiskMetrics.CalmarRatio(RangeHelper.ToDoubleArray(returns),
                         RangeHelper.IsMissing(tradingDays) ? Cfg.DefaultTradingDays : RangeHelper.ScalarInt(tradingDays))
                   : (object)Off;

    [ExcelFunction(Name = "RISK_BETA", Category = "Finance | Risk", IsThreadSafe = true,
        Description = "Beta — systematic risk vs benchmark. β = Cov(portfolio, benchmark) / Var(benchmark).")]
    public static object RiskBeta(
        [ExcelArgument(Name = "portfolioReturns", Description = "Range of portfolio daily returns")]  object portfolioReturns,
        [ExcelArgument(Name = "benchmarkReturns", Description = "Range of benchmark daily returns")]  object benchmarkReturns)
        => Enabled ? RiskMetrics.Beta(RangeHelper.ToDoubleArray(portfolioReturns),
                         RangeHelper.ToDoubleArray(benchmarkReturns))
                   : (object)Off;

    [ExcelFunction(Name = "RISK_ALPHA", Category = "Finance | Risk", IsThreadSafe = true,
        Description = "Jensen's Alpha: α = AnnReturn_p - [rf + β*(AnnReturn_m - rf)].")]
    public static object RiskAlpha(
        [ExcelArgument(Name = "portfolioReturns", Description = "Portfolio daily returns")]           object portfolioReturns,
        [ExcelArgument(Name = "benchmarkReturns", Description = "Benchmark daily returns")]           object benchmarkReturns,
        [ExcelArgument(Name = "rf",               Description = "Annual risk-free rate (default from Settings)")] object rf,
        [ExcelArgument(Name = "tradingDays",      Description = "Trading days per year")]              object tradingDays)
    {
        if (!Enabled) return Off;
        var cfg = Cfg;
        return RiskMetrics.JensensAlpha(RangeHelper.ToDoubleArray(portfolioReturns),
            RangeHelper.ToDoubleArray(benchmarkReturns),
            RangeHelper.IsMissing(rf) ? cfg.DefaultRiskFreeRate : RangeHelper.Scalar(rf),
            RangeHelper.IsMissing(tradingDays) ? cfg.DefaultTradingDays : RangeHelper.ScalarInt(tradingDays));
    }

    [ExcelFunction(Name = "RISK_TREYNOR", Category = "Finance | Risk", IsThreadSafe = true,
        Description = "Treynor ratio: (AnnReturn - rf) / Beta.")]
    public static object RiskTreynor(
        [ExcelArgument(Name = "portfolioReturns", Description = "Portfolio daily returns")]           object portfolioReturns,
        [ExcelArgument(Name = "benchmarkReturns", Description = "Benchmark daily returns")]           object benchmarkReturns,
        [ExcelArgument(Name = "rf",               Description = "Annual risk-free rate (default from Settings)")] object rf,
        [ExcelArgument(Name = "tradingDays",      Description = "Trading days per year")]              object tradingDays)
    {
        if (!Enabled) return Off;
        var cfg = Cfg;
        return RiskMetrics.TreynorRatio(RangeHelper.ToDoubleArray(portfolioReturns),
            RangeHelper.ToDoubleArray(benchmarkReturns),
            RangeHelper.IsMissing(rf) ? cfg.DefaultRiskFreeRate : RangeHelper.Scalar(rf),
            RangeHelper.IsMissing(tradingDays) ? cfg.DefaultTradingDays : RangeHelper.ScalarInt(tradingDays));
    }

    [ExcelFunction(Name = "RISK_TE", Category = "Finance | Risk", IsThreadSafe = true,
        Description = "Tracking error — annualised std dev of active returns (portfolio minus benchmark).")]
    public static object RiskTrackingError(
        [ExcelArgument(Name = "portfolioReturns", Description = "Portfolio daily returns")]          object portfolioReturns,
        [ExcelArgument(Name = "benchmarkReturns", Description = "Benchmark daily returns")]          object benchmarkReturns,
        [ExcelArgument(Name = "tradingDays",      Description = "Trading days per year (default from Settings)")] object tradingDays)
        => Enabled ? RiskMetrics.TrackingError(RangeHelper.ToDoubleArray(portfolioReturns),
                         RangeHelper.ToDoubleArray(benchmarkReturns),
                         RangeHelper.IsMissing(tradingDays) ? Cfg.DefaultTradingDays : RangeHelper.ScalarInt(tradingDays))
                   : (object)Off;

    [ExcelFunction(Name = "RISK_IR", Category = "Finance | Risk", IsThreadSafe = true,
        Description = "Information ratio: active return / tracking error.")]
    public static object RiskInfoRatio(
        [ExcelArgument(Name = "portfolioReturns", Description = "Portfolio daily returns")]          object portfolioReturns,
        [ExcelArgument(Name = "benchmarkReturns", Description = "Benchmark daily returns")]          object benchmarkReturns,
        [ExcelArgument(Name = "tradingDays",      Description = "Trading days per year (default from Settings)")] object tradingDays)
        => Enabled ? RiskMetrics.InformationRatio(RangeHelper.ToDoubleArray(portfolioReturns),
                         RangeHelper.ToDoubleArray(benchmarkReturns),
                         RangeHelper.IsMissing(tradingDays) ? Cfg.DefaultTradingDays : RangeHelper.ScalarInt(tradingDays))
                   : (object)Off;

    // ── Volatility models ─────────────────────────────────────────────────────

    [ExcelFunction(Name = "VOL_EWMA_LATEST", Category = "Finance | Risk", IsThreadSafe = true,
        Description = "Most recent EWMA volatility estimate (annualised) from a return series.")]
    public static object VolEwmaLatest(
        [ExcelArgument(Name = "returns", Description = "Range of daily simple returns")]              object returns,
        [ExcelArgument(Name = "lambda",  Description = "Decay factor (omit for Settings default 0.94)")] object lambda)
        => Enabled ? VolatilityModels.EwmaVolatilityLatest(RangeHelper.ToDoubleArray(returns),
                         RangeHelper.IsMissing(lambda) ? Cfg.DefaultLambda : RangeHelper.Scalar(lambda))
                   : (object)Off;

    [ExcelFunction(Name = "VOL_GARCH_LONGRUN", Category = "Finance | Risk", IsThreadSafe = true,
        Description = "GARCH(1,1) long-run (unconditional) variance: VL = omega / (1 - alpha - beta).")]
    public static object VolGarchLongRun(
        [ExcelArgument(Name = "omega", Description = "GARCH omega parameter")] object omega,
        [ExcelArgument(Name = "alpha", Description = "GARCH alpha parameter")] object alpha,
        [ExcelArgument(Name = "beta",  Description = "GARCH beta parameter")]  object beta)
        => Enabled ? VolatilityModels.GarchLongRunVariance(
                         RangeHelper.Scalar(omega), RangeHelper.Scalar(alpha), RangeHelper.Scalar(beta))
                   : (object)Off;

    [ExcelFunction(Name = "VOL_GARCH_FORECAST", Category = "Finance | Risk", IsThreadSafe = true,
        Description = "GARCH(1,1) N-day ahead annualised volatility forecast. Mean-reverts to long-run vol.")]
    public static object VolGarchForecast(
        [ExcelArgument(Name = "currentVariance", Description = "Today's conditional daily variance")]  object currentVariance,
        [ExcelArgument(Name = "omega",           Description = "GARCH omega parameter")]               object omega,
        [ExcelArgument(Name = "alpha",           Description = "GARCH alpha parameter")]               object alpha,
        [ExcelArgument(Name = "beta",            Description = "GARCH beta parameter")]                object beta,
        [ExcelArgument(Name = "nDays",           Description = "Forecast horizon in trading days")]    object nDays)
        => Enabled ? VolatilityModels.GarchForecast(
                         RangeHelper.Scalar(currentVariance), RangeHelper.Scalar(omega),
                         RangeHelper.Scalar(alpha), RangeHelper.Scalar(beta), RangeHelper.ScalarInt(nDays))
                   : (object)Off;
}
