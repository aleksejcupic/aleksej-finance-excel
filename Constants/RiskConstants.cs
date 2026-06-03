namespace Aleksej.Finance.Excel.Constants;

/// <summary>Names, descriptions, help topics, and argument text for the Risk metric functions.</summary>
internal static class RiskConstants
{
    public const string Help = Cat.HelpBase + "/risk/risk-metrics";

    // ── Core risk metrics ─────────────────────────────────────────────────────
    public const string SharpeName = "SHARPE_RATIO";
    public const string SharpeDesc = "Annualised Sharpe ratio: (AnnReturn - rf) / AnnVolatility.";

    public const string VarHistName = "VAR_HISTORICAL";
    public const string VarHistDesc = "Historical VaR — the loss not exceeded at the confidence level. Returns a positive number.";

    public const string CvarName = "VAR_CVAR";
    public const string CvarDesc = "Conditional VaR (Expected Shortfall) — average loss beyond the VaR threshold.";

    public const string VarParamName = "VAR_PARAMETRIC";
    public const string VarParamDesc = "Parametric VaR assuming normal distribution of returns.";

    public const string AnnReturnName = "ANN_RETURN";
    public const string AnnReturnDesc = "Annualised arithmetic mean return from daily returns.";

    public const string AnnVolName = "ANN_VOL";
    public const string AnnVolDesc = "Annualised volatility (standard deviation) from daily returns.";

    public const string MaxDrawdownName = "MAX_DRAWDOWN";
    public const string MaxDrawdownDesc = "Maximum drawdown — largest peak-to-trough percentage decline.";

    // ── Performance ratios ────────────────────────────────────────────────────
    public const string SortinoName = "RISK_SORTINO";
    public const string SortinoDesc = "Sortino ratio: excess return / downside deviation. Only penalises negative returns.";

    public const string CalmarName = "RISK_CALMAR";
    public const string CalmarDesc = "Calmar ratio: annualised return / max drawdown.";

    public const string BetaName = "RISK_BETA";
    public const string BetaDesc = "Beta — systematic risk vs benchmark. β = Cov(portfolio, benchmark) / Var(benchmark).";

    public const string AlphaName = "RISK_ALPHA";
    public const string AlphaDesc = "Jensen's Alpha: α = AnnReturn_p - [rf + β*(AnnReturn_m - rf)].";

    public const string TreynorName = "RISK_TREYNOR";
    public const string TreynorDesc = "Treynor ratio: (AnnReturn - rf) / Beta.";

    public const string TrackingErrorName = "RISK_TE";
    public const string TrackingErrorDesc = "Tracking error — annualised std dev of active returns (portfolio minus benchmark).";

    public const string InfoRatioName = "RISK_IR";
    public const string InfoRatioDesc = "Information ratio: active return / tracking error.";

    // ── Volatility models ─────────────────────────────────────────────────────
    public const string EwmaLatestName = "VOL_EWMA_LATEST";
    public const string EwmaLatestDesc = "Most recent EWMA volatility estimate (annualised) from a return series.";

    public const string GarchLongRunName = "VOL_GARCH_LONGRUN";
    public const string GarchLongRunDesc = "GARCH(1,1) long-run (unconditional) variance: VL = omega / (1 - alpha - beta).";

    public const string GarchForecastName = "VOL_GARCH_FORECAST";
    public const string GarchForecastDesc = "GARCH(1,1) N-day ahead annualised volatility forecast. Mean-reverts to long-run vol.";

    // ── Argument descriptions ─────────────────────────────────────────────────
    public const string Returns           = "Range of daily simple returns";
    public const string ReturnsExample    = "Range of daily simple returns (e.g. 0.01 = +1%)";
    public const string RfSettings        = "Annual risk-free rate (omit to use Settings default)";
    public const string RfDefault         = "Annual risk-free rate (default from Settings)";
    public const string RfMar             = "Annual risk-free rate / MAR (default from Settings)";
    public const string TradingDays       = "Trading days per year (omit to use Settings default)";
    public const string TradingDaysDefault = "Trading days per year (default from Settings)";
    public const string TradingDaysPlain  = "Trading days per year";
    public const string ConfidenceSettings = "Confidence level (omit for Settings default, e.g. 0.95)";
    public const string ConfidenceDefault = "Confidence level (default from Settings)";
    public const string PortfolioReturns  = "Range of portfolio daily returns";
    public const string BenchmarkReturns  = "Range of benchmark daily returns";
    public const string PortfolioReturnsShort = "Portfolio daily returns";
    public const string BenchmarkReturnsShort = "Benchmark daily returns";
    public const string Lambda            = "Decay factor (omit for Settings default 0.94)";
    public const string Omega             = "GARCH omega parameter";
    public const string Alpha             = "GARCH alpha parameter";
    public const string Beta              = "GARCH beta parameter";
    public const string CurrentVariance   = "Today's conditional daily variance";
    public const string NDays             = "Forecast horizon in trading days";
}
