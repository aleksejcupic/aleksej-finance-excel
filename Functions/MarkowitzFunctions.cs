using ExcelDna.Integration;
using AleksejCupic.FinancialMath.Portfolio;
using Aleksej.Finance.Excel.Helpers;
using Aleksej.Finance.Excel.Settings;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>Markowitz mean-variance portfolio optimisation functions.</summary>
public static class MarkowitzFunctions
{
    private static bool Enabled => UserSettings.Load().EnablePortfolioRisk;
    private static string Off   => RangeHelper.DisabledMessage("Portfolio & Risk");
    private static UserSettings Cfg => UserSettings.Load();

    [ExcelFunction(Name = "PORT_RETURN", Category = "Finance | Portfolio", IsThreadSafe = true,
        Description = "Portfolio expected return: w · mu (dot product of weights and mean returns).",
        HelpTopic = "https://aleksejcupic.github.io/financial-math/portfolio/markowitz")]
    public static object PortReturn(
        [ExcelArgument(Name = "weights", Description = "Portfolio weights (range, should sum to 1)")] object weights,
        [ExcelArgument(Name = "mu",      Description = "Annualised mean returns per asset (range)")]  object mu)
        => Enabled ? Markowitz.PortfolioReturn(RangeHelper.ToDoubleArray(weights), RangeHelper.ToDoubleArray(mu))
                   : (object)Off;

    [ExcelFunction(Name = "PORT_VOL", Category = "Finance | Portfolio", IsThreadSafe = true,
        Description = "Portfolio volatility: sqrt(w' * Sigma * w). Sigma must be a square covariance matrix range.")]
    public static object PortVol(
        [ExcelArgument(Name = "weights", Description = "Portfolio weights (range)")] object weights,
        [ExcelArgument(Name = "cov",     Description = "Annualised covariance matrix (square range)")] object cov)
        => Enabled ? Markowitz.PortfolioVolatility(RangeHelper.ToDoubleArray(weights), RangeHelper.ToDoubleMatrix(cov))
                   : (object)Off;

    [ExcelFunction(Name = "PORT_SHARPE", Category = "Finance | Portfolio", IsThreadSafe = true,
        Description = "Portfolio Sharpe ratio: (w'mu - rf) / sqrt(w'Sigma*w).")]
    public static object PortSharpe(
        [ExcelArgument(Name = "weights", Description = "Portfolio weights (range)")] object weights,
        [ExcelArgument(Name = "mu",      Description = "Mean returns (range)")]       object mu,
        [ExcelArgument(Name = "cov",     Description = "Covariance matrix (range)")] object cov,
        [ExcelArgument(Name = "rf",      Description = "Annual risk-free rate (default from Settings)")] object rf)
    {
        if (!Enabled) return Off;
        var cfg = Cfg;
        return Markowitz.PortfolioSharpe(
            RangeHelper.ToDoubleArray(weights), RangeHelper.ToDoubleArray(mu),
            RangeHelper.ToDoubleMatrix(cov),
            RangeHelper.IsMissing(rf) ? cfg.DefaultRiskFreeRate : RangeHelper.Scalar(rf));
    }

    [ExcelFunction(Name = "PORT_MIN_VAR", Category = "Finance | Portfolio", IsThreadSafe = true,
        Description = "Global minimum variance portfolio weights (analytical). w* = Sigma^-1 * 1 / (1' * Sigma^-1 * 1). Returns a column of weights — use as array formula {Ctrl+Shift+Enter}.")]
    public static object PortMinVar(
        [ExcelArgument(Name = "cov", Description = "Annualised covariance matrix (square range)")] object cov)
    {
        if (!Enabled) return Off;
        var w = Markowitz.MinVariancePortfolio(RangeHelper.ToDoubleMatrix(cov));
        var result = new object[w.Length, 1];
        for (int i = 0; i < w.Length; i++) result[i, 0] = w[i];
        return result;
    }

    [ExcelFunction(Name = "PORT_MAX_SHARPE", Category = "Finance | Portfolio", IsThreadSafe = true,
        Description = "Long-only max Sharpe ratio portfolio weights via projected gradient ascent. Returns a column of weights — use as array formula.")]
    public static object PortMaxSharpe(
        [ExcelArgument(Name = "mu",  Description = "Annualised mean returns (range)")] object mu,
        [ExcelArgument(Name = "cov", Description = "Covariance matrix (range)")]       object cov,
        [ExcelArgument(Name = "rf",  Description = "Annual risk-free rate (default from Settings)")] object rf)
    {
        if (!Enabled) return Off;
        var cfg = Cfg;
        var w = Markowitz.MaxSharpePortfolioConstrained(
            RangeHelper.ToDoubleArray(mu), RangeHelper.ToDoubleMatrix(cov),
            RangeHelper.IsMissing(rf) ? cfg.DefaultRiskFreeRate : RangeHelper.Scalar(rf));
        var result = new object[w.Length, 1];
        for (int i = 0; i < w.Length; i++) result[i, 0] = w[i];
        return result;
    }

    [ExcelFunction(Name = "PORT_RISK_PARITY", Category = "Finance | Portfolio", IsThreadSafe = true,
        Description = "Risk parity (equal risk contribution) portfolio weights. Each asset contributes equally to total portfolio volatility. Returns a column of weights — use as array formula.")]
    public static object PortRiskParity(
        [ExcelArgument(Name = "cov", Description = "Annualised covariance matrix (square range)")] object cov)
    {
        if (!Enabled) return Off;
        var w = Markowitz.RiskParityPortfolio(RangeHelper.ToDoubleMatrix(cov));
        var result = new object[w.Length, 1];
        for (int i = 0; i < w.Length; i++) result[i, 0] = w[i];
        return result;
    }

    [ExcelFunction(Name = "PORT_RISK_CONTRIB", Category = "Finance | Portfolio", IsThreadSafe = true,
        Description = "Risk contributions per asset as a fraction of total portfolio volatility. Sum = 1. For risk parity, all values = 1/n. Returns a column — use as array formula.")]
    public static object PortRiskContrib(
        [ExcelArgument(Name = "weights", Description = "Portfolio weights (range)")] object weights,
        [ExcelArgument(Name = "cov",     Description = "Covariance matrix (range)")] object cov)
    {
        if (!Enabled) return Off;
        var rc = Markowitz.RiskContributions(RangeHelper.ToDoubleArray(weights), RangeHelper.ToDoubleMatrix(cov));
        var result = new object[rc.Length, 1];
        for (int i = 0; i < rc.Length; i++) result[i, 0] = rc[i];
        return result;
    }
}
