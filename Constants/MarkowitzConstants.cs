namespace Aleksej.Finance.Excel.Constants;

/// <summary>Names, descriptions, help topics, and argument text for the Markowitz (PORT_*) functions.</summary>
internal static class MarkowitzConstants
{
    public const string Help = Cat.HelpBase;

    public const string ReturnName = "PORT_RETURN";
    public const string ReturnDesc = "Portfolio expected return: w · mu (dot product of weights and mean returns).";

    public const string VolName = "PORT_VOL";
    public const string VolDesc = "Portfolio volatility: sqrt(w' * Sigma * w). Sigma must be a square covariance matrix range.";

    public const string SharpeName = "PORT_SHARPE";
    public const string SharpeDesc = "Portfolio Sharpe ratio: (w'mu - rf) / sqrt(w'Sigma*w).";

    public const string MinVarName = "PORT_MIN_VAR";
    public const string MinVarDesc = "Global minimum variance portfolio weights (analytical). w* = Sigma^-1 * 1 / (1' * Sigma^-1 * 1). Returns a column of weights — use as array formula {Ctrl+Shift+Enter}.";

    public const string MaxSharpeName = "PORT_MAX_SHARPE";
    public const string MaxSharpeDesc = "Long-only max Sharpe ratio portfolio weights via projected gradient ascent. Returns a column of weights — use as array formula.";

    public const string RiskParityName = "PORT_RISK_PARITY";
    public const string RiskParityDesc = "Risk parity (equal risk contribution) portfolio weights. Each asset contributes equally to total portfolio volatility. Returns a column of weights — use as array formula.";

    public const string RiskContribName = "PORT_RISK_CONTRIB";
    public const string RiskContribDesc = "Risk contributions per asset as a fraction of total portfolio volatility. Sum = 1. For risk parity, all values = 1/n. Returns a column — use as array formula.";

    // ── Argument descriptions ─────────────────────────────────────────────────
    public const string WeightsSum1 = "Portfolio weights (range, should sum to 1)";
    public const string Weights     = "Portfolio weights (range)";
    public const string MuAnn       = "Annualised mean returns per asset (range)";
    public const string MuAnnShort  = "Annualised mean returns (range)";
    public const string Mu          = "Mean returns (range)";
    public const string CovAnn      = "Annualised covariance matrix (square range)";
    public const string Cov         = "Covariance matrix (range)";
    public const string Rf          = "Annual risk-free rate (default from Settings)";
}
