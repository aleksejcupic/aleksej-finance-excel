namespace Aleksej.Finance.Excel.Constants;

/// <summary>Names, descriptions, help topics, and argument text for the FRA (FRA_*) functions.</summary>
internal static class FraConstants
{
    public const string Help = Cat.HelpBase;

    public const string RateName = "FRA_RATE";
    public const string RateDesc = "Continuously compounded forward rate for [t1, t2] from zero rates. f = (r2*t2 - r1*t1)/(t2-t1).";

    public const string RateSimpleName = "FRA_RATE_SIMPLE";
    public const string RateSimpleDesc = "Simply compounded (LIBOR/SOFR style) forward rate for [t1, t2]. R_F = (exp(f*(t2-t1)) - 1)/(t2-t1).";

    public const string ValueName = "FRA_VALUE";
    public const string ValueDesc = "Present value of an FRA. Long receives R_K (FRA rate) and pays market forward rate. V = L*(R_K - R_F)*delta*exp(-r2*t2).";

    public const string SettlementName = "FRA_SETTLEMENT";
    public const string SettlementDesc = "FRA settlement cash flow at t1. Settlement = L*(R_K - R_M)*delta / (1 + R_M*delta). Positive = long profits.";

    public const string Dv01Name = "FRA_DV01";
    public const string Dv01Desc = "FRA DV01 — change in value for a 1bp parallel shift in zero rates.";

    // Argument descriptions (non-common)
    public const string R1Cont    = "Zero rate to t1 (continuous)";
    public const string R2Cont    = "Zero rate to t2 (continuous)";
    public const string R1        = "Zero rate to t1";
    public const string R2        = "Zero rate to t2";
    public const string T1Start   = "Start of forward period (years)";
    public const string T2End     = "End of forward period (years)";
    public const string T1StartShort = "Start of forward period";
    public const string T2EndShort   = "End of forward period";
    public const string T1Accrual = "Start of accrual period (years)";
    public const string T2Accrual = "End of accrual period (years)";
    public const string T1AccrualShort = "Start of accrual period";
    public const string T2AccrualShort = "End of accrual period";
    public const string Notional  = "Notional principal";
    public const string FraRateK  = "Agreed FRA rate R_K (simply compounded)";
    public const string FraRateAgreed = "Agreed FRA rate";
    public const string R1Current = "Current zero rate to t1";
    public const string R2Current = "Current zero rate to t2";
    public const string MarketRate = "Realised market rate R_M at settlement";
    public const string IsLong    = "TRUE = long (receive fixed R_K), FALSE = short";
    public const string IsLongShort = "TRUE = long FRA";

    public const bool IsLongDefault = true;
}
