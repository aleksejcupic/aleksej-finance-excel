namespace Aleksej.Finance.Excel.Constants;

/// <summary>Names, descriptions, help topics, and argument text for the IRS (IRS_*) functions.</summary>
internal static class IrsConstants
{
    public const string Help = Cat.HelpBase;

    public const string ValueName = "IRS_VALUE";
    public const string ValueDesc = "Fixed-for-floating IRS NPV. Value = B_fixed - B_float (receiver) or B_float - B_fixed (payer). Floating leg uses next-reset approximation.";

    public const string ParRateName = "IRS_PAR_RATE";
    public const string ParRateDesc = "Par (fair) swap rate — the fixed rate that makes the swap NPV = 0 at inception. K = (1 - P(Tn)) / sum(delta_i * P(Ti)).";

    public const string FixedLegName = "IRS_FIXED_LEG";
    public const string FixedLegDesc = "Present value of the fixed coupon leg. Set includePrincipal=TRUE to include final notional repayment (bond-equivalent).";

    public const string FloatLegName = "IRS_FLOAT_LEG";
    public const string FloatLegDesc = "Present value of the floating leg using the next-reset approximation. PV = notional*(1+nextCoupon)*exp(-r*t).";

    public const string Dv01Name = "IRS_DV01";
    public const string Dv01Desc = "IRS DV01 — change in swap value for a 1bp parallel shift in zero rates. Payer swap has negative DV01.";

    // Argument descriptions (non-common)
    public const string Notional         = "Notional principal";
    public const string FixedRate        = "Annual fixed coupon rate (e.g. 0.03 = 3%)";
    public const string FixedRateShort   = "Annual fixed coupon rate";
    public const string PaymentTimes     = "Remaining payment times in years (range)";
    public const string PaymentTimesAsc  = "Payment times in years (ascending, range)";
    public const string PaymentTimesPlain = "Payment times in years (range)";
    public const string ZeroRates        = "Zero rates at each payment time (range)";
    public const string NextFloatCoupon  = "Already-fixed floating coupon (fraction of notional)";
    public const string NextFloatCouponPeriod = "Already-fixed coupon for next period (fraction of notional)";
    public const string NextFloatCouponShort  = "Already-fixed floating coupon";
    public const string TimeToNextReset  = "Time to next floating reset date (years)";
    public const string TimeToNextResetYears = "Time to next reset date in years";
    public const string TimeToNextResetShort = "Time to next floating reset";
    public const string ZeroAtNextReset  = "Zero rate at the next reset date";
    public const string ZeroAtNextResetShort = "Zero rate at next reset";
    public const string IsPayFixed       = "TRUE = pay fixed (payer swap), FALSE = receive fixed";
    public const string IsPayFixedShort  = "TRUE = payer swap (default TRUE)";
    public const string Frequency        = "Payments per year (default from Settings)";
    public const string IncludePrincipal = "TRUE to include notional repayment (default FALSE)";

    public const bool IsPayFixedDefault = true;
}
