namespace Aleksej.Finance.Excel.Constants;

/// <summary>
/// Names, descriptions, help topics, default numeric values, and non-common
/// argument descriptions for the bond (BOND_*), yield-curve (YC_*), and
/// mortgage (MORT_*) functions.
/// </summary>
internal static class BondConstants
{
    // ── Help topics ───────────────────────────────────────────────────────────
    public const string HelpBondMath = Cat.HelpBase;

    // ── Default numeric values ────────────────────────────────────────────────
    /// <summary>Default payment frequency per year for mortgage functions.</summary>
    public const int DefaultPaymentsPerYear = 12;

    // ── BondMath ──────────────────────────────────────────────────────────────
    public const string PriceName = "BOND_PRICE";
    public const string PriceDesc = "Present value of a bond given yield to maturity.";

    public const string YtmName = "BOND_YTM";
    public const string YtmDesc = "Yield to maturity from bond price (Newton-Raphson solver).";

    public const string DurationName = "BOND_DURATION";
    public const string DurationDesc = "Macaulay duration in years — weighted average time to cash flow receipt.";

    public const string ModDurationName = "BOND_MOD_DURATION";
    public const string ModDurationDesc = "Modified duration = Macaulay / (1 + ytm/frequency). Price sensitivity to yield.";

    public const string ConvexityName = "BOND_CONVEXITY";
    public const string ConvexityDesc = "Bond convexity — second-order yield sensitivity. Improves duration approximation.";

    public const string Dv01Name = "BOND_DV01";
    public const string Dv01Desc = "DV01 — dollar value of a 1 basis point move in yield.";

    public const string PriceChangeName = "BOND_PRICE_CHANGE";
    public const string PriceChangeDesc = "Approximate price change using duration and convexity for a given yield shift.";

    // ── YieldCurve ────────────────────────────────────────────────────────────
    public const string YcDfName = "YC_DF";
    public const string YcDfDesc = "Discount factor P(0,T) from a continuously compounded zero rate. P = exp(-r*T).";

    public const string YcToContName = "YC_TO_CONT";
    public const string YcToContDesc = "Convert a periodically compounded rate to continuously compounded. r_cont = m*ln(1+R/m).";

    public const string YcFromContName = "YC_FROM_CONT";
    public const string YcFromContDesc = "Convert a continuously compounded rate to periodic compounding. R = m*(exp(r/m)-1).";

    public const string YcFwdRateName = "YC_FWD_RATE";
    public const string YcFwdRateDesc = "Continuously compounded forward rate for period [t1,t2]. f = (r2*t2 - r1*t1)/(t2-t1).";

    public const string YcInterpolateName = "YC_INTERPOLATE";
    public const string YcInterpolateDesc = "Linearly interpolate a zero rate from a zero curve at time T.";

    public const string YcParYieldName = "YC_PAR_YIELD";
    public const string YcParYieldDesc = "Par yield at a target maturity from the zero curve.";

    // ── MortgageMath ──────────────────────────────────────────────────────────
    public const string MortPaymentName = "MORT_PAYMENT";
    public const string MortPaymentDesc = "Periodic payment for a fully amortising fixed-rate loan. M = P*r*(1+r)^n/((1+r)^n-1).";

    public const string MortBalanceName = "MORT_BALANCE";
    public const string MortBalanceDesc = "Outstanding loan balance after k payments have been made.";

    public const string MortTotalInterestName = "MORT_TOTAL_INTEREST";
    public const string MortTotalInterestDesc = "Total interest paid over the life of a loan. = n*M - P.";

    public const string MortEarName = "MORT_EAR";
    public const string MortEarDesc = "Effective Annual Rate from a nominal rate compounded m times per year. EAR = (1+r/m)^m - 1.";

    // ── Argument descriptions (non-common) ────────────────────────────────────
    public const string ArgCouponRatePlain  = "Annual coupon rate";
    public const string ArgYtm              = "Annual yield to maturity";
    public const string ArgYears            = "Years to maturity";
    public const string ArgFrequencyDefault = "Coupon payments per year (default from Settings)";
    public const string ArgFrequencyPlain   = "Payments per year";

    public const string ArgPrice            = "Current bond market price";
    public const string ArgCurrentYtm       = "Current annual yield";
    public const string ArgDeltaYtm         = "Yield change (e.g. 0.01 = +100bps)";

    public const string ArgZeroRate         = "Continuously compounded zero rate";
    public const string ArgMaturityYears    = "Maturity in years";
    public const string ArgRatePeriodic     = "Periodically compounded rate (e.g. 0.05)";
    public const string ArgFreqToContDefault = "Compounding frequency per year (default: Settings)";
    public const string ArgRateContinuous   = "Continuously compounded rate";
    public const string ArgFreqTarget       = "Target compounding frequency";
    public const string ArgR1               = "Zero rate to t1";
    public const string ArgT1               = "Start of forward period (years)";
    public const string ArgR2               = "Zero rate to t2";
    public const string ArgT2               = "End of forward period (years)";
    public const string ArgMaturities       = "Zero curve maturity points (range)";
    public const string ArgZeroRates        = "Zero rates at each maturity (range)";
    public const string ArgTargetMaturityT  = "Target maturity in years";
    public const string ArgTargetMaturityPar = "Desired par yield maturity in years";
    public const string ArgFreqCouponDefault = "Coupon frequency (default: Settings)";

    public const string ArgPrincipal       = "Loan amount";
    public const string ArgPrincipalOrig    = "Original loan amount";
    public const string ArgAnnualRate       = "Annual nominal interest rate";
    public const string ArgLoanTermYears    = "Loan term in years";
    public const string ArgLoanTermTotalYears = "Total loan term in years";
    public const string ArgPaymentsMade     = "Number of payments already made";
    public const string ArgPaymentsPerYear  = "Payment frequency per year (default 12)";
    public const string ArgNominalRate      = "Nominal annual rate";
    public const string ArgCompoundingFreq  = "Compounding frequency";
}
