using ExcelDna.Integration;
using AleksejCupic.FinancialMath.Bonds;
using Aleksej.Finance.Excel.Helpers;
using Aleksej.Finance.Excel.Settings;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>Fixed-income functions: bond math, yield curve, and mortgage calculations.</summary>
public static class BondFunctions
{
    private static bool Enabled => UserSettings.Load().EnableBonds;
    private static string Off   => RangeHelper.DisabledMessage("Bonds");
    private static int Freq     => UserSettings.Load().DefaultFrequency;

    // ── BondMath ──────────────────────────────────────────────────────────────

    [ExcelFunction(Name = "BOND_PRICE", Category = "Finance | Bonds", IsThreadSafe = true,
        Description = "Present value of a bond given yield to maturity.",
        HelpTopic = "https://aleksejcupic.github.io/financial-math/bonds/bond-math")]
    public static object BondPrice(
        [ExcelArgument(Name = "face",       Description = "Face (par) value")]                              object face,
        [ExcelArgument(Name = "couponRate", Description = "Annual coupon rate (e.g. 0.05 = 5%)")]           object couponRate,
        [ExcelArgument(Name = "ytm",        Description = "Annual yield to maturity")]                       object ytm,
        [ExcelArgument(Name = "years",      Description = "Years to maturity")]                              object years,
        [ExcelArgument(Name = "frequency",  Description = "Coupon payments per year (default from Settings)")] object frequency)
        => Enabled ? BondMath.Price(RangeHelper.Scalar(face), RangeHelper.Scalar(couponRate),
                         RangeHelper.Scalar(ytm), RangeHelper.Scalar(years),
                         RangeHelper.IsMissing(frequency) ? Freq : RangeHelper.ScalarInt(frequency))
                   : (object)Off;

    [ExcelFunction(Name = "BOND_YTM", Category = "Finance | Bonds", IsThreadSafe = true,
        Description = "Yield to maturity from bond price (Newton-Raphson solver).")]
    public static object BondYtm(
        [ExcelArgument(Name = "price",      Description = "Current bond market price")]                     object price,
        [ExcelArgument(Name = "face",       Description = "Face (par) value")]                              object face,
        [ExcelArgument(Name = "couponRate", Description = "Annual coupon rate")]                             object couponRate,
        [ExcelArgument(Name = "years",      Description = "Years to maturity")]                              object years,
        [ExcelArgument(Name = "frequency",  Description = "Coupon payments per year (default from Settings)")] object frequency)
        => Enabled ? BondMath.YieldToMaturity(RangeHelper.Scalar(price), RangeHelper.Scalar(face),
                         RangeHelper.Scalar(couponRate), RangeHelper.Scalar(years),
                         RangeHelper.IsMissing(frequency) ? Freq : RangeHelper.ScalarInt(frequency))
                   : (object)Off;

    [ExcelFunction(Name = "BOND_DURATION", Category = "Finance | Bonds", IsThreadSafe = true,
        Description = "Macaulay duration in years — weighted average time to cash flow receipt.")]
    public static object BondDuration(
        [ExcelArgument(Name = "face",       Description = "Face (par) value")]         object face,
        [ExcelArgument(Name = "couponRate", Description = "Annual coupon rate")]        object couponRate,
        [ExcelArgument(Name = "ytm",        Description = "Annual yield to maturity")]  object ytm,
        [ExcelArgument(Name = "years",      Description = "Years to maturity")]         object years,
        [ExcelArgument(Name = "frequency",  Description = "Payments per year")]         object frequency)
        => Enabled ? BondMath.MacaulayDuration(RangeHelper.Scalar(face), RangeHelper.Scalar(couponRate),
                         RangeHelper.Scalar(ytm), RangeHelper.Scalar(years),
                         RangeHelper.IsMissing(frequency) ? Freq : RangeHelper.ScalarInt(frequency))
                   : (object)Off;

    [ExcelFunction(Name = "BOND_MOD_DURATION", Category = "Finance | Bonds", IsThreadSafe = true,
        Description = "Modified duration = Macaulay / (1 + ytm/frequency). Price sensitivity to yield.")]
    public static object BondModDuration(
        [ExcelArgument(Name = "face",       Description = "Face (par) value")]         object face,
        [ExcelArgument(Name = "couponRate", Description = "Annual coupon rate")]        object couponRate,
        [ExcelArgument(Name = "ytm",        Description = "Annual yield to maturity")]  object ytm,
        [ExcelArgument(Name = "years",      Description = "Years to maturity")]         object years,
        [ExcelArgument(Name = "frequency",  Description = "Payments per year")]         object frequency)
        => Enabled ? BondMath.ModifiedDuration(RangeHelper.Scalar(face), RangeHelper.Scalar(couponRate),
                         RangeHelper.Scalar(ytm), RangeHelper.Scalar(years),
                         RangeHelper.IsMissing(frequency) ? Freq : RangeHelper.ScalarInt(frequency))
                   : (object)Off;

    [ExcelFunction(Name = "BOND_CONVEXITY", Category = "Finance | Bonds", IsThreadSafe = true,
        Description = "Bond convexity — second-order yield sensitivity. Improves duration approximation.")]
    public static object BondConvexity(
        [ExcelArgument(Name = "face",       Description = "Face (par) value")]         object face,
        [ExcelArgument(Name = "couponRate", Description = "Annual coupon rate")]        object couponRate,
        [ExcelArgument(Name = "ytm",        Description = "Annual yield to maturity")]  object ytm,
        [ExcelArgument(Name = "years",      Description = "Years to maturity")]         object years,
        [ExcelArgument(Name = "frequency",  Description = "Payments per year")]         object frequency)
        => Enabled ? BondMath.Convexity(RangeHelper.Scalar(face), RangeHelper.Scalar(couponRate),
                         RangeHelper.Scalar(ytm), RangeHelper.Scalar(years),
                         RangeHelper.IsMissing(frequency) ? Freq : RangeHelper.ScalarInt(frequency))
                   : (object)Off;

    [ExcelFunction(Name = "BOND_DV01", Category = "Finance | Bonds", IsThreadSafe = true,
        Description = "DV01 — dollar value of a 1 basis point move in yield.")]
    public static object BondDv01(
        [ExcelArgument(Name = "face",       Description = "Face (par) value")]         object face,
        [ExcelArgument(Name = "couponRate", Description = "Annual coupon rate")]        object couponRate,
        [ExcelArgument(Name = "ytm",        Description = "Annual yield to maturity")]  object ytm,
        [ExcelArgument(Name = "years",      Description = "Years to maturity")]         object years,
        [ExcelArgument(Name = "frequency",  Description = "Payments per year")]         object frequency)
        => Enabled ? BondMath.DV01(RangeHelper.Scalar(face), RangeHelper.Scalar(couponRate),
                         RangeHelper.Scalar(ytm), RangeHelper.Scalar(years),
                         RangeHelper.IsMissing(frequency) ? Freq : RangeHelper.ScalarInt(frequency))
                   : (object)Off;

    [ExcelFunction(Name = "BOND_PRICE_CHANGE", Category = "Finance | Bonds", IsThreadSafe = true,
        Description = "Approximate price change using duration and convexity for a given yield shift.")]
    public static object BondPriceChange(
        [ExcelArgument(Name = "face",       Description = "Face (par) value")]             object face,
        [ExcelArgument(Name = "couponRate", Description = "Annual coupon rate")]            object couponRate,
        [ExcelArgument(Name = "ytm",        Description = "Current annual yield")]          object ytm,
        [ExcelArgument(Name = "years",      Description = "Years to maturity")]             object years,
        [ExcelArgument(Name = "deltaYtm",   Description = "Yield change (e.g. 0.01 = +100bps)")] object deltaYtm,
        [ExcelArgument(Name = "frequency",  Description = "Payments per year")]             object frequency)
        => Enabled ? BondMath.ApproximatePriceChange(RangeHelper.Scalar(face), RangeHelper.Scalar(couponRate),
                         RangeHelper.Scalar(ytm), RangeHelper.Scalar(years), RangeHelper.Scalar(deltaYtm),
                         RangeHelper.IsMissing(frequency) ? Freq : RangeHelper.ScalarInt(frequency))
                   : (object)Off;

    // ── YieldCurve ────────────────────────────────────────────────────────────

    [ExcelFunction(Name = "YC_DF", Category = "Finance | Bonds", IsThreadSafe = true,
        Description = "Discount factor P(0,T) from a continuously compounded zero rate. P = exp(-r*T).")]
    public static object YcDf(
        [ExcelArgument(Name = "zeroRate", Description = "Continuously compounded zero rate")] object zeroRate,
        [ExcelArgument(Name = "T",        Description = "Maturity in years")]                  object t)
        => Enabled ? YieldCurve.DiscountFactor(RangeHelper.Scalar(zeroRate), RangeHelper.Scalar(t))
                   : (object)Off;

    [ExcelFunction(Name = "YC_TO_CONT", Category = "Finance | Bonds", IsThreadSafe = true,
        Description = "Convert a periodically compounded rate to continuously compounded. r_cont = m*ln(1+R/m).")]
    public static object YcToCont(
        [ExcelArgument(Name = "rate",      Description = "Periodically compounded rate (e.g. 0.05)")]    object rate,
        [ExcelArgument(Name = "frequency", Description = "Compounding frequency per year (default: Settings)")] object frequency)
        => Enabled ? YieldCurve.ToContinuous(RangeHelper.Scalar(rate),
                         RangeHelper.IsMissing(frequency) ? Freq : RangeHelper.ScalarInt(frequency))
                   : (object)Off;

    [ExcelFunction(Name = "YC_FROM_CONT", Category = "Finance | Bonds", IsThreadSafe = true,
        Description = "Convert a continuously compounded rate to periodic compounding. R = m*(exp(r/m)-1).")]
    public static object YcFromCont(
        [ExcelArgument(Name = "rate",      Description = "Continuously compounded rate")]        object rate,
        [ExcelArgument(Name = "frequency", Description = "Target compounding frequency")]         object frequency)
        => Enabled ? YieldCurve.FromContinuous(RangeHelper.Scalar(rate),
                         RangeHelper.IsMissing(frequency) ? Freq : RangeHelper.ScalarInt(frequency))
                   : (object)Off;

    [ExcelFunction(Name = "YC_FWD_RATE", Category = "Finance | Bonds", IsThreadSafe = true,
        Description = "Continuously compounded forward rate for period [t1,t2]. f = (r2*t2 - r1*t1)/(t2-t1).")]
    public static object YcFwdRate(
        [ExcelArgument(Name = "r1", Description = "Zero rate to t1")]   object r1,
        [ExcelArgument(Name = "t1", Description = "Start of forward period (years)")] object t1,
        [ExcelArgument(Name = "r2", Description = "Zero rate to t2")]   object r2,
        [ExcelArgument(Name = "t2", Description = "End of forward period (years)")]   object t2)
        => Enabled ? YieldCurve.ForwardRate(RangeHelper.Scalar(r1), RangeHelper.Scalar(t1),
                         RangeHelper.Scalar(r2), RangeHelper.Scalar(t2))
                   : (object)Off;

    [ExcelFunction(Name = "YC_INTERPOLATE", Category = "Finance | Bonds", IsThreadSafe = true,
        Description = "Linearly interpolate a zero rate from a zero curve at time T.")]
    public static object YcInterpolate(
        [ExcelArgument(Name = "maturities", Description = "Zero curve maturity points (range)")] object maturities,
        [ExcelArgument(Name = "zeroRates",  Description = "Zero rates at each maturity (range)")] object zeroRates,
        [ExcelArgument(Name = "T",          Description = "Target maturity in years")]             object t)
        => Enabled ? YieldCurve.InterpolateZeroRate(
                         RangeHelper.ToDoubleArray(maturities),
                         RangeHelper.ToDoubleArray(zeroRates),
                         RangeHelper.Scalar(t))
                   : (object)Off;

    [ExcelFunction(Name = "YC_PAR_YIELD", Category = "Finance | Bonds", IsThreadSafe = true,
        Description = "Par yield at a target maturity from the zero curve.")]
    public static object YcParYield(
        [ExcelArgument(Name = "maturities",     Description = "Zero curve maturity points (range)")] object maturities,
        [ExcelArgument(Name = "zeroRates",      Description = "Zero rates at each maturity (range)")] object zeroRates,
        [ExcelArgument(Name = "targetMaturity", Description = "Desired par yield maturity in years")]  object targetMaturity,
        [ExcelArgument(Name = "frequency",      Description = "Coupon frequency (default: Settings)")]  object frequency)
        => Enabled ? YieldCurve.ParYield(
                         RangeHelper.ToDoubleArray(maturities),
                         RangeHelper.ToDoubleArray(zeroRates),
                         RangeHelper.Scalar(targetMaturity),
                         RangeHelper.IsMissing(frequency) ? Freq : RangeHelper.ScalarInt(frequency))
                   : (object)Off;

    // ── MortgageMath ──────────────────────────────────────────────────────────

    [ExcelFunction(Name = "MORT_PAYMENT", Category = "Finance | Bonds", IsThreadSafe = true,
        Description = "Periodic payment for a fully amortising fixed-rate loan. M = P*r*(1+r)^n/((1+r)^n-1).")]
    public static object MortPayment(
        [ExcelArgument(Name = "principal",       Description = "Loan amount")]                             object principal,
        [ExcelArgument(Name = "annualRate",      Description = "Annual nominal interest rate")]             object annualRate,
        [ExcelArgument(Name = "years",           Description = "Loan term in years")]                      object years,
        [ExcelArgument(Name = "paymentsPerYear", Description = "Payment frequency per year (default 12)")] object paymentsPerYear)
        => Enabled ? MortgageMath.Payment(RangeHelper.Scalar(principal), RangeHelper.Scalar(annualRate),
                         RangeHelper.Scalar(years),
                         RangeHelper.IsMissing(paymentsPerYear) ? 12 : RangeHelper.ScalarInt(paymentsPerYear))
                   : (object)Off;

    [ExcelFunction(Name = "MORT_BALANCE", Category = "Finance | Bonds", IsThreadSafe = true,
        Description = "Outstanding loan balance after k payments have been made.")]
    public static object MortBalance(
        [ExcelArgument(Name = "principal",       Description = "Original loan amount")]                    object principal,
        [ExcelArgument(Name = "annualRate",      Description = "Annual nominal interest rate")]             object annualRate,
        [ExcelArgument(Name = "years",           Description = "Total loan term in years")]                 object years,
        [ExcelArgument(Name = "paymentsMade",    Description = "Number of payments already made")]          object paymentsMade,
        [ExcelArgument(Name = "paymentsPerYear", Description = "Payment frequency per year (default 12)")] object paymentsPerYear)
        => Enabled ? MortgageMath.OutstandingBalance(RangeHelper.Scalar(principal), RangeHelper.Scalar(annualRate),
                         RangeHelper.Scalar(years), RangeHelper.ScalarInt(paymentsMade),
                         RangeHelper.IsMissing(paymentsPerYear) ? 12 : RangeHelper.ScalarInt(paymentsPerYear))
                   : (object)Off;

    [ExcelFunction(Name = "MORT_TOTAL_INTEREST", Category = "Finance | Bonds", IsThreadSafe = true,
        Description = "Total interest paid over the life of a loan. = n*M - P.")]
    public static object MortTotalInterest(
        [ExcelArgument(Name = "principal",       Description = "Loan amount")]                             object principal,
        [ExcelArgument(Name = "annualRate",      Description = "Annual nominal interest rate")]             object annualRate,
        [ExcelArgument(Name = "years",           Description = "Loan term in years")]                      object years,
        [ExcelArgument(Name = "paymentsPerYear", Description = "Payment frequency per year (default 12)")] object paymentsPerYear)
        => Enabled ? MortgageMath.TotalInterest(RangeHelper.Scalar(principal), RangeHelper.Scalar(annualRate),
                         RangeHelper.Scalar(years),
                         RangeHelper.IsMissing(paymentsPerYear) ? 12 : RangeHelper.ScalarInt(paymentsPerYear))
                   : (object)Off;

    [ExcelFunction(Name = "MORT_EAR", Category = "Finance | Bonds", IsThreadSafe = true,
        Description = "Effective Annual Rate from a nominal rate compounded m times per year. EAR = (1+r/m)^m - 1.")]
    public static object MortEar(
        [ExcelArgument(Name = "nominalRate", Description = "Nominal annual rate")]     object nominalRate,
        [ExcelArgument(Name = "frequency",   Description = "Compounding frequency")]   object frequency)
        => Enabled ? MortgageMath.EffectiveAnnualRate(RangeHelper.Scalar(nominalRate),
                         RangeHelper.ScalarInt(frequency))
                   : (object)Off;
}
