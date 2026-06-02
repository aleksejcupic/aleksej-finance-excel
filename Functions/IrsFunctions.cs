using ExcelDna.Integration;
using AleksejCupic.FinancialMath.Derivatives;
using Aleksej.Finance.Excel.Helpers;
using Aleksej.Finance.Excel.Settings;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>Vanilla fixed-for-floating interest rate swap valuation (Hull Ch. 7).</summary>
public static class IrsFunctions
{
    private static bool Enabled => UserSettings.Load().EnableDerivatives;
    private static string Off   => RangeHelper.DisabledMessage("Derivatives");
    private static int Freq     => UserSettings.Load().DefaultFrequency;

    [ExcelFunction(Name = "IRS_VALUE", Category = "Finance | Derivatives", IsThreadSafe = true,
        Description = "Fixed-for-floating IRS NPV. Value = B_fixed - B_float (receiver) or B_float - B_fixed (payer). Floating leg uses next-reset approximation.",
        HelpTopic = "https://aleksejcupic.github.io/financial-math/derivatives/interest-rate-swap")]
    public static object IrsValue(
        [ExcelArgument(Name = "notional",          Description = "Notional principal")]                            object notional,
        [ExcelArgument(Name = "fixedRate",         Description = "Annual fixed coupon rate (e.g. 0.03 = 3%)")]     object fixedRate,
        [ExcelArgument(Name = "paymentTimes",      Description = "Remaining payment times in years (range)")]      object paymentTimes,
        [ExcelArgument(Name = "zeroRates",         Description = "Zero rates at each payment time (range)")]       object zeroRates,
        [ExcelArgument(Name = "nextFloatCoupon",   Description = "Already-fixed floating coupon (fraction of notional)")] object nextFloatCoupon,
        [ExcelArgument(Name = "timeToNextReset",   Description = "Time to next floating reset date (years)")]      object timeToNextReset,
        [ExcelArgument(Name = "zeroAtNextReset",   Description = "Zero rate at the next reset date")]              object zeroAtNextReset,
        [ExcelArgument(Name = "isPayFixed",        Description = "TRUE = pay fixed (payer swap), FALSE = receive fixed")] object isPayFixed)
        => Enabled ? InterestRateSwap.SwapValue(
                         RangeHelper.Scalar(notional), RangeHelper.Scalar(fixedRate),
                         RangeHelper.ToDoubleArray(paymentTimes), RangeHelper.ToDoubleArray(zeroRates),
                         RangeHelper.Scalar(nextFloatCoupon), RangeHelper.Scalar(timeToNextReset),
                         RangeHelper.Scalar(zeroAtNextReset),
                         RangeHelper.IsMissing(isPayFixed) ? true : RangeHelper.ScalarBool(isPayFixed))
                   : (object)Off;

    [ExcelFunction(Name = "IRS_PAR_RATE", Category = "Finance | Derivatives", IsThreadSafe = true,
        Description = "Par (fair) swap rate — the fixed rate that makes the swap NPV = 0 at inception. K = (1 - P(Tn)) / sum(delta_i * P(Ti)).")]
    public static object IrsParRate(
        [ExcelArgument(Name = "paymentTimes", Description = "Payment times in years (ascending, range)")] object paymentTimes,
        [ExcelArgument(Name = "zeroRates",   Description = "Zero rates at each payment time (range)")]   object zeroRates,
        [ExcelArgument(Name = "frequency",   Description = "Payments per year (default from Settings)")]  object frequency)
        => Enabled ? InterestRateSwap.ParSwapRate(
                         RangeHelper.ToDoubleArray(paymentTimes), RangeHelper.ToDoubleArray(zeroRates),
                         RangeHelper.IsMissing(frequency) ? Freq : RangeHelper.ScalarInt(frequency))
                   : (object)Off;

    [ExcelFunction(Name = "IRS_FIXED_LEG", Category = "Finance | Derivatives", IsThreadSafe = true,
        Description = "Present value of the fixed coupon leg. Set includePrincipal=TRUE to include final notional repayment (bond-equivalent).")]
    public static object IrsFixedLeg(
        [ExcelArgument(Name = "notional",         Description = "Notional principal")]                    object notional,
        [ExcelArgument(Name = "fixedRate",        Description = "Annual fixed coupon rate")]               object fixedRate,
        [ExcelArgument(Name = "paymentTimes",     Description = "Payment times in years (range)")]         object paymentTimes,
        [ExcelArgument(Name = "zeroRates",        Description = "Zero rates at each payment time (range)")] object zeroRates,
        [ExcelArgument(Name = "frequency",        Description = "Payments per year (default from Settings)")] object frequency,
        [ExcelArgument(Name = "includePrincipal", Description = "TRUE to include notional repayment (default FALSE)")] object includePrincipal)
        => Enabled ? InterestRateSwap.FixedLegPV(
                         RangeHelper.Scalar(notional), RangeHelper.Scalar(fixedRate),
                         RangeHelper.ToDoubleArray(paymentTimes), RangeHelper.ToDoubleArray(zeroRates),
                         RangeHelper.IsMissing(frequency) ? Freq : RangeHelper.ScalarInt(frequency),
                         RangeHelper.ScalarBool(includePrincipal))
                   : (object)Off;

    [ExcelFunction(Name = "IRS_FLOAT_LEG", Category = "Finance | Derivatives", IsThreadSafe = true,
        Description = "Present value of the floating leg using the next-reset approximation. PV = notional*(1+nextCoupon)*exp(-r*t).")]
    public static object IrsFloatLeg(
        [ExcelArgument(Name = "notional",         Description = "Notional principal")]                         object notional,
        [ExcelArgument(Name = "nextFloatCoupon",  Description = "Already-fixed coupon for next period (fraction of notional)")] object nextFloatCoupon,
        [ExcelArgument(Name = "timeToNextReset",  Description = "Time to next reset date in years")]            object timeToNextReset,
        [ExcelArgument(Name = "zeroAtNextReset",  Description = "Zero rate at the next reset date")]            object zeroAtNextReset)
        => Enabled ? InterestRateSwap.FloatingLegPV(
                         RangeHelper.Scalar(notional), RangeHelper.Scalar(nextFloatCoupon),
                         RangeHelper.Scalar(timeToNextReset), RangeHelper.Scalar(zeroAtNextReset))
                   : (object)Off;

    [ExcelFunction(Name = "IRS_DV01", Category = "Finance | Derivatives", IsThreadSafe = true,
        Description = "IRS DV01 — change in swap value for a 1bp parallel shift in zero rates. Payer swap has negative DV01.")]
    public static object IrsDv01(
        [ExcelArgument(Name = "notional",        Description = "Notional principal")]                          object notional,
        [ExcelArgument(Name = "fixedRate",       Description = "Annual fixed coupon rate")]                    object fixedRate,
        [ExcelArgument(Name = "paymentTimes",    Description = "Payment times in years (range)")]              object paymentTimes,
        [ExcelArgument(Name = "zeroRates",       Description = "Zero rates at each payment time (range)")]     object zeroRates,
        [ExcelArgument(Name = "nextFloatCoupon", Description = "Already-fixed floating coupon")]               object nextFloatCoupon,
        [ExcelArgument(Name = "timeToNextReset", Description = "Time to next floating reset")]                 object timeToNextReset,
        [ExcelArgument(Name = "zeroAtNextReset", Description = "Zero rate at next reset")]                     object zeroAtNextReset,
        [ExcelArgument(Name = "isPayFixed",      Description = "TRUE = payer swap (default TRUE)")]            object isPayFixed)
        => Enabled ? InterestRateSwap.DV01(
                         RangeHelper.Scalar(notional), RangeHelper.Scalar(fixedRate),
                         RangeHelper.ToDoubleArray(paymentTimes), RangeHelper.ToDoubleArray(zeroRates),
                         RangeHelper.Scalar(nextFloatCoupon), RangeHelper.Scalar(timeToNextReset),
                         RangeHelper.Scalar(zeroAtNextReset),
                         RangeHelper.IsMissing(isPayFixed) ? true : RangeHelper.ScalarBool(isPayFixed))
                   : (object)Off;
}
