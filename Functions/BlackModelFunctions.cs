using ExcelDna.Integration;
using AleksejCupic.FinancialMath.Derivatives;
using Aleksej.Finance.Excel.Helpers;
using Aleksej.Finance.Excel.Settings;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>Black's (1976) model for interest rate caps, floors, and swaptions (Hull Ch. 29).</summary>
public static class BlackModelFunctions
{
    private static bool Enabled => UserSettings.Load().EnableDerivatives;
    private static string Off   => RangeHelper.DisabledMessage("Derivatives");

    [ExcelFunction(Name = "BM_CAPLET", Category = "Finance | Derivatives", IsThreadSafe = true,
        Description = "Single caplet (or floorlet) price using Black's model. Pays max(F-K,0)*delta*notional at reset. Set isFloor=TRUE for floorlet.",
        HelpTopic = "https://aleksejcupic.github.io/financial-math/derivatives/black-model")]
    public static object BmCaplet(
        [ExcelArgument(Name = "notional",        Description = "Notional principal")]                           object notional,
        [ExcelArgument(Name = "forwardRate",     Description = "Forward interest rate for the period (annual)")]object forwardRate,
        [ExcelArgument(Name = "strike",          Description = "Cap/floor strike rate (annual)")]               object strike,
        [ExcelArgument(Name = "T",               Description = "Time to start of accrual period (option expiry)")] object t,
        [ExcelArgument(Name = "r",               Description = "Zero rate to T (continuously compounded)")]     object r,
        [ExcelArgument(Name = "sigma",           Description = "Black volatility of the forward rate")]         object sigma,
        [ExcelArgument(Name = "accrualFraction", Description = "Accrual period length in years (e.g. 0.5 for semi-annual)")] object accrualFraction,
        [ExcelArgument(Name = "isFloor",         Description = "TRUE for floorlet, FALSE for caplet (default)")]object isFloor)
        => Enabled ? BlackModel.CapletPrice(
                         RangeHelper.Scalar(notional), RangeHelper.Scalar(forwardRate), RangeHelper.Scalar(strike),
                         RangeHelper.Scalar(t), RangeHelper.Scalar(r), RangeHelper.Scalar(sigma),
                         RangeHelper.IsMissing(accrualFraction) ? 0.5 : RangeHelper.Scalar(accrualFraction),
                         RangeHelper.ScalarBool(isFloor))
                   : (object)Off;

    [ExcelFunction(Name = "BM_CAP", Category = "Finance | Derivatives", IsThreadSafe = true,
        Description = "Interest rate cap price — sum of caplets over the payment schedule. All arrays must be the same length.")]
    public static object BmCap(
        [ExcelArgument(Name = "notional",        Description = "Notional principal")]                    object notional,
        [ExcelArgument(Name = "strike",          Description = "Cap strike rate")]                        object strike,
        [ExcelArgument(Name = "sigma",           Description = "Flat Black volatility")]                  object sigma,
        [ExcelArgument(Name = "paymentTimes",    Description = "Reset/start times of each caplet (range)")] object paymentTimes,
        [ExcelArgument(Name = "zeroRates",       Description = "Zero rates at each payment time (range)")] object zeroRates,
        [ExcelArgument(Name = "forwardRates",    Description = "Forward rates for each period (range)")]  object forwardRates,
        [ExcelArgument(Name = "accrualFracs",    Description = "Accrual fractions for each period (range)")] object accrualFracs)
        => Enabled ? BlackModel.CapPrice(
                         RangeHelper.Scalar(notional), RangeHelper.Scalar(strike), RangeHelper.Scalar(sigma),
                         RangeHelper.ToDoubleArray(paymentTimes), RangeHelper.ToDoubleArray(zeroRates),
                         RangeHelper.ToDoubleArray(forwardRates), RangeHelper.ToDoubleArray(accrualFracs))
                   : (object)Off;

    [ExcelFunction(Name = "BM_FLOOR", Category = "Finance | Derivatives", IsThreadSafe = true,
        Description = "Interest rate floor price — sum of floorlets. Same inputs as BM_CAP.")]
    public static object BmFloor(
        [ExcelArgument(Name = "notional",     Description = "Notional principal")]                    object notional,
        [ExcelArgument(Name = "strike",       Description = "Floor strike rate")]                     object strike,
        [ExcelArgument(Name = "sigma",        Description = "Flat Black volatility")]                 object sigma,
        [ExcelArgument(Name = "paymentTimes", Description = "Reset times (range)")]                   object paymentTimes,
        [ExcelArgument(Name = "zeroRates",    Description = "Zero rates (range)")]                    object zeroRates,
        [ExcelArgument(Name = "forwardRates", Description = "Forward rates (range)")]                 object forwardRates,
        [ExcelArgument(Name = "accrualFracs", Description = "Accrual fractions (range)")]             object accrualFracs)
        => Enabled ? BlackModel.FloorPrice(
                         RangeHelper.Scalar(notional), RangeHelper.Scalar(strike), RangeHelper.Scalar(sigma),
                         RangeHelper.ToDoubleArray(paymentTimes), RangeHelper.ToDoubleArray(zeroRates),
                         RangeHelper.ToDoubleArray(forwardRates), RangeHelper.ToDoubleArray(accrualFracs))
                   : (object)Off;

    [ExcelFunction(Name = "BM_FWD_SWAP_RATE", Category = "Finance | Derivatives", IsThreadSafe = true,
        Description = "Forward swap rate R = (P(t0) - P(tn)) / sum(delta_i * P(ti)). This is the ATM swaption strike.")]
    public static object BmFwdSwapRate(
        [ExcelArgument(Name = "paymentTimes",  Description = "Swap payment times in years (range, t0 = swaption expiry)")] object paymentTimes,
        [ExcelArgument(Name = "zeroRates",     Description = "Zero rates at each payment time (range)")] object zeroRates,
        [ExcelArgument(Name = "accrualFracs",  Description = "Accrual fractions (range)")]               object accrualFracs)
        => Enabled ? BlackModel.ForwardSwapRate(
                         RangeHelper.ToDoubleArray(paymentTimes), RangeHelper.ToDoubleArray(zeroRates),
                         RangeHelper.ToDoubleArray(accrualFracs))
                   : (object)Off;

    [ExcelFunction(Name = "BM_SWAPTION", Category = "Finance | Derivatives", IsThreadSafe = true,
        Description = "Swaption price via Black's model. Payer = right to pay fixed K; Receiver = right to receive fixed K.")]
    public static object BmSwaption(
        [ExcelArgument(Name = "notional",     Description = "Notional principal")]                              object notional,
        [ExcelArgument(Name = "strike",       Description = "Fixed strike rate of the underlying swap")]         object strike,
        [ExcelArgument(Name = "T",            Description = "Swaption expiry in years (= swap start date)")]    object t,
        [ExcelArgument(Name = "sigma",        Description = "Black vol of the forward swap rate")]               object sigma,
        [ExcelArgument(Name = "paymentTimes", Description = "Swap coupon payment times (range, starting at T)")] object paymentTimes,
        [ExcelArgument(Name = "zeroRates",    Description = "Zero rates at each payment time (range)")]          object zeroRates,
        [ExcelArgument(Name = "accrualFracs", Description = "Accrual fractions (range)")]                       object accrualFracs,
        [ExcelArgument(Name = "isPayer",      Description = "TRUE = payer swaption (right to pay fixed), FALSE = receiver")] object isPayer)
        => Enabled ? BlackModel.SwaptionPrice(
                         RangeHelper.Scalar(notional), RangeHelper.Scalar(strike), RangeHelper.Scalar(t),
                         RangeHelper.Scalar(sigma),
                         RangeHelper.ToDoubleArray(paymentTimes), RangeHelper.ToDoubleArray(zeroRates),
                         RangeHelper.ToDoubleArray(accrualFracs),
                         RangeHelper.IsMissing(isPayer) ? true : RangeHelper.ScalarBool(isPayer))
                   : (object)Off;
}
