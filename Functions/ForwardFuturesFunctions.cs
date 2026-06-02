using ExcelDna.Integration;
using AleksejCupic.FinancialMath.Derivatives;
using Aleksej.Finance.Excel.Helpers;
using Aleksej.Finance.Excel.Settings;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>Forward and futures pricing via the cost-of-carry model (Hull Ch. 5-6).</summary>
public static class ForwardFuturesFunctions
{
    private static bool Enabled => UserSettings.Load().EnableDerivatives;
    private static string Off   => RangeHelper.DisabledMessage("Derivatives");

    [ExcelFunction(Name = "FWD_PRICE", Category = "Finance | Derivatives", IsThreadSafe = true,
        Description = "Forward price on an asset with no income. F = S * exp(r * T).")]
    public static object FwdPrice(
        [ExcelArgument(Name = "S", Description = "Current spot price")]              object s,
        [ExcelArgument(Name = "r", Description = "Continuous risk-free rate")]       object r,
        [ExcelArgument(Name = "T", Description = "Time to delivery in years")]       object t)
        => Enabled ? ForwardFutures.ForwardPrice(RangeHelper.Scalar(s), RangeHelper.Scalar(r), RangeHelper.Scalar(t))
                   : (object)Off;

    [ExcelFunction(Name = "FWD_PRICE_YIELD", Category = "Finance | Derivatives", IsThreadSafe = true,
        Description = "Forward price with continuous dividend yield q. F = S * exp((r - q) * T).")]
    public static object FwdPriceYield(
        [ExcelArgument(Name = "S", Description = "Current spot price")]              object s,
        [ExcelArgument(Name = "r", Description = "Continuous risk-free rate")]       object r,
        [ExcelArgument(Name = "q", Description = "Continuous dividend yield")]       object q,
        [ExcelArgument(Name = "T", Description = "Time to delivery in years")]       object t)
        => Enabled ? ForwardFutures.ForwardPriceWithYield(
                         RangeHelper.Scalar(s), RangeHelper.Scalar(r), RangeHelper.Scalar(q), RangeHelper.Scalar(t))
                   : (object)Off;

    [ExcelFunction(Name = "FWD_PRICE_INCOME", Category = "Finance | Derivatives", IsThreadSafe = true,
        Description = "Forward price with known discrete income (PV). F = (S - I) * exp(r * T). Use FWD_PV_INCOME to compute I.")]
    public static object FwdPriceIncome(
        [ExcelArgument(Name = "S",        Description = "Current spot price")]          object s,
        [ExcelArgument(Name = "incomesPV", Description = "Present value of income I")]  object incomesPv,
        [ExcelArgument(Name = "r",        Description = "Continuous risk-free rate")]   object r,
        [ExcelArgument(Name = "T",        Description = "Time to delivery")]            object t)
        => Enabled ? ForwardFutures.ForwardPriceWithIncome(
                         RangeHelper.Scalar(s), RangeHelper.Scalar(incomesPv), RangeHelper.Scalar(r), RangeHelper.Scalar(t))
                   : (object)Off;

    [ExcelFunction(Name = "FWD_FX", Category = "Finance | Derivatives", IsThreadSafe = true,
        Description = "FX forward price via covered interest rate parity. F = S * exp((r - rf) * T).")]
    public static object FwdFx(
        [ExcelArgument(Name = "S",  Description = "Spot rate (domestic per foreign)")]  object s,
        [ExcelArgument(Name = "r",  Description = "Domestic risk-free rate")]           object r,
        [ExcelArgument(Name = "rf", Description = "Foreign risk-free rate")]            object rf,
        [ExcelArgument(Name = "T",  Description = "Time to delivery")]                  object t)
        => Enabled ? ForwardFutures.FxForwardPrice(
                         RangeHelper.Scalar(s), RangeHelper.Scalar(r), RangeHelper.Scalar(rf), RangeHelper.Scalar(t))
                   : (object)Off;

    [ExcelFunction(Name = "FWD_COMMODITY", Category = "Finance | Derivatives", IsThreadSafe = true,
        Description = "Commodity forward price via cost-of-carry. F = S * exp((r + u - y) * T) where u = storage cost, y = convenience yield.")]
    public static object FwdCommodity(
        [ExcelArgument(Name = "S",              Description = "Current spot price")]                  object s,
        [ExcelArgument(Name = "r",              Description = "Risk-free rate")]                      object r,
        [ExcelArgument(Name = "storageCost",    Description = "Annual storage cost rate")]             object storageCost,
        [ExcelArgument(Name = "convenienceYield",Description = "Annual convenience yield")]            object convYield,
        [ExcelArgument(Name = "T",              Description = "Time to delivery")]                    object t)
        => Enabled ? ForwardFutures.ForwardPriceCommodity(
                         RangeHelper.Scalar(s), RangeHelper.Scalar(r), RangeHelper.Scalar(storageCost),
                         RangeHelper.Scalar(convYield), RangeHelper.Scalar(t))
                   : (object)Off;

    [ExcelFunction(Name = "FWD_VALUE", Category = "Finance | Derivatives", IsThreadSafe = true,
        Description = "Current value of an existing long forward position. f = (F - K) * exp(-r * T).")]
    public static object FwdValue(
        [ExcelArgument(Name = "F", Description = "Current fair forward price (use FWD_PRICE*)")] object f,
        [ExcelArgument(Name = "K", Description = "Delivery price agreed at inception")]          object k,
        [ExcelArgument(Name = "r", Description = "Continuous risk-free rate")]                   object r,
        [ExcelArgument(Name = "T", Description = "Remaining time to delivery")]                  object t)
        => Enabled ? ForwardFutures.ForwardValue(
                         RangeHelper.Scalar(f), RangeHelper.Scalar(k), RangeHelper.Scalar(r), RangeHelper.Scalar(t))
                   : (object)Off;

    [ExcelFunction(Name = "FWD_VALUE_SHORT", Category = "Finance | Derivatives", IsThreadSafe = true,
        Description = "Current value of an existing short forward position. f = (K - F) * exp(-r * T).")]
    public static object FwdValueShort(
        [ExcelArgument(Name = "F", Description = "Current fair forward price")] object f,
        [ExcelArgument(Name = "K", Description = "Delivery price")]             object k,
        [ExcelArgument(Name = "r", Description = "Risk-free rate")]             object r,
        [ExcelArgument(Name = "T", Description = "Remaining time to delivery")] object t)
        => Enabled ? ForwardFutures.ForwardValueShort(
                         RangeHelper.Scalar(f), RangeHelper.Scalar(k), RangeHelper.Scalar(r), RangeHelper.Scalar(t))
                   : (object)Off;

    [ExcelFunction(Name = "FWD_PV_INCOME", Category = "Finance | Derivatives", IsThreadSafe = true,
        Description = "Present value of discrete cash flows: I = sum(CF_i * exp(-r * t_i)). Use as input to FWD_PRICE_INCOME.")]
    public static object FwdPvIncome(
        [ExcelArgument(Name = "cashFlows", Description = "Cash flow amounts (range)")] object cashFlows,
        [ExcelArgument(Name = "times",     Description = "Cash flow times in years (range)")] object times,
        [ExcelArgument(Name = "r",         Description = "Continuous discount rate")]  object r)
        => Enabled ? ForwardFutures.PresentValueOfIncome(
                         RangeHelper.ToDoubleArray(cashFlows), RangeHelper.ToDoubleArray(times), RangeHelper.Scalar(r))
                   : (object)Off;
}
