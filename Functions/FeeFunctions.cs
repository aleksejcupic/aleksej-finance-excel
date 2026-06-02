using ExcelDna.Integration;
using AleksejCupic.FinancialMath.Portfolio;
using Aleksej.Finance.Excel.Helpers;
using Aleksej.Finance.Excel.Settings;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>Fee calculations: management fees, performance fees, carried interest, expense drag.</summary>
public static class FeeFunctions
{
    private static bool Enabled => UserSettings.Load().EnableFeesAttribution;
    private static string Off   => RangeHelper.DisabledMessage("Fees & Attribution");

    [ExcelFunction(Name = "FEE_MGMT", Category = "Finance | Fees", IsThreadSafe = true,
        Description = "Management fee accrued over a period. Fee = AUM × annualRate × (days/daysInYear).")]
    public static object FeeMgmt(
        [ExcelArgument(Name = "aum",          Description = "Assets under management (beginning of period)")] object aum,
        [ExcelArgument(Name = "annualRate",   Description = "Annual management fee rate (e.g. 0.02 = 2%)")]   object annualRate,
        [ExcelArgument(Name = "daysInPeriod", Description = "Number of days in the accrual period")]           object daysInPeriod,
        [ExcelArgument(Name = "daysInYear",   Description = "Day count convention (default 365)")]             object daysInYear)
        => Enabled ? FeeCalculations.ManagementFee(
                         RangeHelper.Scalar(aum), RangeHelper.Scalar(annualRate), RangeHelper.Scalar(daysInPeriod),
                         RangeHelper.IsMissing(daysInYear) ? 365 : RangeHelper.Scalar(daysInYear))
                   : (object)Off;

    [ExcelFunction(Name = "FEE_PERF", Category = "Finance | Fees", IsThreadSafe = true,
        Description = "Performance fee with high-water mark and hurdle. Fee = max(0, (NAV - effectiveHWM) × carryRate). effectiveHWM = max(HWM, previousNAV*(1+hurdle)).")]
    public static object FeePerf(
        [ExcelArgument(Name = "currentNav",    Description = "Current NAV per share (or total)")]              object currentNav,
        [ExcelArgument(Name = "highWaterMark", Description = "Previous peak NAV")]                              object highWaterMark,
        [ExcelArgument(Name = "previousNav",   Description = "NAV at start of performance period")]             object previousNav,
        [ExcelArgument(Name = "carryRate",     Description = "Performance fee rate (default 0.20 = 20%)")]      object carryRate,
        [ExcelArgument(Name = "hurdleRate",    Description = "Minimum return before carry (default 0 = none)")] object hurdleRate)
        => Enabled ? FeeCalculations.PerformanceFee(
                         RangeHelper.Scalar(currentNav), RangeHelper.Scalar(highWaterMark), RangeHelper.Scalar(previousNav),
                         RangeHelper.IsMissing(carryRate) ? 0.20 : RangeHelper.Scalar(carryRate),
                         RangeHelper.IsMissing(hurdleRate) ? 0.0 : RangeHelper.Scalar(hurdleRate))
                   : (object)Off;

    [ExcelFunction(Name = "FEE_HWM", Category = "Finance | Fees", IsThreadSafe = true,
        Description = "High-water mark — running maximum NAV from a series. Performance fees only charged on gains above this level.")]
    public static object FeeHwm(
        [ExcelArgument(Name = "navSeries", Description = "NAV values in chronological order (range)")] object navSeries)
        => Enabled ? FeeCalculations.HighWaterMark(RangeHelper.ToDoubleArray(navSeries)) : (object)Off;

    [ExcelFunction(Name = "FEE_EXPENSE_DRAG", Category = "Finance | Fees", IsThreadSafe = true,
        Description = "Total wealth lost to annual expense ratio over N years. Impact = (1+gross)^n - (1+net)^n.")]
    public static object FeeExpenseDrag(
        [ExcelArgument(Name = "grossReturn",   Description = "Gross annual return (e.g. 0.08 = 8%)")]      object grossReturn,
        [ExcelArgument(Name = "expenseRatio",  Description = "Annual expense ratio (e.g. 0.01 = 1%)")]     object expenseRatio,
        [ExcelArgument(Name = "years",         Description = "Investment horizon in years")]                object years)
        => Enabled ? FeeCalculations.CumulativeFeeImpact(
                         RangeHelper.Scalar(grossReturn), RangeHelper.Scalar(expenseRatio), RangeHelper.Scalar(years))
                   : (object)Off;

    [ExcelFunction(Name = "FEE_NET_RETURN", Category = "Finance | Fees", IsThreadSafe = true,
        Description = "Net return after an annual expense ratio. netReturn = (1+gross)/(1+expense) - 1.")]
    public static object FeeNetReturn(
        [ExcelArgument(Name = "grossReturn",  Description = "Gross annual return")] object grossReturn,
        [ExcelArgument(Name = "expenseRatio", Description = "Annual expense ratio")] object expenseRatio)
        => Enabled ? FeeCalculations.NetReturnAfterFees(RangeHelper.Scalar(grossReturn), RangeHelper.Scalar(expenseRatio))
                   : (object)Off;

    [ExcelFunction(Name = "FEE_CARRIED_INT", Category = "Finance | Fees", IsThreadSafe = true,
        Description = "PE carried interest: GP share of profits above the preferred return hurdle. Carry = carryRate × max(0, distributions - investedCapital*(1+prefReturn)^holdYears).")]
    public static object FeeCarriedInt(
        [ExcelArgument(Name = "totalDistributions", Description = "Total cash returned to LPs")]                object totalDistributions,
        [ExcelArgument(Name = "investedCapital",    Description = "Total LP invested capital")]                 object investedCapital,
        [ExcelArgument(Name = "preferredReturn",    Description = "Hurdle rate per year (e.g. 0.08 = 8%)")]    object preferredReturn,
        [ExcelArgument(Name = "holdingYears",       Description = "Number of years capital was held")]          object holdingYears,
        [ExcelArgument(Name = "carryRate",          Description = "GP carry rate (default 0.20 = 20%)")]       object carryRate)
        => Enabled ? FeeCalculations.CarriedInterest(
                         RangeHelper.Scalar(totalDistributions), RangeHelper.Scalar(investedCapital),
                         RangeHelper.IsMissing(preferredReturn) ? 0.08 : RangeHelper.Scalar(preferredReturn),
                         RangeHelper.IsMissing(holdingYears) ? 1.0 : RangeHelper.Scalar(holdingYears),
                         RangeHelper.IsMissing(carryRate) ? 0.20 : RangeHelper.Scalar(carryRate))
                   : (object)Off;

    [ExcelFunction(Name = "FEE_TRANSACTION_COST", Category = "Finance | Fees", IsThreadSafe = true,
        Description = "Total round-trip transaction cost: commission + half the bid-ask spread. Cost = value*commission + value*spreadBps/20000.")]
    public static object FeeTransactionCost(
        [ExcelArgument(Name = "tradeValue",     Description = "Notional value of the trade")]                  object tradeValue,
        [ExcelArgument(Name = "commissionRate", Description = "Commission as a fraction (e.g. 0.001 = 10bps)")] object commissionRate,
        [ExcelArgument(Name = "spreadBps",      Description = "Bid-ask spread in basis points (e.g. 2 = 2bps)")] object spreadBps)
        => Enabled ? FeeCalculations.TotalTransactionCost(
                         RangeHelper.Scalar(tradeValue), RangeHelper.Scalar(commissionRate), RangeHelper.Scalar(spreadBps))
                   : (object)Off;
}
