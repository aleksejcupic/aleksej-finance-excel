using ExcelDna.Integration;
using Aleksej.Finance.Portfolio;
using Aleksej.Finance.Excel.Constants;
using Aleksej.Finance.Excel.Helpers;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>Fee calculations: management fees, performance fees, carried interest, expense drag.</summary>
public static class FeeFunctions
{
    [ExcelFunction(Name = FeeConstants.MgmtName, Category = Cat.Fees, IsThreadSafe = true,
        Description = FeeConstants.MgmtDesc, HelpTopic = FeeConstants.Help)]
    public static object FeeMgmt(
        [ExcelArgument(Name = "aum",          Description = FeeConstants.Aum)]          object aum,
        [ExcelArgument(Name = "annualRate",   Description = FeeConstants.AnnualRate)]   object annualRate,
        [ExcelArgument(Name = "daysInPeriod", Description = FeeConstants.DaysInPeriod)] object daysInPeriod,
        [ExcelArgument(Name = "daysInYear",   Description = FeeConstants.DaysInYear)]   object daysInYear)
        => Fn.Run(Category.FeesAttribution, () => FeeCalculations.ManagementFee(
               In.Price("aum", aum), In.Rate("annualRate", annualRate), In.PosInt("daysInPeriod", daysInPeriod),
               In.PosInt("daysInYear", daysInYear, FeeConstants.DefaultDaysInYear)));

    [ExcelFunction(Name = FeeConstants.PerfName, Category = Cat.Fees, IsThreadSafe = true,
        Description = FeeConstants.PerfDesc, HelpTopic = FeeConstants.Help)]
    public static object FeePerf(
        [ExcelArgument(Name = "currentNav",    Description = FeeConstants.CurrentNav)]    object currentNav,
        [ExcelArgument(Name = "highWaterMark", Description = FeeConstants.HighWaterMark)] object highWaterMark,
        [ExcelArgument(Name = "previousNav",   Description = FeeConstants.PreviousNav)]   object previousNav,
        [ExcelArgument(Name = "carryRate",     Description = FeeConstants.CarryRatePerf)] object carryRate,
        [ExcelArgument(Name = "hurdleRate",    Description = FeeConstants.HurdleRate)]    object hurdleRate)
        => Fn.Run(Category.FeesAttribution, () => FeeCalculations.PerformanceFee(
               In.Price("currentNav", currentNav), In.Price("highWaterMark", highWaterMark), In.Price("previousNav", previousNav),
               In.Rate("carryRate", carryRate, FeeConstants.DefaultCarryRate),
               In.Rate("hurdleRate", hurdleRate, FeeConstants.DefaultHurdleRate)));

    [ExcelFunction(Name = FeeConstants.HwmName, Category = Cat.Fees, IsThreadSafe = true,
        Description = FeeConstants.HwmDesc, HelpTopic = FeeConstants.Help)]
    public static object FeeHwm(
        [ExcelArgument(Name = "navSeries", Description = FeeConstants.NavSeries)] object navSeries)
        => Fn.Run(Category.FeesAttribution, () => FeeCalculations.HighWaterMark(In.Vector("navSeries", navSeries)));

    [ExcelFunction(Name = FeeConstants.ExpenseDragName, Category = Cat.Fees, IsThreadSafe = true,
        Description = FeeConstants.ExpenseDragDesc, HelpTopic = FeeConstants.Help)]
    public static object FeeExpenseDrag(
        [ExcelArgument(Name = "grossReturn",   Description = FeeConstants.GrossReturn)]  object grossReturn,
        [ExcelArgument(Name = "expenseRatio",  Description = FeeConstants.ExpenseRatio)] object expenseRatio,
        [ExcelArgument(Name = "years",         Description = FeeConstants.Years)]        object years)
        => Fn.Run(Category.FeesAttribution, () => FeeCalculations.CumulativeFeeImpact(
               In.Rate("grossReturn", grossReturn), In.Rate("expenseRatio", expenseRatio), In.Years("years", years)));

    [ExcelFunction(Name = FeeConstants.NetReturnName, Category = Cat.Fees, IsThreadSafe = true,
        Description = FeeConstants.NetReturnDesc, HelpTopic = FeeConstants.Help)]
    public static object FeeNetReturn(
        [ExcelArgument(Name = "grossReturn",  Description = FeeConstants.GrossReturnNet)]  object grossReturn,
        [ExcelArgument(Name = "expenseRatio", Description = FeeConstants.ExpenseRatioNet)] object expenseRatio)
        => Fn.Run(Category.FeesAttribution, () => FeeCalculations.NetReturnAfterFees(
               In.Rate("grossReturn", grossReturn), In.Rate("expenseRatio", expenseRatio)));

    [ExcelFunction(Name = FeeConstants.CarriedIntName, Category = Cat.Fees, IsThreadSafe = true,
        Description = FeeConstants.CarriedIntDesc, HelpTopic = FeeConstants.Help)]
    public static object FeeCarriedInt(
        [ExcelArgument(Name = "totalDistributions", Description = FeeConstants.TotalDistributions)] object totalDistributions,
        [ExcelArgument(Name = "investedCapital",    Description = FeeConstants.InvestedCapital)]    object investedCapital,
        [ExcelArgument(Name = "preferredReturn",    Description = FeeConstants.PreferredReturn)]    object preferredReturn,
        [ExcelArgument(Name = "holdingYears",       Description = FeeConstants.HoldingYears)]       object holdingYears,
        [ExcelArgument(Name = "carryRate",          Description = FeeConstants.CarryRateCarried)]   object carryRate)
        => Fn.Run(Category.FeesAttribution, () => FeeCalculations.CarriedInterest(
               In.Price("totalDistributions", totalDistributions), In.Price("investedCapital", investedCapital),
               In.Rate("preferredReturn", preferredReturn, FeeConstants.DefaultPreferredRet),
               RangeHelper.IsMissing(holdingYears) ? FeeConstants.DefaultHoldingYears : In.Years("holdingYears", holdingYears),
               In.Rate("carryRate", carryRate, FeeConstants.DefaultCarryRate)));

    [ExcelFunction(Name = FeeConstants.TransactionCostName, Category = Cat.Fees, IsThreadSafe = true,
        Description = FeeConstants.TransactionCostDesc, HelpTopic = FeeConstants.Help)]
    public static object FeeTransactionCost(
        [ExcelArgument(Name = "tradeValue",     Description = FeeConstants.TradeValue)]     object tradeValue,
        [ExcelArgument(Name = "commissionRate", Description = FeeConstants.CommissionRate)] object commissionRate,
        [ExcelArgument(Name = "spreadBps",      Description = FeeConstants.SpreadBps)]      object spreadBps)
        => Fn.Run(Category.FeesAttribution, () => FeeCalculations.TotalTransactionCost(
               In.Price("tradeValue", tradeValue), In.Rate("commissionRate", commissionRate), In.Num("spreadBps", spreadBps)));
}
