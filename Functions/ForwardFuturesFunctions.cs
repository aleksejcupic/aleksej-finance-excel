using ExcelDna.Integration;
using Aleksej.Finance.Derivatives;
using Aleksej.Finance.Excel.Constants;
using Aleksej.Finance.Excel.Helpers;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>Forward and futures pricing via the cost-of-carry model (Hull Ch. 5-6).</summary>
public static class ForwardFuturesFunctions
{
    [ExcelFunction(Name = ForwardFuturesConstants.PriceName, Category = Cat.Derivatives, IsThreadSafe = true,
        Description = ForwardFuturesConstants.PriceDesc, HelpTopic = ForwardFuturesConstants.Help)]
    public static object FwdPrice(
        [ExcelArgument(Name = "S", Description = ForwardFuturesConstants.Spot)]        object s,
        [ExcelArgument(Name = "r", Description = ForwardFuturesConstants.RContinuous)] object r,
        [ExcelArgument(Name = "T", Description = ForwardFuturesConstants.TDelivery)]   object t)
        => Fn.Run(Category.Derivatives, () => ForwardFutures.ForwardPrice(
               In.Price("S", s), In.Rate("r", r), In.Years("T", t)));

    [ExcelFunction(Name = ForwardFuturesConstants.PriceYieldName, Category = Cat.Derivatives, IsThreadSafe = true,
        Description = ForwardFuturesConstants.PriceYieldDesc, HelpTopic = ForwardFuturesConstants.Help)]
    public static object FwdPriceYield(
        [ExcelArgument(Name = "S", Description = ForwardFuturesConstants.Spot)]          object s,
        [ExcelArgument(Name = "r", Description = ForwardFuturesConstants.RContinuous)]   object r,
        [ExcelArgument(Name = "q", Description = ForwardFuturesConstants.DividendYield)] object q,
        [ExcelArgument(Name = "T", Description = ForwardFuturesConstants.TDelivery)]     object t)
        => Fn.Run(Category.Derivatives, () => ForwardFutures.ForwardPriceWithYield(
               In.Price("S", s), In.Rate("r", r), In.Rate("q", q), In.Years("T", t)));

    [ExcelFunction(Name = ForwardFuturesConstants.PriceIncomeName, Category = Cat.Derivatives, IsThreadSafe = true,
        Description = ForwardFuturesConstants.PriceIncomeDesc, HelpTopic = ForwardFuturesConstants.Help)]
    public static object FwdPriceIncome(
        [ExcelArgument(Name = "S",         Description = ForwardFuturesConstants.Spot)]            object s,
        [ExcelArgument(Name = "incomesPV", Description = ForwardFuturesConstants.IncomesPv)]       object incomesPv,
        [ExcelArgument(Name = "r",         Description = ForwardFuturesConstants.RContinuous)]     object r,
        [ExcelArgument(Name = "T",         Description = ForwardFuturesConstants.TDeliveryShort)]  object t)
        => Fn.Run(Category.Derivatives, () => ForwardFutures.ForwardPriceWithIncome(
               In.Price("S", s), In.Num("incomesPV", incomesPv), In.Rate("r", r), In.Years("T", t)));

    [ExcelFunction(Name = ForwardFuturesConstants.FxName, Category = Cat.Derivatives, IsThreadSafe = true,
        Description = ForwardFuturesConstants.FxDesc, HelpTopic = ForwardFuturesConstants.Help)]
    public static object FwdFx(
        [ExcelArgument(Name = "S",  Description = ForwardFuturesConstants.FxSpot)]         object s,
        [ExcelArgument(Name = "r",  Description = ForwardFuturesConstants.RDomestic)]      object r,
        [ExcelArgument(Name = "rf", Description = ForwardFuturesConstants.RForeign)]       object rf,
        [ExcelArgument(Name = "T",  Description = ForwardFuturesConstants.TDeliveryShort)] object t)
        => Fn.Run(Category.Derivatives, () => ForwardFutures.FxForwardPrice(
               In.Price("S", s), In.Rate("r", r), In.Rate("rf", rf), In.Years("T", t)));

    [ExcelFunction(Name = ForwardFuturesConstants.CommodityName, Category = Cat.Derivatives, IsThreadSafe = true,
        Description = ForwardFuturesConstants.CommodityDesc, HelpTopic = ForwardFuturesConstants.Help)]
    public static object FwdCommodity(
        [ExcelArgument(Name = "S",               Description = ForwardFuturesConstants.Spot)]             object s,
        [ExcelArgument(Name = "r",               Description = ForwardFuturesConstants.RiskFree)]         object r,
        [ExcelArgument(Name = "storageCost",     Description = ForwardFuturesConstants.StorageCost)]      object storageCost,
        [ExcelArgument(Name = "convenienceYield",Description = ForwardFuturesConstants.ConvenienceYield)] object convYield,
        [ExcelArgument(Name = "T",               Description = ForwardFuturesConstants.TDeliveryShort)]   object t)
        => Fn.Run(Category.Derivatives, () => ForwardFutures.ForwardPriceCommodity(
               In.Price("S", s), In.Rate("r", r), In.Rate("storageCost", storageCost),
               In.Rate("convenienceYield", convYield), In.Years("T", t)));

    [ExcelFunction(Name = ForwardFuturesConstants.ValueName, Category = Cat.Derivatives, IsThreadSafe = true,
        Description = ForwardFuturesConstants.ValueDesc, HelpTopic = ForwardFuturesConstants.Help)]
    public static object FwdValue(
        [ExcelArgument(Name = "F", Description = ForwardFuturesConstants.FairForward)]   object f,
        [ExcelArgument(Name = "K", Description = ForwardFuturesConstants.DeliveryPrice)] object k,
        [ExcelArgument(Name = "r", Description = ForwardFuturesConstants.RContinuous)]   object r,
        [ExcelArgument(Name = "T", Description = ForwardFuturesConstants.TRemaining)]    object t)
        => Fn.Run(Category.Derivatives, () => ForwardFutures.ForwardValue(
               In.Price("F", f), In.Price("K", k), In.Rate("r", r), In.Years("T", t)));

    [ExcelFunction(Name = ForwardFuturesConstants.ValueShortName, Category = Cat.Derivatives, IsThreadSafe = true,
        Description = ForwardFuturesConstants.ValueShortDesc, HelpTopic = ForwardFuturesConstants.Help)]
    public static object FwdValueShort(
        [ExcelArgument(Name = "F", Description = ForwardFuturesConstants.FairForwardShort)]   object f,
        [ExcelArgument(Name = "K", Description = ForwardFuturesConstants.DeliveryPriceShort)] object k,
        [ExcelArgument(Name = "r", Description = ForwardFuturesConstants.RiskFree)]           object r,
        [ExcelArgument(Name = "T", Description = ForwardFuturesConstants.TRemaining)]         object t)
        => Fn.Run(Category.Derivatives, () => ForwardFutures.ForwardValueShort(
               In.Price("F", f), In.Price("K", k), In.Rate("r", r), In.Years("T", t)));

    [ExcelFunction(Name = ForwardFuturesConstants.PvIncomeName, Category = Cat.Derivatives, IsThreadSafe = true,
        Description = ForwardFuturesConstants.PvIncomeDesc, HelpTopic = ForwardFuturesConstants.Help)]
    public static object FwdPvIncome(
        [ExcelArgument(Name = "cashFlows", Description = ForwardFuturesConstants.CashFlows)] object cashFlows,
        [ExcelArgument(Name = "times",     Description = ForwardFuturesConstants.Times)]     object times,
        [ExcelArgument(Name = "r",         Description = ForwardFuturesConstants.RDiscount)] object r)
        => Fn.Run(Category.Derivatives, () => ForwardFutures.PresentValueOfIncome(
               In.Vector("cashFlows", cashFlows), In.Vector("times", times), In.Rate("r", r)));
}
