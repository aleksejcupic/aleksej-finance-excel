using ExcelDna.Integration;
using Aleksej.Finance.Portfolio;
using Aleksej.Finance.Excel.Constants;
using Aleksej.Finance.Excel.Helpers;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>Performance attribution: TWR, Modified Dietz, IRR, NPV, and Brinson-Hood-Beebower attribution.</summary>
public static class AttributionFunctions
{
    [ExcelFunction(Name = AttributionConstants.TwrName, Category = Cat.Attribution, IsThreadSafe = true,
        Description = AttributionConstants.TwrDesc, HelpTopic = AttributionConstants.Help)]
    public static object AttrTwr(
        [ExcelArgument(Name = "subPeriodReturns", Description = AttributionConstants.SubPeriodReturns)] object subPeriodReturns)
        => Fn.Run(Category.FeesAttribution, () => PerformanceAttribution.TimeWeightedReturn(
               In.Vector("subPeriodReturns", subPeriodReturns)));

    [ExcelFunction(Name = AttributionConstants.MdietzName, Category = Cat.Attribution, IsThreadSafe = true,
        Description = AttributionConstants.MdietzDesc, HelpTopic = AttributionConstants.Help)]
    public static object AttrMdietz(
        [ExcelArgument(Name = "startValue",    Description = AttributionConstants.StartValue)]   object startValue,
        [ExcelArgument(Name = "endValue",      Description = AttributionConstants.EndValue)]     object endValue,
        [ExcelArgument(Name = "cashFlows",     Description = AttributionConstants.CashFlowsMd)]  object cashFlows,
        [ExcelArgument(Name = "cashFlowDays",  Description = AttributionConstants.CashFlowDays)] object cashFlowDays,
        [ExcelArgument(Name = "totalDays",     Description = AttributionConstants.TotalDays)]    object totalDays)
        => Fn.Run(Category.FeesAttribution, () => PerformanceAttribution.ModifiedDietz(
               In.Price("startValue", startValue), In.Price("endValue", endValue),
               In.Vector("cashFlows", cashFlows), In.Vector("cashFlowDays", cashFlowDays),
               In.PosInt("totalDays", totalDays)));

    [ExcelFunction(Name = AttributionConstants.IrrName, Category = Cat.Attribution, IsThreadSafe = true,
        Description = AttributionConstants.IrrDesc, HelpTopic = AttributionConstants.Help)]
    public static object AttrIrr(
        [ExcelArgument(Name = "cashFlows", Description = AttributionConstants.CashFlowsIrr)] object cashFlows,
        [ExcelArgument(Name = "times",     Description = AttributionConstants.TimesIrr)]     object times,
        [ExcelArgument(Name = "guess",     Description = AttributionConstants.Guess)]        object guess)
        => Fn.Run(Category.FeesAttribution, () => PerformanceAttribution.InternalRateOfReturn(
               In.Vector("cashFlows", cashFlows), In.Vector("times", times),
               In.Rate("guess", guess, AttributionConstants.DefaultGuess)));

    [ExcelFunction(Name = AttributionConstants.NpvName, Category = Cat.Attribution, IsThreadSafe = true,
        Description = AttributionConstants.NpvDesc, HelpTopic = AttributionConstants.Help)]
    public static object AttrNpv(
        [ExcelArgument(Name = "cashFlows",    Description = AttributionConstants.CashFlowsNpv)] object cashFlows,
        [ExcelArgument(Name = "times",        Description = AttributionConstants.TimesNpv)]     object times,
        [ExcelArgument(Name = "discountRate", Description = AttributionConstants.DiscountRate)] object discountRate)
        => Fn.Run(Category.FeesAttribution, () => PerformanceAttribution.NetPresentValue(
               In.Vector("cashFlows", cashFlows), In.Vector("times", times),
               In.Rate("discountRate", discountRate)));

    [ExcelFunction(Name = AttributionConstants.AllocName, Category = Cat.Attribution, IsThreadSafe = true,
        Description = AttributionConstants.AllocDesc, HelpTopic = AttributionConstants.Help)]
    public static object AttrAlloc(
        [ExcelArgument(Name = "portfolioWeight",    Description = AttributionConstants.PortfolioWeight)]    object portfolioWeight,
        [ExcelArgument(Name = "benchmarkWeight",    Description = AttributionConstants.BenchmarkWeight)]    object benchmarkWeight,
        [ExcelArgument(Name = "benchmarkSectorRet", Description = AttributionConstants.BenchmarkSectorRet)] object benchmarkSectorRet,
        [ExcelArgument(Name = "totalBenchmarkRet",  Description = AttributionConstants.TotalBenchmarkRet)]  object totalBenchmarkRet)
        => Fn.Run(Category.FeesAttribution, () => PerformanceAttribution.AllocationEffect(
               In.Num("portfolioWeight", portfolioWeight), In.Num("benchmarkWeight", benchmarkWeight),
               In.Num("benchmarkSectorRet", benchmarkSectorRet), In.Num("totalBenchmarkRet", totalBenchmarkRet)));

    [ExcelFunction(Name = AttributionConstants.SelectName, Category = Cat.Attribution, IsThreadSafe = true,
        Description = AttributionConstants.SelectDesc, HelpTopic = AttributionConstants.Help)]
    public static object AttrSelect(
        [ExcelArgument(Name = "benchmarkWeight",    Description = AttributionConstants.BenchmarkWeight)]    object benchmarkWeight,
        [ExcelArgument(Name = "portfolioSectorRet", Description = AttributionConstants.PortfolioSectorRet)] object portfolioSectorRet,
        [ExcelArgument(Name = "benchmarkSectorRet", Description = AttributionConstants.BenchmarkSectorRet)] object benchmarkSectorRet)
        => Fn.Run(Category.FeesAttribution, () => PerformanceAttribution.SelectionEffect(
               In.Num("benchmarkWeight", benchmarkWeight), In.Num("portfolioSectorRet", portfolioSectorRet),
               In.Num("benchmarkSectorRet", benchmarkSectorRet)));

    [ExcelFunction(Name = AttributionConstants.InteractName, Category = Cat.Attribution, IsThreadSafe = true,
        Description = AttributionConstants.InteractDesc, HelpTopic = AttributionConstants.Help)]
    public static object AttrInteract(
        [ExcelArgument(Name = "portfolioWeight",    Description = AttributionConstants.PortfolioWeight)]    object portfolioWeight,
        [ExcelArgument(Name = "benchmarkWeight",    Description = AttributionConstants.BenchmarkWeight)]    object benchmarkWeight,
        [ExcelArgument(Name = "portfolioSectorRet", Description = AttributionConstants.PortfolioSectorRet)] object portfolioSectorRet,
        [ExcelArgument(Name = "benchmarkSectorRet", Description = AttributionConstants.BenchmarkSectorRet)] object benchmarkSectorRet)
        => Fn.Run(Category.FeesAttribution, () => PerformanceAttribution.InteractionEffect(
               In.Num("portfolioWeight", portfolioWeight), In.Num("benchmarkWeight", benchmarkWeight),
               In.Num("portfolioSectorRet", portfolioSectorRet), In.Num("benchmarkSectorRet", benchmarkSectorRet)));

    [ExcelFunction(Name = AttributionConstants.BhbAllocName, Category = Cat.Attribution, IsThreadSafe = true,
        Description = AttributionConstants.BhbAllocDesc, HelpTopic = AttributionConstants.Help)]
    public static object AttrBhbAlloc(
        [ExcelArgument(Name = "portfolioWeights",  Description = AttributionConstants.PortfolioWeights)] object portfolioWeights,
        [ExcelArgument(Name = "benchmarkWeights",  Description = AttributionConstants.BenchmarkWeights)] object benchmarkWeights,
        [ExcelArgument(Name = "portfolioReturns",  Description = AttributionConstants.PortfolioReturns)] object portfolioReturns,
        [ExcelArgument(Name = "benchmarkReturns",  Description = AttributionConstants.BenchmarkReturns)] object benchmarkReturns)
        => Fn.Run(Category.FeesAttribution, () =>
        {
            var (alloc, _, _, _) = PerformanceAttribution.BhbAttribution(
                In.Vector("portfolioWeights", portfolioWeights), In.Vector("benchmarkWeights", benchmarkWeights),
                In.Vector("portfolioReturns", portfolioReturns), In.Vector("benchmarkReturns", benchmarkReturns));
            return alloc;
        });

    [ExcelFunction(Name = AttributionConstants.BhbSelectName, Category = Cat.Attribution, IsThreadSafe = true,
        Description = AttributionConstants.BhbSelectDesc, HelpTopic = AttributionConstants.Help)]
    public static object AttrBhbSelect(
        [ExcelArgument(Name = "portfolioWeights",  Description = AttributionConstants.PortfolioWeights)] object portfolioWeights,
        [ExcelArgument(Name = "benchmarkWeights",  Description = AttributionConstants.BenchmarkWeights)] object benchmarkWeights,
        [ExcelArgument(Name = "portfolioReturns",  Description = AttributionConstants.PortfolioReturns)] object portfolioReturns,
        [ExcelArgument(Name = "benchmarkReturns",  Description = AttributionConstants.BenchmarkReturns)] object benchmarkReturns)
        => Fn.Run(Category.FeesAttribution, () =>
        {
            var (_, select, _, _) = PerformanceAttribution.BhbAttribution(
                In.Vector("portfolioWeights", portfolioWeights), In.Vector("benchmarkWeights", benchmarkWeights),
                In.Vector("portfolioReturns", portfolioReturns), In.Vector("benchmarkReturns", benchmarkReturns));
            return select;
        });

    [ExcelFunction(Name = AttributionConstants.BhbInteractName, Category = Cat.Attribution, IsThreadSafe = true,
        Description = AttributionConstants.BhbInteractDesc, HelpTopic = AttributionConstants.Help)]
    public static object AttrBhbInteract(
        [ExcelArgument(Name = "portfolioWeights",  Description = AttributionConstants.PortfolioWeights)] object portfolioWeights,
        [ExcelArgument(Name = "benchmarkWeights",  Description = AttributionConstants.BenchmarkWeights)] object benchmarkWeights,
        [ExcelArgument(Name = "portfolioReturns",  Description = AttributionConstants.PortfolioReturns)] object portfolioReturns,
        [ExcelArgument(Name = "benchmarkReturns",  Description = AttributionConstants.BenchmarkReturns)] object benchmarkReturns)
        => Fn.Run(Category.FeesAttribution, () =>
        {
            var (_, _, interact, _) = PerformanceAttribution.BhbAttribution(
                In.Vector("portfolioWeights", portfolioWeights), In.Vector("benchmarkWeights", benchmarkWeights),
                In.Vector("portfolioReturns", portfolioReturns), In.Vector("benchmarkReturns", benchmarkReturns));
            return interact;
        });

    [ExcelFunction(Name = AttributionConstants.ActiveReturnName, Category = Cat.Attribution, IsThreadSafe = true,
        Description = AttributionConstants.ActiveReturnDesc, HelpTopic = AttributionConstants.Help)]
    public static object AttrActiveReturn(
        [ExcelArgument(Name = "portfolioReturn",  Description = AttributionConstants.PortfolioReturn)] object portfolioReturn,
        [ExcelArgument(Name = "benchmarkReturn",  Description = AttributionConstants.BenchmarkReturn)] object benchmarkReturn)
        => Fn.Run(Category.FeesAttribution, () => PerformanceAttribution.ActiveReturn(
               In.Num("portfolioReturn", portfolioReturn), In.Num("benchmarkReturn", benchmarkReturn)));
}
