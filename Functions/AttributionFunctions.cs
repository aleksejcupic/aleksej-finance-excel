using ExcelDna.Integration;
using AleksejCupic.FinancialMath.Portfolio;
using Aleksej.Finance.Excel.Helpers;
using Aleksej.Finance.Excel.Settings;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>Performance attribution: TWR, Modified Dietz, IRR, NPV, and Brinson-Hood-Beebower attribution.</summary>
public static class AttributionFunctions
{
    private static bool Enabled => UserSettings.Load().EnableFeesAttribution;
    private static string Off   => RangeHelper.DisabledMessage("Fees & Attribution");

    [ExcelFunction(Name = "ATTR_TWR", Category = "Finance | Attribution", IsThreadSafe = true,
        Description = "Time-Weighted Return: chains sub-period returns eliminating cash flow timing bias. TWR = Π(1+r_i) - 1. Industry standard for manager comparison (GIPS).",
        HelpTopic = "https://aleksejcupic.github.io/financial-math/portfolio/performance-attribution")]
    public static object AttrTwr(
        [ExcelArgument(Name = "subPeriodReturns", Description = "Simple returns for each sub-period (range, one row/column)")] object subPeriodReturns)
        => Enabled ? PerformanceAttribution.TimeWeightedReturn(RangeHelper.ToDoubleArray(subPeriodReturns))
                   : (object)Off;

    [ExcelFunction(Name = "ATTR_MDIETZ", Category = "Finance | Attribution", IsThreadSafe = true,
        Description = "Modified Dietz return. Approximates TWR when sub-period valuations are unavailable. Weights cash flows by time invested. MDietz = (EV-BV-CF) / (BV + weighted_CF).")]
    public static object AttrMdietz(
        [ExcelArgument(Name = "startValue",    Description = "Portfolio value at start of period")]                        object startValue,
        [ExcelArgument(Name = "endValue",      Description = "Portfolio value at end of period")]                          object endValue,
        [ExcelArgument(Name = "cashFlows",     Description = "External cash flows (+ = contribution, - = withdrawal)")]   object cashFlows,
        [ExcelArgument(Name = "cashFlowDays",  Description = "Day of each cash flow within the period")]                  object cashFlowDays,
        [ExcelArgument(Name = "totalDays",     Description = "Total days in the period")]                                  object totalDays)
        => Enabled ? PerformanceAttribution.ModifiedDietz(
                         RangeHelper.Scalar(startValue), RangeHelper.Scalar(endValue),
                         RangeHelper.ToDoubleArray(cashFlows), RangeHelper.ToDoubleArray(cashFlowDays),
                         RangeHelper.Scalar(totalDays))
                   : (object)Off;

    [ExcelFunction(Name = "ATTR_IRR", Category = "Finance | Attribution", IsThreadSafe = true,
        Description = "Internal Rate of Return (IRR) — the continuously compounded rate that makes NPV = 0. This is the money-weighted return (MWR). First cash flow is typically negative (investment), last is positive (exit).")]
    public static object AttrIrr(
        [ExcelArgument(Name = "cashFlows", Description = "Cash flows: negative = outflows, positive = inflows (range)")] object cashFlows,
        [ExcelArgument(Name = "times",     Description = "Time of each cash flow in years (range)")] object times,
        [ExcelArgument(Name = "guess",     Description = "Initial IRR guess (default 0.10 = 10%)")] object guess)
        => Enabled ? PerformanceAttribution.InternalRateOfReturn(
                         RangeHelper.ToDoubleArray(cashFlows), RangeHelper.ToDoubleArray(times),
                         RangeHelper.IsMissing(guess) ? 0.10 : RangeHelper.Scalar(guess))
                   : (object)Off;

    [ExcelFunction(Name = "ATTR_NPV", Category = "Finance | Attribution", IsThreadSafe = true,
        Description = "Net Present Value: sum of cash flows discounted continuously at rate r. NPV = sum(CF_i * exp(-r * t_i)).")]
    public static object AttrNpv(
        [ExcelArgument(Name = "cashFlows",    Description = "Cash flows (range)")] object cashFlows,
        [ExcelArgument(Name = "times",        Description = "Cash flow times in years (range)")] object times,
        [ExcelArgument(Name = "discountRate", Description = "Continuously compounded discount rate")] object discountRate)
        => Enabled ? PerformanceAttribution.NetPresentValue(
                         RangeHelper.ToDoubleArray(cashFlows), RangeHelper.ToDoubleArray(times),
                         RangeHelper.Scalar(discountRate))
                   : (object)Off;

    [ExcelFunction(Name = "ATTR_ALLOC", Category = "Finance | Attribution", IsThreadSafe = true,
        Description = "Brinson allocation effect for one sector: (w_p - w_b) × (r_b_sector - r_b_total). Measures value of over/underweighting the sector.")]
    public static object AttrAlloc(
        [ExcelArgument(Name = "portfolioWeight",    Description = "Portfolio weight in sector")]           object portfolioWeight,
        [ExcelArgument(Name = "benchmarkWeight",    Description = "Benchmark weight in sector")]           object benchmarkWeight,
        [ExcelArgument(Name = "benchmarkSectorRet", Description = "Benchmark return for this sector")]     object benchmarkSectorRet,
        [ExcelArgument(Name = "totalBenchmarkRet",  Description = "Total benchmark return (all sectors)")] object totalBenchmarkRet)
        => Enabled ? PerformanceAttribution.AllocationEffect(
                         RangeHelper.Scalar(portfolioWeight), RangeHelper.Scalar(benchmarkWeight),
                         RangeHelper.Scalar(benchmarkSectorRet), RangeHelper.Scalar(totalBenchmarkRet))
                   : (object)Off;

    [ExcelFunction(Name = "ATTR_SELECT", Category = "Finance | Attribution", IsThreadSafe = true,
        Description = "Brinson selection effect for one sector: w_b × (r_p_sector - r_b_sector). Measures value of security selection within the sector.")]
    public static object AttrSelect(
        [ExcelArgument(Name = "benchmarkWeight",    Description = "Benchmark weight in sector")]         object benchmarkWeight,
        [ExcelArgument(Name = "portfolioSectorRet", Description = "Portfolio return for this sector")]   object portfolioSectorRet,
        [ExcelArgument(Name = "benchmarkSectorRet", Description = "Benchmark return for this sector")]  object benchmarkSectorRet)
        => Enabled ? PerformanceAttribution.SelectionEffect(
                         RangeHelper.Scalar(benchmarkWeight), RangeHelper.Scalar(portfolioSectorRet),
                         RangeHelper.Scalar(benchmarkSectorRet))
                   : (object)Off;

    [ExcelFunction(Name = "ATTR_INTERACT", Category = "Finance | Attribution", IsThreadSafe = true,
        Description = "Brinson interaction effect for one sector: (w_p - w_b) × (r_p_sector - r_b_sector). Captures combined effect of weight and security selection.")]
    public static object AttrInteract(
        [ExcelArgument(Name = "portfolioWeight",    Description = "Portfolio weight in sector")]          object portfolioWeight,
        [ExcelArgument(Name = "benchmarkWeight",    Description = "Benchmark weight in sector")]          object benchmarkWeight,
        [ExcelArgument(Name = "portfolioSectorRet", Description = "Portfolio return for this sector")]    object portfolioSectorRet,
        [ExcelArgument(Name = "benchmarkSectorRet", Description = "Benchmark return for this sector")]   object benchmarkSectorRet)
        => Enabled ? PerformanceAttribution.InteractionEffect(
                         RangeHelper.Scalar(portfolioWeight), RangeHelper.Scalar(benchmarkWeight),
                         RangeHelper.Scalar(portfolioSectorRet), RangeHelper.Scalar(benchmarkSectorRet))
                   : (object)Off;

    [ExcelFunction(Name = "ATTR_BHB_ALLOC", Category = "Finance | Attribution", IsThreadSafe = true,
        Description = "Total Brinson-Hood-Beebower allocation effect across all sectors (sum). Enter all arrays as column ranges of equal length.")]
    public static object AttrBhbAlloc(
        [ExcelArgument(Name = "portfolioWeights",  Description = "Portfolio weights per sector (range)")] object portfolioWeights,
        [ExcelArgument(Name = "benchmarkWeights",  Description = "Benchmark weights per sector (range)")] object benchmarkWeights,
        [ExcelArgument(Name = "portfolioReturns",  Description = "Portfolio returns per sector (range)")] object portfolioReturns,
        [ExcelArgument(Name = "benchmarkReturns",  Description = "Benchmark returns per sector (range)")] object benchmarkReturns)
    {
        if (!Enabled) return Off;
        var (alloc, _, _, _) = PerformanceAttribution.BhbAttribution(
            RangeHelper.ToDoubleArray(portfolioWeights), RangeHelper.ToDoubleArray(benchmarkWeights),
            RangeHelper.ToDoubleArray(portfolioReturns), RangeHelper.ToDoubleArray(benchmarkReturns));
        return alloc;
    }

    [ExcelFunction(Name = "ATTR_BHB_SELECT", Category = "Finance | Attribution", IsThreadSafe = true,
        Description = "Total Brinson-Hood-Beebower selection effect across all sectors.")]
    public static object AttrBhbSelect(
        [ExcelArgument(Name = "portfolioWeights",  Description = "Portfolio weights per sector (range)")] object portfolioWeights,
        [ExcelArgument(Name = "benchmarkWeights",  Description = "Benchmark weights per sector (range)")] object benchmarkWeights,
        [ExcelArgument(Name = "portfolioReturns",  Description = "Portfolio returns per sector (range)")] object portfolioReturns,
        [ExcelArgument(Name = "benchmarkReturns",  Description = "Benchmark returns per sector (range)")] object benchmarkReturns)
    {
        if (!Enabled) return Off;
        var (_, select, _, _) = PerformanceAttribution.BhbAttribution(
            RangeHelper.ToDoubleArray(portfolioWeights), RangeHelper.ToDoubleArray(benchmarkWeights),
            RangeHelper.ToDoubleArray(portfolioReturns), RangeHelper.ToDoubleArray(benchmarkReturns));
        return select;
    }

    [ExcelFunction(Name = "ATTR_BHB_INTERACT", Category = "Finance | Attribution", IsThreadSafe = true,
        Description = "Total Brinson-Hood-Beebower interaction effect across all sectors.")]
    public static object AttrBhbInteract(
        [ExcelArgument(Name = "portfolioWeights",  Description = "Portfolio weights per sector (range)")] object portfolioWeights,
        [ExcelArgument(Name = "benchmarkWeights",  Description = "Benchmark weights per sector (range)")] object benchmarkWeights,
        [ExcelArgument(Name = "portfolioReturns",  Description = "Portfolio returns per sector (range)")] object portfolioReturns,
        [ExcelArgument(Name = "benchmarkReturns",  Description = "Benchmark returns per sector (range)")] object benchmarkReturns)
    {
        if (!Enabled) return Off;
        var (_, _, interact, _) = PerformanceAttribution.BhbAttribution(
            RangeHelper.ToDoubleArray(portfolioWeights), RangeHelper.ToDoubleArray(benchmarkWeights),
            RangeHelper.ToDoubleArray(portfolioReturns), RangeHelper.ToDoubleArray(benchmarkReturns));
        return interact;
    }

    [ExcelFunction(Name = "ATTR_ACTIVE_RETURN", Category = "Finance | Attribution", IsThreadSafe = true,
        Description = "Active return (portfolio return - benchmark return). Also equals the sum of all three BHB attribution effects.")]
    public static object AttrActiveReturn(
        [ExcelArgument(Name = "portfolioReturn",  Description = "Portfolio return")] object portfolioReturn,
        [ExcelArgument(Name = "benchmarkReturn",  Description = "Benchmark return")] object benchmarkReturn)
        => Enabled ? PerformanceAttribution.ActiveReturn(
                         RangeHelper.Scalar(portfolioReturn), RangeHelper.Scalar(benchmarkReturn))
                   : (object)Off;
}
