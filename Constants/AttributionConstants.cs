namespace Aleksej.Finance.Excel.Constants;

/// <summary>Names, descriptions, help topics, defaults, and argument descriptions for the attribution (ATTR_*) functions.</summary>
internal static class AttributionConstants
{
    public const string Help = Cat.HelpBase + "/portfolio/performance-attribution";

    // ── Literal defaults ──────────────────────────────────────────────────────
    public const double DefaultGuess = 0.10;

    // ── Shared argument descriptions ──────────────────────────────────────────
    public const string PortfolioWeights = "Portfolio weights per sector (range)";
    public const string BenchmarkWeights = "Benchmark weights per sector (range)";
    public const string PortfolioReturns = "Portfolio returns per sector (range)";
    public const string BenchmarkReturns = "Benchmark returns per sector (range)";
    public const string BenchmarkSectorRet = "Benchmark return for this sector";
    public const string PortfolioSectorRet = "Portfolio return for this sector";
    public const string PortfolioWeight    = "Portfolio weight in sector";
    public const string BenchmarkWeight    = "Benchmark weight in sector";

    // ── ATTR_TWR ──────────────────────────────────────────────────────────────
    public const string TwrName = "ATTR_TWR";
    public const string TwrDesc = "Time-Weighted Return: chains sub-period returns eliminating cash flow timing bias. TWR = Π(1+r_i) - 1. Industry standard for manager comparison (GIPS).";
    public const string SubPeriodReturns = "Simple returns for each sub-period (range, one row/column)";

    // ── ATTR_MDIETZ ───────────────────────────────────────────────────────────
    public const string MdietzName = "ATTR_MDIETZ";
    public const string MdietzDesc = "Modified Dietz return. Approximates TWR when sub-period valuations are unavailable. Weights cash flows by time invested. MDietz = (EV-BV-CF) / (BV + weighted_CF).";
    public const string StartValue   = "Portfolio value at start of period";
    public const string EndValue     = "Portfolio value at end of period";
    public const string CashFlowsMd  = "External cash flows (+ = contribution, - = withdrawal)";
    public const string CashFlowDays = "Day of each cash flow within the period";
    public const string TotalDays    = "Total days in the period";

    // ── ATTR_IRR ──────────────────────────────────────────────────────────────
    public const string IrrName = "ATTR_IRR";
    public const string IrrDesc = "Internal Rate of Return (IRR) — the continuously compounded rate that makes NPV = 0. This is the money-weighted return (MWR). First cash flow is typically negative (investment), last is positive (exit).";
    public const string CashFlowsIrr = "Cash flows: negative = outflows, positive = inflows (range)";
    public const string TimesIrr     = "Time of each cash flow in years (range)";
    public const string Guess        = "Initial IRR guess (default 0.10 = 10%)";

    // ── ATTR_NPV ──────────────────────────────────────────────────────────────
    public const string NpvName = "ATTR_NPV";
    public const string NpvDesc = "Net Present Value: sum of cash flows discounted continuously at rate r. NPV = sum(CF_i * exp(-r * t_i)).";
    public const string CashFlowsNpv = "Cash flows (range)";
    public const string TimesNpv     = "Cash flow times in years (range)";
    public const string DiscountRate = "Continuously compounded discount rate";

    // ── ATTR_ALLOC ────────────────────────────────────────────────────────────
    public const string AllocName = "ATTR_ALLOC";
    public const string AllocDesc = "Brinson allocation effect for one sector: (w_p - w_b) × (r_b_sector - r_b_total). Measures value of over/underweighting the sector.";
    public const string TotalBenchmarkRet = "Total benchmark return (all sectors)";

    // ── ATTR_SELECT ───────────────────────────────────────────────────────────
    public const string SelectName = "ATTR_SELECT";
    public const string SelectDesc = "Brinson selection effect for one sector: w_b × (r_p_sector - r_b_sector). Measures value of security selection within the sector.";

    // ── ATTR_INTERACT ─────────────────────────────────────────────────────────
    public const string InteractName = "ATTR_INTERACT";
    public const string InteractDesc = "Brinson interaction effect for one sector: (w_p - w_b) × (r_p_sector - r_b_sector). Captures combined effect of weight and security selection.";

    // ── ATTR_BHB_ALLOC ────────────────────────────────────────────────────────
    public const string BhbAllocName = "ATTR_BHB_ALLOC";
    public const string BhbAllocDesc = "Total Brinson-Hood-Beebower allocation effect across all sectors (sum). Enter all arrays as column ranges of equal length.";

    // ── ATTR_BHB_SELECT ───────────────────────────────────────────────────────
    public const string BhbSelectName = "ATTR_BHB_SELECT";
    public const string BhbSelectDesc = "Total Brinson-Hood-Beebower selection effect across all sectors.";

    // ── ATTR_BHB_INTERACT ─────────────────────────────────────────────────────
    public const string BhbInteractName = "ATTR_BHB_INTERACT";
    public const string BhbInteractDesc = "Total Brinson-Hood-Beebower interaction effect across all sectors.";

    // ── ATTR_ACTIVE_RETURN ────────────────────────────────────────────────────
    public const string ActiveReturnName = "ATTR_ACTIVE_RETURN";
    public const string ActiveReturnDesc = "Active return (portfolio return - benchmark return). Also equals the sum of all three BHB attribution effects.";
    public const string PortfolioReturn = "Portfolio return";
    public const string BenchmarkReturn = "Benchmark return";
}
