namespace Aleksej.Finance.Excel.Constants;

/// <summary>Names, descriptions, help topics, defaults, and argument descriptions for the fee (FEE_*) functions.</summary>
internal static class FeeConstants
{
    public const string Help = Cat.HelpBase + "/portfolio/fees";

    // ── Literal defaults ──────────────────────────────────────────────────────
    public const int    DefaultDaysInYear   = 365;
    public const double DefaultCarryRate    = 0.20;
    public const double DefaultHurdleRate   = 0.0;
    public const double DefaultPreferredRet = 0.08;
    public const double DefaultHoldingYears = 1.0;

    // ── FEE_MGMT ──────────────────────────────────────────────────────────────
    public const string MgmtName = "FEE_MGMT";
    public const string MgmtDesc = "Management fee accrued over a period. Fee = AUM × annualRate × (days/daysInYear).";
    public const string Aum          = "Assets under management (beginning of period)";
    public const string AnnualRate   = "Annual management fee rate (e.g. 0.02 = 2%)";
    public const string DaysInPeriod = "Number of days in the accrual period";
    public const string DaysInYear   = "Day count convention (default 365)";

    // ── FEE_PERF ──────────────────────────────────────────────────────────────
    public const string PerfName = "FEE_PERF";
    public const string PerfDesc = "Performance fee with high-water mark and hurdle. Fee = max(0, (NAV - effectiveHWM) × carryRate). effectiveHWM = max(HWM, previousNAV*(1+hurdle)).";
    public const string CurrentNav    = "Current NAV per share (or total)";
    public const string HighWaterMark = "Previous peak NAV";
    public const string PreviousNav   = "NAV at start of performance period";
    public const string CarryRatePerf = "Performance fee rate (default 0.20 = 20%)";
    public const string HurdleRate    = "Minimum return before carry (default 0 = none)";

    // ── FEE_HWM ───────────────────────────────────────────────────────────────
    public const string HwmName = "FEE_HWM";
    public const string HwmDesc = "High-water mark — running maximum NAV from a series. Performance fees only charged on gains above this level.";
    public const string NavSeries = "NAV values in chronological order (range)";

    // ── FEE_EXPENSE_DRAG ──────────────────────────────────────────────────────
    public const string ExpenseDragName = "FEE_EXPENSE_DRAG";
    public const string ExpenseDragDesc = "Total wealth lost to annual expense ratio over N years. Impact = (1+gross)^n - (1+net)^n.";
    public const string GrossReturn  = "Gross annual return (e.g. 0.08 = 8%)";
    public const string ExpenseRatio = "Annual expense ratio (e.g. 0.01 = 1%)";
    public const string Years        = "Investment horizon in years";

    // ── FEE_NET_RETURN ────────────────────────────────────────────────────────
    public const string NetReturnName = "FEE_NET_RETURN";
    public const string NetReturnDesc = "Net return after an annual expense ratio. netReturn = (1+gross)/(1+expense) - 1.";
    public const string GrossReturnNet  = "Gross annual return";
    public const string ExpenseRatioNet = "Annual expense ratio";

    // ── FEE_CARRIED_INT ───────────────────────────────────────────────────────
    public const string CarriedIntName = "FEE_CARRIED_INT";
    public const string CarriedIntDesc = "PE carried interest: GP share of profits above the preferred return hurdle. Carry = carryRate × max(0, distributions - investedCapital*(1+prefReturn)^holdYears).";
    public const string TotalDistributions = "Total cash returned to LPs";
    public const string InvestedCapital     = "Total LP invested capital";
    public const string PreferredReturn     = "Hurdle rate per year (e.g. 0.08 = 8%)";
    public const string HoldingYears        = "Number of years capital was held";
    public const string CarryRateCarried    = "GP carry rate (default 0.20 = 20%)";

    // ── FEE_TRANSACTION_COST ──────────────────────────────────────────────────
    public const string TransactionCostName = "FEE_TRANSACTION_COST";
    public const string TransactionCostDesc = "Total round-trip transaction cost: commission + half the bid-ask spread. Cost = value*commission + value*spreadBps/20000.";
    public const string TradeValue     = "Notional value of the trade";
    public const string CommissionRate = "Commission as a fraction (e.g. 0.001 = 10bps)";
    public const string SpreadBps      = "Bid-ask spread in basis points (e.g. 2 = 2bps)";
}
