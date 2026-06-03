namespace Aleksej.Finance.Excel.Constants;

/// <summary>Names, descriptions, help topics, and argument text for the forward/futures (FWD_*) functions.</summary>
internal static class ForwardFuturesConstants
{
    public const string Help = Cat.HelpBase;

    public const string PriceName = "FWD_PRICE";
    public const string PriceDesc = "Forward price on an asset with no income. F = S * exp(r * T).";

    public const string PriceYieldName = "FWD_PRICE_YIELD";
    public const string PriceYieldDesc = "Forward price with continuous dividend yield q. F = S * exp((r - q) * T).";

    public const string PriceIncomeName = "FWD_PRICE_INCOME";
    public const string PriceIncomeDesc = "Forward price with known discrete income (PV). F = (S - I) * exp(r * T). Use FWD_PV_INCOME to compute I.";

    public const string FxName = "FWD_FX";
    public const string FxDesc = "FX forward price via covered interest rate parity. F = S * exp((r - rf) * T).";

    public const string CommodityName = "FWD_COMMODITY";
    public const string CommodityDesc = "Commodity forward price via cost-of-carry. F = S * exp((r + u - y) * T) where u = storage cost, y = convenience yield.";

    public const string ValueName = "FWD_VALUE";
    public const string ValueDesc = "Current value of an existing long forward position. f = (F - K) * exp(-r * T).";

    public const string ValueShortName = "FWD_VALUE_SHORT";
    public const string ValueShortDesc = "Current value of an existing short forward position. f = (K - F) * exp(-r * T).";

    public const string PvIncomeName = "FWD_PV_INCOME";
    public const string PvIncomeDesc = "Present value of discrete cash flows: I = sum(CF_i * exp(-r * t_i)). Use as input to FWD_PRICE_INCOME.";

    // Argument descriptions (non-common)
    public const string Spot             = "Current spot price";
    public const string RContinuous      = "Continuous risk-free rate";
    public const string TDelivery        = "Time to delivery in years";
    public const string TDeliveryShort   = "Time to delivery";
    public const string TRemaining       = "Remaining time to delivery";
    public const string DividendYield    = "Continuous dividend yield";
    public const string IncomesPv        = "Present value of income I";
    public const string FxSpot           = "Spot rate (domestic per foreign)";
    public const string RDomestic        = "Domestic risk-free rate";
    public const string RForeign         = "Foreign risk-free rate";
    public const string StorageCost      = "Annual storage cost rate";
    public const string ConvenienceYield = "Annual convenience yield";
    public const string RiskFree         = "Risk-free rate";
    public const string FairForward      = "Current fair forward price (use FWD_PRICE*)";
    public const string FairForwardShort = "Current fair forward price";
    public const string DeliveryPrice    = "Delivery price agreed at inception";
    public const string DeliveryPriceShort = "Delivery price";
    public const string CashFlows        = "Cash flow amounts (range)";
    public const string Times            = "Cash flow times in years (range)";
    public const string RDiscount        = "Continuous discount rate";
}
