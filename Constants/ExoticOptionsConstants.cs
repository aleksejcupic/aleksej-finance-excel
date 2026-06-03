namespace Aleksej.Finance.Excel.Constants;

/// <summary>Names, descriptions, and defaults for the exotic option (EX_*) functions.</summary>
internal static class ExoticOptionsConstants
{
    public const double DefaultCashPayoff = 1.0;
    public const int    DefaultMonitoring = 252;
    public const int    DefaultPaths      = 10_000;
    public const int    DefaultSeed       = 42;

    public const string BinaryCashName = "EX_BINARY_CASH";
    public const string BinaryCashDesc = "Cash-or-nothing binary option. Pays cashPayoff if expires ITM, otherwise zero. Call: Q*exp(-rT)*N(d2). Put: Q*exp(-rT)*N(-d2).";

    public const string BinaryAssetName = "EX_BINARY_ASSET";
    public const string BinaryAssetDesc = "Asset-or-nothing binary option. Pays asset price S_T if expires ITM, otherwise zero. Call: S*N(d1). Put: S*N(-d1).";

    public const string BarrierCallName = "EX_BARRIER_CALL";
    public const string BarrierCallDesc = "European barrier call option (closed-form). Knock-out expires worthless if S touches H. Knock-in only activates if S touches H. isUp=TRUE means H is above current spot.";

    public const string BarrierPutName = "EX_BARRIER_PUT";
    public const string BarrierPutDesc = "European barrier put option (closed-form). See EX_BARRIER_CALL for parameter details.";

    public const string AsianGeoName = "EX_ASIAN_GEO";
    public const string AsianGeoDesc = "European geometric Asian option (closed-form). Payoff based on geometric average of asset price. Always below arithmetic Asian price.";

    public const string AsianArithName = "EX_ASIAN_ARITH";
    public const string AsianArithDesc = "Arithmetic Asian option price via Monte Carlo. Payoff on arithmetic average price (the market standard). Slower than EX_ASIAN_GEO.";

    public const string LookbackCallName = "EX_LOOKBACK_CALL";
    public const string LookbackCallDesc = "Floating-strike lookback call: payoff = S_T - min(S). Right to buy at the lowest price seen. sMin = current observed minimum (= S at inception).";

    public const string LookbackPutName = "EX_LOOKBACK_PUT";
    public const string LookbackPutDesc = "Floating-strike lookback put: payoff = max(S) - S_T. Right to sell at the highest price seen. sMax = current observed maximum (= S at inception).";

    // Argument descriptions (non-common or differing from Arg.*)
    public const string Time              = "Time to expiry";
    public const string RiskFree          = "Risk-free rate";
    public const string CashPayoff        = "Fixed cash amount paid if ITM (default 1.0)";
    public const string BinaryCashIsPut   = "TRUE for put (pays if S<K), FALSE for call";
    public const string IsPut             = "TRUE for put";
    public const string Barrier           = "Barrier level";
    public const string KnockInDefault    = "TRUE = knock-in, FALSE = knock-out (default FALSE)";
    public const string KnockIn           = "TRUE = knock-in";
    public const string IsUp              = "TRUE = barrier above spot (up), FALSE = below (down)";
    public const string IsUpShort         = "TRUE = up barrier";
    public const string Monitoring        = "Monitoring steps (default 252 = daily)";
    public const string Paths             = "MC paths (default 10000)";
    public const string Seed              = "Random seed (default 42)";
    public const string SMin              = "Minimum asset price observed so far (= S at inception)";
    public const string SMax              = "Maximum asset price observed so far (= S at inception)";
    public const string LookbackTime      = "Remaining time to expiry in years";
}
