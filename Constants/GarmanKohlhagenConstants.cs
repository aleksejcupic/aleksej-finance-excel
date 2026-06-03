namespace Aleksej.Finance.Excel.Constants;

/// <summary>Names, descriptions, and argument text for the Garman-Kohlhagen FX (GK_*) functions.</summary>
internal static class GarmanKohlhagenConstants
{
    public const string CallName = "GK_CALL";
    public const string CallDesc = "Garman-Kohlhagen European FX call option price. C = S*exp(-rf*T)*N(d1) - K*exp(-r*T)*N(d2).";

    public const string PutName = "GK_PUT";
    public const string PutDesc = "Garman-Kohlhagen European FX put option price.";

    public const string DeltaName = "GK_DELTA";
    public const string DeltaDesc = "Garman-Kohlhagen Delta (dV/dS). Call: exp(-rf*T)*N(d1). Put: exp(-rf*T)*(N(d1)-1).";

    public const string GammaName = "GK_GAMMA";
    public const string GammaDesc = "Garman-Kohlhagen Gamma (d²V/dS²). Same for puts and calls.";

    public const string VegaName = "GK_VEGA";
    public const string VegaDesc = "Garman-Kohlhagen Vega (dV/dσ per 1% vol). Same for puts and calls.";

    public const string ThetaName = "GK_THETA";
    public const string ThetaDesc = "Garman-Kohlhagen Theta — daily time decay.";

    public const string RhoName = "GK_RHO";
    public const string RhoDesc = "Garman-Kohlhagen Rho — sensitivity to domestic risk-free rate per 1% move.";

    public const string RhoForeignName = "GK_RHO_FOREIGN";
    public const string RhoForeignDesc = "Garman-Kohlhagen RhoForeign — sensitivity to foreign risk-free rate per 1% move.";

    public const string IvName = "GK_IV";
    public const string IvDesc = "Garman-Kohlhagen implied volatility from a market price. Returns #NUM if no solution.";

    // Argument descriptions
    public const string SpotFull       = "Spot exchange rate (domestic per foreign)";
    public const string Spot           = "Spot exchange rate";
    public const string StrikeFx       = "Strike exchange rate";
    public const string DomesticFull   = "Domestic risk-free rate (continuous)";
    public const string Domestic       = "Domestic risk-free rate";
    public const string DomesticShort  = "Domestic rate";
    public const string ForeignFull    = "Foreign risk-free rate (continuous)";
    public const string Foreign        = "Foreign risk-free rate";
    public const string ForeignShort   = "Foreign rate";
    public const string SigmaFx        = "Annualised volatility of the exchange rate";
    public const string Time           = "Time to expiry";
    public const string MarketPrice    = "Observed market price";
    public const string IsPut          = "TRUE for put";
}
