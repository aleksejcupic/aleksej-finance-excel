namespace Aleksej.Finance.Excel.Constants;

/// <summary>Names, descriptions, and help topics for the Black-Scholes (BS_*) functions.</summary>
internal static class OptionsConstants
{
    public const string Help = Cat.HelpBase + "/options/black-scholes";

    public const string CallName = "BS_CALL";
    public const string CallDesc = "Black-Scholes European call option price.";

    public const string PutName = "BS_PUT";
    public const string PutDesc = "Black-Scholes European put option price.";

    public const string DeltaName = "BS_DELTA";
    public const string DeltaDesc = "Black-Scholes Delta (dV/dS). Calls: (0,1). Puts: (-1,0).";

    public const string GammaName = "BS_GAMMA";
    public const string GammaDesc = "Black-Scholes Gamma (d²V/dS²). Same for puts and calls. Always positive.";

    public const string VegaName = "BS_VEGA";
    public const string VegaDesc = "Black-Scholes Vega (dV/dσ per 1% vol). Same for puts and calls.";

    public const string ThetaName = "BS_THETA";
    public const string ThetaDesc = "Black-Scholes Theta — daily time decay. Typically negative.";

    public const string RhoName = "BS_RHO";
    public const string RhoDesc = "Black-Scholes Rho (dV/dr per 1% rate move).";

    public const string IvName = "BS_IV";
    public const string IvDesc = "Black-Scholes implied volatility from a market price. Returns #NUM if no solution found.";

    public const string VannaName = "BS_VANNA";
    public const string VannaDesc = "Vanna — ∂²V/∂S∂σ. Rate of change of delta with respect to volatility.";

    public const string CharmName = "BS_CHARM";
    public const string CharmDesc = "Charm — daily ∂Delta/∂t. Delta bleed per calendar day.";

    public const string VolgaName = "BS_VOLGA";
    public const string VolgaDesc = "Volga (Vomma) — ∂²V/∂σ² per 1% vol move. Convexity of price to volatility.";

    public const string SpeedName = "BS_SPEED";
    public const string SpeedDesc = "Speed — ∂Gamma/∂S. Rate of change of gamma with respect to asset price.";

    public const string ZommaName = "BS_ZOMMA";
    public const string ZommaDesc = "Zomma — ∂Gamma/∂σ. Rate of change of gamma with respect to volatility.";
}
