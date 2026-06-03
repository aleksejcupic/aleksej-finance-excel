namespace Aleksej.Finance.Excel.Constants;

/// <summary>Names, descriptions, help, and argument text for the options-on-futures (OF_*) functions.</summary>
internal static class OptionsOnFuturesConstants
{
    public const string Help = Cat.HelpBase + "/options/options-on-futures";

    public const string CallName = "OF_CALL";
    public const string CallDesc = "European call on a futures contract (Black 1976). C = exp(-rT)*[F*N(d1) - K*N(d2)]. Uses futures price F, not spot S.";

    public const string PutName = "OF_PUT";
    public const string PutDesc = "European put on a futures contract (Black 1976). P = exp(-rT)*[K*N(-d2) - F*N(-d1)].";

    public const string CallFromPutName = "OF_CALL_FROM_PUT";
    public const string CallFromPutDesc = "Futures call price derived from put via put-call parity. C = P + exp(-rT)*(F-K).";

    public const string DeltaName = "OF_DELTA";
    public const string DeltaDesc = "Futures option Delta (dV/dF). Call: exp(-rT)*N(d1). Put: exp(-rT)*(N(d1)-1).";

    public const string GammaName = "OF_GAMMA";
    public const string GammaDesc = "Futures option Gamma (d²V/dF²). Same for puts and calls.";

    public const string VegaName = "OF_VEGA";
    public const string VegaDesc = "Futures option Vega (dV/dσ per 1% vol move). Same for puts and calls.";

    public const string IvName = "OF_IV";
    public const string IvDesc = "Futures option implied volatility from a market price. Returns #NUM if no solution.";

    // Argument descriptions
    public const string Futures        = "Current futures price";
    public const string FuturesShort   = "Futures price";
    public const string OptionExpiry   = "Time to option expiry";
    public const string FuturesVol     = "Annualised futures price vol";
    public const string FuturesVolShort = "Annualised futures vol";
    public const string Strike         = "Strike";
    public const string Time           = "Time to expiry";
    public const string RiskFree       = "Risk-free rate";
    public const string Vol            = "Annualised vol";
    public const string IsPut          = "TRUE for put";
    public const string PutPrice       = "Known put price";
    public const string MarketPrice    = "Observed market price";
}
