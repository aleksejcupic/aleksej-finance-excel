namespace Aleksej.Finance.Excel.Constants;

/// <summary>Names, descriptions, help topics, and argument text for the short-rate (SR_*) functions.</summary>
internal static class ShortRateConstants
{
    public const string Help = Cat.HelpBase + "/derivatives/short-rate-models";

    public const string VasicekPriceName = "SR_VASICEK_PRICE";
    public const string VasicekPriceDesc = "Zero-coupon bond price under the Vasicek model. dr = kappa*(theta-r)*dt + sigma*dW. P = A(tau)*exp(-B(tau)*r).";

    public const string VasicekYieldName = "SR_VASICEK_YIELD";
    public const string VasicekYieldDesc = "Continuously compounded zero yield under the Vasicek model. R = -ln(P)/tau.";

    public const string VasicekLrYieldName = "SR_VASICEK_LRYIELD";
    public const string VasicekLrYieldDesc = "Vasicek long-run yield as tau → infinity. R(∞) = theta - sigma²/(2*kappa²).";

    public const string VasicekOptionName = "SR_VASICEK_OPTION";
    public const string VasicekOptionDesc = "European call or put option on a zero-coupon bond under the Vasicek model (Jamshidian 1989).";

    public const string CirPriceName = "SR_CIR_PRICE";
    public const string CirPriceDesc = "Zero-coupon bond price under the CIR model. dr = kappa*(theta-r)*dt + sigma*sqrt(r)*dW. Square-root keeps r non-negative.";

    public const string CirYieldName = "SR_CIR_YIELD";
    public const string CirYieldDesc = "Continuously compounded zero yield under the CIR model. R = -ln(P)/tau.";

    public const string CirLrYieldName = "SR_CIR_LRYIELD";
    public const string CirLrYieldDesc = "CIR long-run yield. R(∞) = 2*kappa*theta / (kappa + gamma), where gamma = sqrt(kappa² + 2*sigma²).";

    // Argument descriptions (non-common)
    public const string ShortRate     = "Current short rate";
    public const string TauMaturity   = "Time to maturity in years";
    public const string Maturity      = "Maturity in years";
    public const string Kappa         = "Mean-reversion speed";
    public const string Theta         = "Long-run mean rate";
    public const string SigmaVol      = "Short-rate volatility";
    public const string SigmaCoeff    = "Volatility coefficient";
    public const string OptionExpiry  = "Option expiry in years";
    public const string BondMaturity  = "Bond maturity in years (> T)";
    public const string Strike        = "Option strike price";
    public const string IsPut         = "TRUE for put, FALSE for call";
}
