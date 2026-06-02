using ExcelDna.Integration;
using AleksejCupic.FinancialMath.Options;
using Aleksej.Finance.Excel.Helpers;
using Aleksej.Finance.Excel.Settings;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>Garman-Kohlhagen FX option pricing and Greeks.</summary>
public static class GarmanKohlhagenFunctions
{
    private static bool Enabled => UserSettings.Load().EnableOptions;
    private static string Off   => RangeHelper.DisabledMessage("Options");

    [ExcelFunction(Name = "GK_CALL", Category = "Finance | Options", IsThreadSafe = true,
        Description = "Garman-Kohlhagen European FX call option price. C = S*exp(-rf*T)*N(d1) - K*exp(-r*T)*N(d2).")]
    public static object GkCall(
        [ExcelArgument(Name = "S",     Description = "Spot exchange rate (domestic per foreign)")] object s,
        [ExcelArgument(Name = "K",     Description = "Strike exchange rate")]                       object k,
        [ExcelArgument(Name = "T",     Description = "Time to expiry in years")]                    object t,
        [ExcelArgument(Name = "r",     Description = "Domestic risk-free rate (continuous)")]       object r,
        [ExcelArgument(Name = "rf",    Description = "Foreign risk-free rate (continuous)")]        object rf,
        [ExcelArgument(Name = "sigma", Description = "Annualised volatility of the exchange rate")] object sigma)
        => Enabled ? GarmanKohlhagen.Call(RangeHelper.Scalar(s), RangeHelper.Scalar(k), RangeHelper.Scalar(t),
                         RangeHelper.Scalar(r), RangeHelper.Scalar(rf), RangeHelper.Scalar(sigma))
                   : (object)Off;

    [ExcelFunction(Name = "GK_PUT", Category = "Finance | Options", IsThreadSafe = true,
        Description = "Garman-Kohlhagen European FX put option price.")]
    public static object GkPut(
        [ExcelArgument(Name = "S",     Description = "Spot exchange rate")] object s,
        [ExcelArgument(Name = "K",     Description = "Strike exchange rate")] object k,
        [ExcelArgument(Name = "T",     Description = "Time to expiry in years")] object t,
        [ExcelArgument(Name = "r",     Description = "Domestic risk-free rate")] object r,
        [ExcelArgument(Name = "rf",    Description = "Foreign risk-free rate")] object rf,
        [ExcelArgument(Name = "sigma", Description = "Annualised volatility")] object sigma)
        => Enabled ? GarmanKohlhagen.Put(RangeHelper.Scalar(s), RangeHelper.Scalar(k), RangeHelper.Scalar(t),
                         RangeHelper.Scalar(r), RangeHelper.Scalar(rf), RangeHelper.Scalar(sigma))
                   : (object)Off;

    [ExcelFunction(Name = "GK_DELTA", Category = "Finance | Options", IsThreadSafe = true,
        Description = "Garman-Kohlhagen Delta (dV/dS). Call: exp(-rf*T)*N(d1). Put: exp(-rf*T)*(N(d1)-1).")]
    public static object GkDelta(
        [ExcelArgument(Name = "S",     Description = "Spot exchange rate")] object s,
        [ExcelArgument(Name = "K",     Description = "Strike exchange rate")] object k,
        [ExcelArgument(Name = "T",     Description = "Time to expiry in years")] object t,
        [ExcelArgument(Name = "r",     Description = "Domestic risk-free rate")] object r,
        [ExcelArgument(Name = "rf",    Description = "Foreign risk-free rate")] object rf,
        [ExcelArgument(Name = "sigma", Description = "Annualised volatility")] object sigma,
        [ExcelArgument(Name = "isPut", Description = "TRUE for put, FALSE for call")] object isPut)
        => Enabled ? GarmanKohlhagen.Delta(RangeHelper.Scalar(s), RangeHelper.Scalar(k), RangeHelper.Scalar(t),
                         RangeHelper.Scalar(r), RangeHelper.Scalar(rf), RangeHelper.Scalar(sigma),
                         RangeHelper.ScalarBool(isPut))
                   : (object)Off;

    [ExcelFunction(Name = "GK_GAMMA", Category = "Finance | Options", IsThreadSafe = true,
        Description = "Garman-Kohlhagen Gamma (d²V/dS²). Same for puts and calls.")]
    public static object GkGamma(
        [ExcelArgument(Name = "S",     Description = "Spot exchange rate")] object s,
        [ExcelArgument(Name = "K",     Description = "Strike exchange rate")] object k,
        [ExcelArgument(Name = "T",     Description = "Time to expiry in years")] object t,
        [ExcelArgument(Name = "r",     Description = "Domestic risk-free rate")] object r,
        [ExcelArgument(Name = "rf",    Description = "Foreign risk-free rate")] object rf,
        [ExcelArgument(Name = "sigma", Description = "Annualised volatility")] object sigma)
        => Enabled ? GarmanKohlhagen.Gamma(RangeHelper.Scalar(s), RangeHelper.Scalar(k), RangeHelper.Scalar(t),
                         RangeHelper.Scalar(r), RangeHelper.Scalar(rf), RangeHelper.Scalar(sigma))
                   : (object)Off;

    [ExcelFunction(Name = "GK_VEGA", Category = "Finance | Options", IsThreadSafe = true,
        Description = "Garman-Kohlhagen Vega (dV/dσ per 1% vol). Same for puts and calls.")]
    public static object GkVega(
        [ExcelArgument(Name = "S",     Description = "Spot exchange rate")] object s,
        [ExcelArgument(Name = "K",     Description = "Strike exchange rate")] object k,
        [ExcelArgument(Name = "T",     Description = "Time to expiry in years")] object t,
        [ExcelArgument(Name = "r",     Description = "Domestic risk-free rate")] object r,
        [ExcelArgument(Name = "rf",    Description = "Foreign risk-free rate")] object rf,
        [ExcelArgument(Name = "sigma", Description = "Annualised volatility")] object sigma)
        => Enabled ? GarmanKohlhagen.Vega(RangeHelper.Scalar(s), RangeHelper.Scalar(k), RangeHelper.Scalar(t),
                         RangeHelper.Scalar(r), RangeHelper.Scalar(rf), RangeHelper.Scalar(sigma))
                   : (object)Off;

    [ExcelFunction(Name = "GK_THETA", Category = "Finance | Options", IsThreadSafe = true,
        Description = "Garman-Kohlhagen Theta — daily time decay.")]
    public static object GkTheta(
        [ExcelArgument(Name = "S",     Description = "Spot exchange rate")] object s,
        [ExcelArgument(Name = "K",     Description = "Strike exchange rate")] object k,
        [ExcelArgument(Name = "T",     Description = "Time to expiry in years")] object t,
        [ExcelArgument(Name = "r",     Description = "Domestic risk-free rate")] object r,
        [ExcelArgument(Name = "rf",    Description = "Foreign risk-free rate")] object rf,
        [ExcelArgument(Name = "sigma", Description = "Annualised volatility")] object sigma,
        [ExcelArgument(Name = "isPut", Description = "TRUE for put, FALSE for call")] object isPut)
        => Enabled ? GarmanKohlhagen.Theta(RangeHelper.Scalar(s), RangeHelper.Scalar(k), RangeHelper.Scalar(t),
                         RangeHelper.Scalar(r), RangeHelper.Scalar(rf), RangeHelper.Scalar(sigma),
                         RangeHelper.ScalarBool(isPut))
                   : (object)Off;

    [ExcelFunction(Name = "GK_RHO", Category = "Finance | Options", IsThreadSafe = true,
        Description = "Garman-Kohlhagen Rho — sensitivity to domestic risk-free rate per 1% move.")]
    public static object GkRho(
        [ExcelArgument(Name = "S",     Description = "Spot exchange rate")] object s,
        [ExcelArgument(Name = "K",     Description = "Strike exchange rate")] object k,
        [ExcelArgument(Name = "T",     Description = "Time to expiry in years")] object t,
        [ExcelArgument(Name = "r",     Description = "Domestic risk-free rate")] object r,
        [ExcelArgument(Name = "rf",    Description = "Foreign risk-free rate")] object rf,
        [ExcelArgument(Name = "sigma", Description = "Annualised volatility")] object sigma,
        [ExcelArgument(Name = "isPut", Description = "TRUE for put, FALSE for call")] object isPut)
        => Enabled ? GarmanKohlhagen.Rho(RangeHelper.Scalar(s), RangeHelper.Scalar(k), RangeHelper.Scalar(t),
                         RangeHelper.Scalar(r), RangeHelper.Scalar(rf), RangeHelper.Scalar(sigma),
                         RangeHelper.ScalarBool(isPut))
                   : (object)Off;

    [ExcelFunction(Name = "GK_RHO_FOREIGN", Category = "Finance | Options", IsThreadSafe = true,
        Description = "Garman-Kohlhagen RhoForeign — sensitivity to foreign risk-free rate per 1% move.")]
    public static object GkRhoForeign(
        [ExcelArgument(Name = "S",     Description = "Spot exchange rate")] object s,
        [ExcelArgument(Name = "K",     Description = "Strike exchange rate")] object k,
        [ExcelArgument(Name = "T",     Description = "Time to expiry in years")] object t,
        [ExcelArgument(Name = "r",     Description = "Domestic risk-free rate")] object r,
        [ExcelArgument(Name = "rf",    Description = "Foreign risk-free rate")] object rf,
        [ExcelArgument(Name = "sigma", Description = "Annualised volatility")] object sigma,
        [ExcelArgument(Name = "isPut", Description = "TRUE for put, FALSE for call")] object isPut)
        => Enabled ? GarmanKohlhagen.RhoForeign(RangeHelper.Scalar(s), RangeHelper.Scalar(k), RangeHelper.Scalar(t),
                         RangeHelper.Scalar(r), RangeHelper.Scalar(rf), RangeHelper.Scalar(sigma),
                         RangeHelper.ScalarBool(isPut))
                   : (object)Off;

    [ExcelFunction(Name = "GK_IV", Category = "Finance | Options", IsThreadSafe = true,
        Description = "Garman-Kohlhagen implied volatility from a market price. Returns #NUM if no solution.")]
    public static object GkIv(
        [ExcelArgument(Name = "marketPrice", Description = "Observed market price")] object marketPrice,
        [ExcelArgument(Name = "S",           Description = "Spot exchange rate")]     object s,
        [ExcelArgument(Name = "K",           Description = "Strike exchange rate")]   object k,
        [ExcelArgument(Name = "T",           Description = "Time to expiry")]         object t,
        [ExcelArgument(Name = "r",           Description = "Domestic rate")]          object r,
        [ExcelArgument(Name = "rf",          Description = "Foreign rate")]           object rf,
        [ExcelArgument(Name = "isPut",       Description = "TRUE for put")]           object isPut)
        => Enabled ? GarmanKohlhagen.ImpliedVolatility(
                         RangeHelper.Scalar(marketPrice), RangeHelper.Scalar(s), RangeHelper.Scalar(k),
                         RangeHelper.Scalar(t), RangeHelper.Scalar(r), RangeHelper.Scalar(rf),
                         RangeHelper.ScalarBool(isPut))
                   : (object)Off;
}
