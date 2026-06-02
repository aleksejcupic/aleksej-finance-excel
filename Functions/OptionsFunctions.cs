using ExcelDna.Integration;
using AleksejCupic.FinancialMath.Options;
using Aleksej.Finance.Excel.Helpers;
using Aleksej.Finance.Excel.Settings;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>Black-Scholes European option pricing, Greeks, and higher-order Greeks.</summary>
public static class OptionsFunctions
{
    private static bool Enabled => UserSettings.Load().EnableOptions;
    private static string Off   => RangeHelper.DisabledMessage("Options");

    [ExcelFunction(Name = "BS_CALL", Category = "Finance | Options", IsThreadSafe = true,
        Description = "Black-Scholes European call option price.",
        HelpTopic = "https://aleksejcupic.github.io/financial-math/options/black-scholes")]
    public static object BsCall(
        [ExcelArgument(Name = "S",     Description = "Current asset price")]         object s,
        [ExcelArgument(Name = "K",     Description = "Strike price")]                 object k,
        [ExcelArgument(Name = "T",     Description = "Time to expiry in years")]      object t,
        [ExcelArgument(Name = "r",     Description = "Continuous risk-free rate")]    object r,
        [ExcelArgument(Name = "sigma", Description = "Annualised volatility")]        object sigma)
        => Enabled ? BlackScholes.Call(RangeHelper.Scalar(s), RangeHelper.Scalar(k),
                         RangeHelper.Scalar(t), RangeHelper.Scalar(r), RangeHelper.Scalar(sigma))
                   : (object)Off;

    [ExcelFunction(Name = "BS_PUT", Category = "Finance | Options", IsThreadSafe = true,
        Description = "Black-Scholes European put option price.",
        HelpTopic = "https://aleksejcupic.github.io/financial-math/options/black-scholes")]
    public static object BsPut(
        [ExcelArgument(Name = "S",     Description = "Current asset price")]         object s,
        [ExcelArgument(Name = "K",     Description = "Strike price")]                 object k,
        [ExcelArgument(Name = "T",     Description = "Time to expiry in years")]      object t,
        [ExcelArgument(Name = "r",     Description = "Continuous risk-free rate")]    object r,
        [ExcelArgument(Name = "sigma", Description = "Annualised volatility")]        object sigma)
        => Enabled ? BlackScholes.Put(RangeHelper.Scalar(s), RangeHelper.Scalar(k),
                         RangeHelper.Scalar(t), RangeHelper.Scalar(r), RangeHelper.Scalar(sigma))
                   : (object)Off;

    [ExcelFunction(Name = "BS_DELTA", Category = "Finance | Options", IsThreadSafe = true,
        Description = "Black-Scholes Delta (dV/dS). Calls: (0,1). Puts: (-1,0).")]
    public static object BsDelta(
        [ExcelArgument(Name = "S",     Description = "Current asset price")]          object s,
        [ExcelArgument(Name = "K",     Description = "Strike price")]                  object k,
        [ExcelArgument(Name = "T",     Description = "Time to expiry in years")]       object t,
        [ExcelArgument(Name = "r",     Description = "Continuous risk-free rate")]     object r,
        [ExcelArgument(Name = "sigma", Description = "Annualised volatility")]         object sigma,
        [ExcelArgument(Name = "isPut", Description = "TRUE for put, FALSE for call")]  object isPut)
        => Enabled ? BlackScholes.Delta(RangeHelper.Scalar(s), RangeHelper.Scalar(k),
                         RangeHelper.Scalar(t), RangeHelper.Scalar(r), RangeHelper.Scalar(sigma),
                         RangeHelper.ScalarBool(isPut))
                   : (object)Off;

    [ExcelFunction(Name = "BS_GAMMA", Category = "Finance | Options", IsThreadSafe = true,
        Description = "Black-Scholes Gamma (d²V/dS²). Same for puts and calls. Always positive.")]
    public static object BsGamma(
        [ExcelArgument(Name = "S",     Description = "Current asset price")]          object s,
        [ExcelArgument(Name = "K",     Description = "Strike price")]                  object k,
        [ExcelArgument(Name = "T",     Description = "Time to expiry in years")]       object t,
        [ExcelArgument(Name = "r",     Description = "Continuous risk-free rate")]     object r,
        [ExcelArgument(Name = "sigma", Description = "Annualised volatility")]         object sigma)
        => Enabled ? BlackScholes.Gamma(RangeHelper.Scalar(s), RangeHelper.Scalar(k),
                         RangeHelper.Scalar(t), RangeHelper.Scalar(r), RangeHelper.Scalar(sigma))
                   : (object)Off;

    [ExcelFunction(Name = "BS_VEGA", Category = "Finance | Options", IsThreadSafe = true,
        Description = "Black-Scholes Vega (dV/dσ per 1% vol). Same for puts and calls.")]
    public static object BsVega(
        [ExcelArgument(Name = "S",     Description = "Current asset price")]          object s,
        [ExcelArgument(Name = "K",     Description = "Strike price")]                  object k,
        [ExcelArgument(Name = "T",     Description = "Time to expiry in years")]       object t,
        [ExcelArgument(Name = "r",     Description = "Continuous risk-free rate")]     object r,
        [ExcelArgument(Name = "sigma", Description = "Annualised volatility")]         object sigma)
        => Enabled ? BlackScholes.Vega(RangeHelper.Scalar(s), RangeHelper.Scalar(k),
                         RangeHelper.Scalar(t), RangeHelper.Scalar(r), RangeHelper.Scalar(sigma))
                   : (object)Off;

    [ExcelFunction(Name = "BS_THETA", Category = "Finance | Options", IsThreadSafe = true,
        Description = "Black-Scholes Theta — daily time decay. Typically negative.")]
    public static object BsTheta(
        [ExcelArgument(Name = "S",     Description = "Current asset price")]          object s,
        [ExcelArgument(Name = "K",     Description = "Strike price")]                  object k,
        [ExcelArgument(Name = "T",     Description = "Time to expiry in years")]       object t,
        [ExcelArgument(Name = "r",     Description = "Continuous risk-free rate")]     object r,
        [ExcelArgument(Name = "sigma", Description = "Annualised volatility")]         object sigma,
        [ExcelArgument(Name = "isPut", Description = "TRUE for put, FALSE for call")]  object isPut)
        => Enabled ? BlackScholes.Theta(RangeHelper.Scalar(s), RangeHelper.Scalar(k),
                         RangeHelper.Scalar(t), RangeHelper.Scalar(r), RangeHelper.Scalar(sigma),
                         RangeHelper.ScalarBool(isPut))
                   : (object)Off;

    [ExcelFunction(Name = "BS_RHO", Category = "Finance | Options", IsThreadSafe = true,
        Description = "Black-Scholes Rho (dV/dr per 1% rate move).")]
    public static object BsRho(
        [ExcelArgument(Name = "S",     Description = "Current asset price")]          object s,
        [ExcelArgument(Name = "K",     Description = "Strike price")]                  object k,
        [ExcelArgument(Name = "T",     Description = "Time to expiry in years")]       object t,
        [ExcelArgument(Name = "r",     Description = "Continuous risk-free rate")]     object r,
        [ExcelArgument(Name = "sigma", Description = "Annualised volatility")]         object sigma,
        [ExcelArgument(Name = "isPut", Description = "TRUE for put, FALSE for call")]  object isPut)
        => Enabled ? BlackScholes.Rho(RangeHelper.Scalar(s), RangeHelper.Scalar(k),
                         RangeHelper.Scalar(t), RangeHelper.Scalar(r), RangeHelper.Scalar(sigma),
                         RangeHelper.ScalarBool(isPut))
                   : (object)Off;

    [ExcelFunction(Name = "BS_IV", Category = "Finance | Options", IsThreadSafe = true,
        Description = "Black-Scholes implied volatility from a market price. Returns #NUM if no solution found.")]
    public static object BsIv(
        [ExcelArgument(Name = "marketPrice", Description = "Observed market option price")]    object marketPrice,
        [ExcelArgument(Name = "S",           Description = "Current asset price")]              object s,
        [ExcelArgument(Name = "K",           Description = "Strike price")]                     object k,
        [ExcelArgument(Name = "T",           Description = "Time to expiry in years")]          object t,
        [ExcelArgument(Name = "r",           Description = "Continuous risk-free rate")]        object r,
        [ExcelArgument(Name = "isPut",       Description = "TRUE for put, FALSE for call")]     object isPut)
        => Enabled ? BlackScholes.ImpliedVolatility(
                         RangeHelper.Scalar(marketPrice), RangeHelper.Scalar(s), RangeHelper.Scalar(k),
                         RangeHelper.Scalar(t), RangeHelper.Scalar(r), RangeHelper.ScalarBool(isPut))
                   : (object)Off;

    [ExcelFunction(Name = "BS_VANNA", Category = "Finance | Options", IsThreadSafe = true,
        Description = "Vanna — ∂²V/∂S∂σ. Rate of change of delta with respect to volatility.")]
    public static object BsVanna(
        [ExcelArgument(Name = "S",     Description = "Current asset price")]          object s,
        [ExcelArgument(Name = "K",     Description = "Strike price")]                  object k,
        [ExcelArgument(Name = "T",     Description = "Time to expiry in years")]       object t,
        [ExcelArgument(Name = "r",     Description = "Continuous risk-free rate")]     object r,
        [ExcelArgument(Name = "sigma", Description = "Annualised volatility")]         object sigma)
        => Enabled ? BlackScholes.Vanna(RangeHelper.Scalar(s), RangeHelper.Scalar(k),
                         RangeHelper.Scalar(t), RangeHelper.Scalar(r), RangeHelper.Scalar(sigma))
                   : (object)Off;

    [ExcelFunction(Name = "BS_CHARM", Category = "Finance | Options", IsThreadSafe = true,
        Description = "Charm — daily ∂Delta/∂t. Delta bleed per calendar day.")]
    public static object BsCharm(
        [ExcelArgument(Name = "S",     Description = "Current asset price")]          object s,
        [ExcelArgument(Name = "K",     Description = "Strike price")]                  object k,
        [ExcelArgument(Name = "T",     Description = "Time to expiry in years")]       object t,
        [ExcelArgument(Name = "r",     Description = "Continuous risk-free rate")]     object r,
        [ExcelArgument(Name = "sigma", Description = "Annualised volatility")]         object sigma)
        => Enabled ? BlackScholes.Charm(RangeHelper.Scalar(s), RangeHelper.Scalar(k),
                         RangeHelper.Scalar(t), RangeHelper.Scalar(r), RangeHelper.Scalar(sigma))
                   : (object)Off;

    [ExcelFunction(Name = "BS_VOLGA", Category = "Finance | Options", IsThreadSafe = true,
        Description = "Volga (Vomma) — ∂²V/∂σ² per 1% vol move. Convexity of price to volatility.")]
    public static object BsVolga(
        [ExcelArgument(Name = "S",     Description = "Current asset price")]          object s,
        [ExcelArgument(Name = "K",     Description = "Strike price")]                  object k,
        [ExcelArgument(Name = "T",     Description = "Time to expiry in years")]       object t,
        [ExcelArgument(Name = "r",     Description = "Continuous risk-free rate")]     object r,
        [ExcelArgument(Name = "sigma", Description = "Annualised volatility")]         object sigma)
        => Enabled ? BlackScholes.Volga(RangeHelper.Scalar(s), RangeHelper.Scalar(k),
                         RangeHelper.Scalar(t), RangeHelper.Scalar(r), RangeHelper.Scalar(sigma))
                   : (object)Off;

    [ExcelFunction(Name = "BS_SPEED", Category = "Finance | Options", IsThreadSafe = true,
        Description = "Speed — ∂Gamma/∂S. Rate of change of gamma with respect to asset price.")]
    public static object BsSpeed(
        [ExcelArgument(Name = "S",     Description = "Current asset price")]          object s,
        [ExcelArgument(Name = "K",     Description = "Strike price")]                  object k,
        [ExcelArgument(Name = "T",     Description = "Time to expiry in years")]       object t,
        [ExcelArgument(Name = "r",     Description = "Continuous risk-free rate")]     object r,
        [ExcelArgument(Name = "sigma", Description = "Annualised volatility")]         object sigma)
        => Enabled ? BlackScholes.Speed(RangeHelper.Scalar(s), RangeHelper.Scalar(k),
                         RangeHelper.Scalar(t), RangeHelper.Scalar(r), RangeHelper.Scalar(sigma))
                   : (object)Off;

    [ExcelFunction(Name = "BS_ZOMMA", Category = "Finance | Options", IsThreadSafe = true,
        Description = "Zomma — ∂Gamma/∂σ. Rate of change of gamma with respect to volatility.")]
    public static object BsZomma(
        [ExcelArgument(Name = "S",     Description = "Current asset price")]          object s,
        [ExcelArgument(Name = "K",     Description = "Strike price")]                  object k,
        [ExcelArgument(Name = "T",     Description = "Time to expiry in years")]       object t,
        [ExcelArgument(Name = "r",     Description = "Continuous risk-free rate")]     object r,
        [ExcelArgument(Name = "sigma", Description = "Annualised volatility")]         object sigma)
        => Enabled ? BlackScholes.Zomma(RangeHelper.Scalar(s), RangeHelper.Scalar(k),
                         RangeHelper.Scalar(t), RangeHelper.Scalar(r), RangeHelper.Scalar(sigma))
                   : (object)Off;
}
