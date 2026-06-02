using ExcelDna.Integration;
using AleksejCupic.FinancialMath.Options;
using Aleksej.Finance.Excel.Helpers;
using Aleksej.Finance.Excel.Settings;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>Options on futures priced via Black's (1976) model.</summary>
public static class OptionsOnFuturesFunctions
{
    private static bool Enabled => UserSettings.Load().EnableOptions;
    private static string Off   => RangeHelper.DisabledMessage("Options");

    [ExcelFunction(Name = "OF_CALL", Category = "Finance | Options", IsThreadSafe = true,
        Description = "European call on a futures contract (Black 1976). C = exp(-rT)*[F*N(d1) - K*N(d2)]. Uses futures price F, not spot S.",
        HelpTopic = "https://aleksejcupic.github.io/financial-math/options/options-on-futures")]
    public static object OfCall(
        [ExcelArgument(Name = "F",     Description = "Current futures price")]      object f,
        [ExcelArgument(Name = "K",     Description = "Strike price")]               object k,
        [ExcelArgument(Name = "T",     Description = "Time to option expiry")]      object t,
        [ExcelArgument(Name = "r",     Description = "Continuous risk-free rate")]  object r,
        [ExcelArgument(Name = "sigma", Description = "Annualised futures price vol")] object sigma)
        => Enabled ? OptionsOnFutures.Call(RangeHelper.Scalar(f), RangeHelper.Scalar(k),
                         RangeHelper.Scalar(t), RangeHelper.Scalar(r), RangeHelper.Scalar(sigma))
                   : (object)Off;

    [ExcelFunction(Name = "OF_PUT", Category = "Finance | Options", IsThreadSafe = true,
        Description = "European put on a futures contract (Black 1976). P = exp(-rT)*[K*N(-d2) - F*N(-d1)].")]
    public static object OfPut(
        [ExcelArgument(Name = "F",     Description = "Current futures price")]      object f,
        [ExcelArgument(Name = "K",     Description = "Strike price")]               object k,
        [ExcelArgument(Name = "T",     Description = "Time to option expiry")]      object t,
        [ExcelArgument(Name = "r",     Description = "Continuous risk-free rate")]  object r,
        [ExcelArgument(Name = "sigma", Description = "Annualised futures vol")]     object sigma)
        => Enabled ? OptionsOnFutures.Put(RangeHelper.Scalar(f), RangeHelper.Scalar(k),
                         RangeHelper.Scalar(t), RangeHelper.Scalar(r), RangeHelper.Scalar(sigma))
                   : (object)Off;

    [ExcelFunction(Name = "OF_CALL_FROM_PUT", Category = "Finance | Options", IsThreadSafe = true,
        Description = "Futures call price derived from put via put-call parity. C = P + exp(-rT)*(F-K).")]
    public static object OfCallFromPut(
        [ExcelArgument(Name = "putPrice", Description = "Known put price")]        object putPrice,
        [ExcelArgument(Name = "F",        Description = "Futures price")]          object f,
        [ExcelArgument(Name = "K",        Description = "Strike")]                 object k,
        [ExcelArgument(Name = "T",        Description = "Time to expiry")]         object t,
        [ExcelArgument(Name = "r",        Description = "Risk-free rate")]         object r)
        => Enabled ? OptionsOnFutures.CallFromPut(RangeHelper.Scalar(putPrice), RangeHelper.Scalar(f),
                         RangeHelper.Scalar(k), RangeHelper.Scalar(t), RangeHelper.Scalar(r))
                   : (object)Off;

    [ExcelFunction(Name = "OF_DELTA", Category = "Finance | Options", IsThreadSafe = true,
        Description = "Futures option Delta (dV/dF). Call: exp(-rT)*N(d1). Put: exp(-rT)*(N(d1)-1).")]
    public static object OfDelta(
        [ExcelArgument(Name = "F",     Description = "Current futures price")]     object f,
        [ExcelArgument(Name = "K",     Description = "Strike")]                    object k,
        [ExcelArgument(Name = "T",     Description = "Time to expiry")]            object t,
        [ExcelArgument(Name = "r",     Description = "Risk-free rate")]            object r,
        [ExcelArgument(Name = "sigma", Description = "Annualised vol")]            object sigma,
        [ExcelArgument(Name = "isPut", Description = "TRUE for put")]              object isPut)
        => Enabled ? OptionsOnFutures.Delta(RangeHelper.Scalar(f), RangeHelper.Scalar(k),
                         RangeHelper.Scalar(t), RangeHelper.Scalar(r), RangeHelper.Scalar(sigma),
                         RangeHelper.ScalarBool(isPut))
                   : (object)Off;

    [ExcelFunction(Name = "OF_GAMMA", Category = "Finance | Options", IsThreadSafe = true,
        Description = "Futures option Gamma (d²V/dF²). Same for puts and calls.")]
    public static object OfGamma(
        [ExcelArgument(Name = "F",     Description = "Current futures price")] object f,
        [ExcelArgument(Name = "K",     Description = "Strike")] object k,
        [ExcelArgument(Name = "T",     Description = "Time to expiry")] object t,
        [ExcelArgument(Name = "r",     Description = "Risk-free rate")] object r,
        [ExcelArgument(Name = "sigma", Description = "Annualised vol")] object sigma)
        => Enabled ? OptionsOnFutures.Gamma(RangeHelper.Scalar(f), RangeHelper.Scalar(k),
                         RangeHelper.Scalar(t), RangeHelper.Scalar(r), RangeHelper.Scalar(sigma))
                   : (object)Off;

    [ExcelFunction(Name = "OF_VEGA", Category = "Finance | Options", IsThreadSafe = true,
        Description = "Futures option Vega (dV/dσ per 1% vol move). Same for puts and calls.")]
    public static object OfVega(
        [ExcelArgument(Name = "F",     Description = "Current futures price")] object f,
        [ExcelArgument(Name = "K",     Description = "Strike")] object k,
        [ExcelArgument(Name = "T",     Description = "Time to expiry")] object t,
        [ExcelArgument(Name = "r",     Description = "Risk-free rate")] object r,
        [ExcelArgument(Name = "sigma", Description = "Annualised vol")] object sigma)
        => Enabled ? OptionsOnFutures.Vega(RangeHelper.Scalar(f), RangeHelper.Scalar(k),
                         RangeHelper.Scalar(t), RangeHelper.Scalar(r), RangeHelper.Scalar(sigma))
                   : (object)Off;

    [ExcelFunction(Name = "OF_IV", Category = "Finance | Options", IsThreadSafe = true,
        Description = "Futures option implied volatility from a market price. Returns #NUM if no solution.")]
    public static object OfIv(
        [ExcelArgument(Name = "marketPrice", Description = "Observed market price")] object marketPrice,
        [ExcelArgument(Name = "F",           Description = "Futures price")]          object f,
        [ExcelArgument(Name = "K",           Description = "Strike")]                 object k,
        [ExcelArgument(Name = "T",           Description = "Time to expiry")]         object t,
        [ExcelArgument(Name = "r",           Description = "Risk-free rate")]         object r,
        [ExcelArgument(Name = "isPut",       Description = "TRUE for put")]           object isPut)
        => Enabled ? OptionsOnFutures.ImpliedVolatility(
                         RangeHelper.Scalar(marketPrice), RangeHelper.Scalar(f), RangeHelper.Scalar(k),
                         RangeHelper.Scalar(t), RangeHelper.Scalar(r), RangeHelper.ScalarBool(isPut))
                   : (object)Off;
}
