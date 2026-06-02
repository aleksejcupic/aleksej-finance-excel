using ExcelDna.Integration;
using AleksejCupic.FinancialMath.Options;
using Aleksej.Finance.Excel.Helpers;
using Aleksej.Finance.Excel.Settings;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>Exotic option pricing: binary, barrier, Asian, and lookback options.</summary>
public static class ExoticOptionsFunctions
{
    private static bool Enabled => UserSettings.Load().EnableOptions;
    private static string Off   => RangeHelper.DisabledMessage("Options");

    [ExcelFunction(Name = "EX_BINARY_CASH", Category = "Finance | Options", IsThreadSafe = true,
        Description = "Cash-or-nothing binary option. Pays cashPayoff if expires ITM, otherwise zero. Call: Q*exp(-rT)*N(d2). Put: Q*exp(-rT)*N(-d2).")]
    public static object ExBinaryCash(
        [ExcelArgument(Name = "S",          Description = "Current asset price")]                        object s,
        [ExcelArgument(Name = "K",          Description = "Strike price")]                                object k,
        [ExcelArgument(Name = "T",          Description = "Time to expiry in years")]                     object t,
        [ExcelArgument(Name = "r",          Description = "Continuous risk-free rate")]                   object r,
        [ExcelArgument(Name = "sigma",      Description = "Annualised volatility")]                       object sigma,
        [ExcelArgument(Name = "cashPayoff", Description = "Fixed cash amount paid if ITM (default 1.0)")] object cashPayoff,
        [ExcelArgument(Name = "isPut",      Description = "TRUE for put (pays if S<K), FALSE for call")]  object isPut)
        => Enabled ? ExoticOptions.CashOrNothing(
                         RangeHelper.Scalar(s), RangeHelper.Scalar(k), RangeHelper.Scalar(t),
                         RangeHelper.Scalar(r), RangeHelper.Scalar(sigma),
                         RangeHelper.IsMissing(cashPayoff) ? 1.0 : RangeHelper.Scalar(cashPayoff),
                         RangeHelper.ScalarBool(isPut))
                   : (object)Off;

    [ExcelFunction(Name = "EX_BINARY_ASSET", Category = "Finance | Options", IsThreadSafe = true,
        Description = "Asset-or-nothing binary option. Pays asset price S_T if expires ITM, otherwise zero. Call: S*N(d1). Put: S*N(-d1).")]
    public static object ExBinaryAsset(
        [ExcelArgument(Name = "S",     Description = "Current asset price")]   object s,
        [ExcelArgument(Name = "K",     Description = "Strike price")]           object k,
        [ExcelArgument(Name = "T",     Description = "Time to expiry")]         object t,
        [ExcelArgument(Name = "r",     Description = "Risk-free rate")]         object r,
        [ExcelArgument(Name = "sigma", Description = "Annualised volatility")]  object sigma,
        [ExcelArgument(Name = "isPut", Description = "TRUE for put")]           object isPut)
        => Enabled ? ExoticOptions.AssetOrNothing(
                         RangeHelper.Scalar(s), RangeHelper.Scalar(k), RangeHelper.Scalar(t),
                         RangeHelper.Scalar(r), RangeHelper.Scalar(sigma), RangeHelper.ScalarBool(isPut))
                   : (object)Off;

    [ExcelFunction(Name = "EX_BARRIER_CALL", Category = "Finance | Options", IsThreadSafe = true,
        Description = "European barrier call option (closed-form). Knock-out expires worthless if S touches H. Knock-in only activates if S touches H. isUp=TRUE means H is above current spot.")]
    public static object ExBarrierCall(
        [ExcelArgument(Name = "S",       Description = "Current asset price")]                                  object s,
        [ExcelArgument(Name = "K",       Description = "Strike price")]                                          object k,
        [ExcelArgument(Name = "H",       Description = "Barrier level")]                                         object h,
        [ExcelArgument(Name = "T",       Description = "Time to expiry in years")]                               object t,
        [ExcelArgument(Name = "r",       Description = "Continuous risk-free rate")]                             object r,
        [ExcelArgument(Name = "sigma",   Description = "Annualised volatility")]                                 object sigma,
        [ExcelArgument(Name = "knockIn", Description = "TRUE = knock-in, FALSE = knock-out (default FALSE)")]   object knockIn,
        [ExcelArgument(Name = "isUp",    Description = "TRUE = barrier above spot (up), FALSE = below (down)")] object isUp)
        => Enabled ? ExoticOptions.BarrierCall(
                         RangeHelper.Scalar(s), RangeHelper.Scalar(k), RangeHelper.Scalar(h),
                         RangeHelper.Scalar(t), RangeHelper.Scalar(r), RangeHelper.Scalar(sigma),
                         RangeHelper.ScalarBool(knockIn), RangeHelper.ScalarBool(isUp))
                   : (object)Off;

    [ExcelFunction(Name = "EX_BARRIER_PUT", Category = "Finance | Options", IsThreadSafe = true,
        Description = "European barrier put option (closed-form). See EX_BARRIER_CALL for parameter details.")]
    public static object ExBarrierPut(
        [ExcelArgument(Name = "S",       Description = "Current asset price")]   object s,
        [ExcelArgument(Name = "K",       Description = "Strike price")]           object k,
        [ExcelArgument(Name = "H",       Description = "Barrier level")]          object h,
        [ExcelArgument(Name = "T",       Description = "Time to expiry")]         object t,
        [ExcelArgument(Name = "r",       Description = "Risk-free rate")]         object r,
        [ExcelArgument(Name = "sigma",   Description = "Annualised volatility")]  object sigma,
        [ExcelArgument(Name = "knockIn", Description = "TRUE = knock-in")]        object knockIn,
        [ExcelArgument(Name = "isUp",    Description = "TRUE = up barrier")]      object isUp)
        => Enabled ? ExoticOptions.BarrierPut(
                         RangeHelper.Scalar(s), RangeHelper.Scalar(k), RangeHelper.Scalar(h),
                         RangeHelper.Scalar(t), RangeHelper.Scalar(r), RangeHelper.Scalar(sigma),
                         RangeHelper.ScalarBool(knockIn), RangeHelper.ScalarBool(isUp))
                   : (object)Off;

    [ExcelFunction(Name = "EX_ASIAN_GEO", Category = "Finance | Options", IsThreadSafe = true,
        Description = "European geometric Asian option (closed-form). Payoff based on geometric average of asset price. Always below arithmetic Asian price.")]
    public static object ExAsianGeo(
        [ExcelArgument(Name = "S",     Description = "Current asset price")]  object s,
        [ExcelArgument(Name = "K",     Description = "Strike price")]          object k,
        [ExcelArgument(Name = "T",     Description = "Time to expiry")]        object t,
        [ExcelArgument(Name = "r",     Description = "Risk-free rate")]        object r,
        [ExcelArgument(Name = "sigma", Description = "Annualised volatility")] object sigma,
        [ExcelArgument(Name = "isPut", Description = "TRUE for put")]          object isPut)
        => Enabled ? ExoticOptions.GeometricAsian(
                         RangeHelper.Scalar(s), RangeHelper.Scalar(k), RangeHelper.Scalar(t),
                         RangeHelper.Scalar(r), RangeHelper.Scalar(sigma), RangeHelper.ScalarBool(isPut))
                   : (object)Off;

    [ExcelFunction(Name = "EX_ASIAN_ARITH", Category = "Finance | Options", IsThreadSafe = true,
        Description = "Arithmetic Asian option price via Monte Carlo. Payoff on arithmetic average price (the market standard). Slower than EX_ASIAN_GEO.")]
    public static object ExAsianArith(
        [ExcelArgument(Name = "S",          Description = "Current asset price")]                    object s,
        [ExcelArgument(Name = "K",          Description = "Strike price")]                            object k,
        [ExcelArgument(Name = "T",          Description = "Time to expiry")]                          object t,
        [ExcelArgument(Name = "r",          Description = "Risk-free rate")]                          object r,
        [ExcelArgument(Name = "sigma",      Description = "Annualised volatility")]                   object sigma,
        [ExcelArgument(Name = "monitoring", Description = "Monitoring steps (default 252 = daily)")]  object monitoring,
        [ExcelArgument(Name = "paths",      Description = "MC paths (default 10000)")]                object paths,
        [ExcelArgument(Name = "isPut",      Description = "TRUE for put")]                            object isPut,
        [ExcelArgument(Name = "seed",       Description = "Random seed (default 42)")]                object seed)
        => Enabled ? ExoticOptions.ArithmeticAsian(
                         RangeHelper.Scalar(s), RangeHelper.Scalar(k), RangeHelper.Scalar(t),
                         RangeHelper.Scalar(r), RangeHelper.Scalar(sigma),
                         RangeHelper.IsMissing(monitoring) ? 252 : RangeHelper.ScalarInt(monitoring),
                         RangeHelper.IsMissing(paths) ? 10_000 : RangeHelper.ScalarInt(paths),
                         RangeHelper.ScalarBool(isPut),
                         RangeHelper.IsMissing(seed) ? 42 : RangeHelper.ScalarInt(seed))
                   : (object)Off;

    [ExcelFunction(Name = "EX_LOOKBACK_CALL", Category = "Finance | Options", IsThreadSafe = true,
        Description = "Floating-strike lookback call: payoff = S_T - min(S). Right to buy at the lowest price seen. sMin = current observed minimum (= S at inception).")]
    public static object ExLookbackCall(
        [ExcelArgument(Name = "S",    Description = "Current asset price")]                                     object s,
        [ExcelArgument(Name = "sMin", Description = "Minimum asset price observed so far (= S at inception)")] object sMin,
        [ExcelArgument(Name = "T",    Description = "Remaining time to expiry in years")]                       object t,
        [ExcelArgument(Name = "r",    Description = "Continuous risk-free rate")]                               object r,
        [ExcelArgument(Name = "sigma",Description = "Annualised volatility")]                                   object sigma)
        => Enabled ? ExoticOptions.LookbackCall(
                         RangeHelper.Scalar(s), RangeHelper.Scalar(sMin), RangeHelper.Scalar(t),
                         RangeHelper.Scalar(r), RangeHelper.Scalar(sigma))
                   : (object)Off;

    [ExcelFunction(Name = "EX_LOOKBACK_PUT", Category = "Finance | Options", IsThreadSafe = true,
        Description = "Floating-strike lookback put: payoff = max(S) - S_T. Right to sell at the highest price seen. sMax = current observed maximum (= S at inception).")]
    public static object ExLookbackPut(
        [ExcelArgument(Name = "S",    Description = "Current asset price")]                                     object s,
        [ExcelArgument(Name = "sMax", Description = "Maximum asset price observed so far (= S at inception)")] object sMax,
        [ExcelArgument(Name = "T",    Description = "Remaining time to expiry in years")]                       object t,
        [ExcelArgument(Name = "r",    Description = "Continuous risk-free rate")]                               object r,
        [ExcelArgument(Name = "sigma",Description = "Annualised volatility")]                                   object sigma)
        => Enabled ? ExoticOptions.LookbackPut(
                         RangeHelper.Scalar(s), RangeHelper.Scalar(sMax), RangeHelper.Scalar(t),
                         RangeHelper.Scalar(r), RangeHelper.Scalar(sigma))
                   : (object)Off;
}
