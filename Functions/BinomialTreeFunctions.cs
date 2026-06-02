using ExcelDna.Integration;
using AleksejCupic.FinancialMath.Options;
using Aleksej.Finance.Excel.Helpers;
using Aleksej.Finance.Excel.Settings;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>CRR binomial tree pricing for European and American options.</summary>
public static class BinomialTreeFunctions
{
    private static bool Enabled => UserSettings.Load().EnableOptions;
    private static string Off   => RangeHelper.DisabledMessage("Options");

    [ExcelFunction(Name = "BT_PRICE", Category = "Finance | Options", IsThreadSafe = true,
        Description = "CRR binomial tree option price. Supports both European and American exercise. Converges to Black-Scholes for European options.",
        HelpTopic = "https://aleksejcupic.github.io/financial-math/options/binomial-tree")]
    public static object BtPrice(
        [ExcelArgument(Name = "S",          Description = "Current asset price")]                          object s,
        [ExcelArgument(Name = "K",          Description = "Strike price")]                                  object k,
        [ExcelArgument(Name = "T",          Description = "Time to expiry in years")]                       object t,
        [ExcelArgument(Name = "r",          Description = "Continuous risk-free rate")]                     object r,
        [ExcelArgument(Name = "sigma",      Description = "Annualised volatility")]                         object sigma,
        [ExcelArgument(Name = "steps",      Description = "Number of time steps (default 200, higher = more accurate)")] object steps,
        [ExcelArgument(Name = "isPut",      Description = "TRUE for put, FALSE for call (default FALSE)")]  object isPut,
        [ExcelArgument(Name = "isAmerican", Description = "TRUE for American (early exercise), FALSE for European (default FALSE)")] object isAmerican)
        => Enabled ? BinomialTree.Price(
                         RangeHelper.Scalar(s), RangeHelper.Scalar(k), RangeHelper.Scalar(t),
                         RangeHelper.Scalar(r), RangeHelper.Scalar(sigma),
                         RangeHelper.IsMissing(steps) ? 200 : RangeHelper.ScalarInt(steps),
                         RangeHelper.ScalarBool(isPut), RangeHelper.ScalarBool(isAmerican))
                   : (object)Off;

    [ExcelFunction(Name = "BT_DELTA", Category = "Finance | Options", IsThreadSafe = true,
        Description = "Binomial tree Delta — extracted from level-1 nodes. More accurate for American options than BS_DELTA.")]
    public static object BtDelta(
        [ExcelArgument(Name = "S",          Description = "Current asset price")]       object s,
        [ExcelArgument(Name = "K",          Description = "Strike price")]               object k,
        [ExcelArgument(Name = "T",          Description = "Time to expiry in years")]    object t,
        [ExcelArgument(Name = "r",          Description = "Continuous risk-free rate")]  object r,
        [ExcelArgument(Name = "sigma",      Description = "Annualised volatility")]      object sigma,
        [ExcelArgument(Name = "steps",      Description = "Number of time steps")]       object steps,
        [ExcelArgument(Name = "isPut",      Description = "TRUE for put")]               object isPut,
        [ExcelArgument(Name = "isAmerican", Description = "TRUE for American")]          object isAmerican)
        => Enabled ? BinomialTree.Delta(
                         RangeHelper.Scalar(s), RangeHelper.Scalar(k), RangeHelper.Scalar(t),
                         RangeHelper.Scalar(r), RangeHelper.Scalar(sigma),
                         RangeHelper.IsMissing(steps) ? 200 : RangeHelper.ScalarInt(steps),
                         RangeHelper.ScalarBool(isPut), RangeHelper.ScalarBool(isAmerican))
                   : (object)Off;

    [ExcelFunction(Name = "BT_GAMMA", Category = "Finance | Options", IsThreadSafe = true,
        Description = "Binomial tree Gamma — second-order finite difference from level-1 nodes.")]
    public static object BtGamma(
        [ExcelArgument(Name = "S",          Description = "Current asset price")]       object s,
        [ExcelArgument(Name = "K",          Description = "Strike price")]               object k,
        [ExcelArgument(Name = "T",          Description = "Time to expiry in years")]    object t,
        [ExcelArgument(Name = "r",          Description = "Continuous risk-free rate")]  object r,
        [ExcelArgument(Name = "sigma",      Description = "Annualised volatility")]      object sigma,
        [ExcelArgument(Name = "steps",      Description = "Number of time steps")]       object steps,
        [ExcelArgument(Name = "isPut",      Description = "TRUE for put")]               object isPut,
        [ExcelArgument(Name = "isAmerican", Description = "TRUE for American")]          object isAmerican)
        => Enabled ? BinomialTree.Gamma(
                         RangeHelper.Scalar(s), RangeHelper.Scalar(k), RangeHelper.Scalar(t),
                         RangeHelper.Scalar(r), RangeHelper.Scalar(sigma),
                         RangeHelper.IsMissing(steps) ? 200 : RangeHelper.ScalarInt(steps),
                         RangeHelper.ScalarBool(isPut), RangeHelper.ScalarBool(isAmerican))
                   : (object)Off;
}
