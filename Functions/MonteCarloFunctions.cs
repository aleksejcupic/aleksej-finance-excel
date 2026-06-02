using ExcelDna.Integration;
using AleksejCupic.FinancialMath.Options;
using Aleksej.Finance.Excel.Helpers;
using Aleksej.Finance.Excel.Settings;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>Monte Carlo option pricing: European via GBM, American via Longstaff-Schwartz LSM.</summary>
public static class MonteCarloFunctions
{
    private static bool Enabled => UserSettings.Load().EnableOptions;
    private static string Off   => RangeHelper.DisabledMessage("Options");

    [ExcelFunction(Name = "MC_EUROPEAN", Category = "Finance | Options", IsThreadSafe = true,
        Description = "European option price via Monte Carlo GBM simulation. Converges to Black-Scholes. Use BS_CALL/PUT for production; this is for validation and exotic comparisons.",
        HelpTopic = "https://aleksejcupic.github.io/financial-math/options/monte-carlo")]
    public static object McEuropean(
        [ExcelArgument(Name = "S",     Description = "Current asset price")]                          object s,
        [ExcelArgument(Name = "K",     Description = "Strike price")]                                  object k,
        [ExcelArgument(Name = "T",     Description = "Time to expiry in years")]                       object t,
        [ExcelArgument(Name = "r",     Description = "Continuous risk-free rate")]                     object r,
        [ExcelArgument(Name = "sigma", Description = "Annualised volatility")]                         object sigma,
        [ExcelArgument(Name = "paths", Description = "Number of simulated paths (default 10000, more = more accurate but slower)")] object paths,
        [ExcelArgument(Name = "steps", Description = "Time steps per path (default 50)")]              object steps,
        [ExcelArgument(Name = "isPut", Description = "TRUE for put, FALSE for call")]                  object isPut,
        [ExcelArgument(Name = "seed",  Description = "Random seed for reproducibility (default 42)")]  object seed)
        => Enabled ? MonteCarlo.EuropeanPrice(
                         RangeHelper.Scalar(s), RangeHelper.Scalar(k), RangeHelper.Scalar(t),
                         RangeHelper.Scalar(r), RangeHelper.Scalar(sigma),
                         RangeHelper.IsMissing(paths) ? 10_000 : RangeHelper.ScalarInt(paths),
                         RangeHelper.IsMissing(steps) ? 50     : RangeHelper.ScalarInt(steps),
                         RangeHelper.ScalarBool(isPut),
                         RangeHelper.IsMissing(seed)  ? 42     : RangeHelper.ScalarInt(seed))
                   : (object)Off;

    [ExcelFunction(Name = "MC_AMERICAN", Category = "Finance | Options", IsThreadSafe = true,
        Description = "American option price via Longstaff-Schwartz LSM (2001). Uses Laguerre basis regression for early exercise decisions. NOTE: slow for large path counts — use BT_PRICE for quick estimates.",
        HelpTopic = "https://aleksejcupic.github.io/financial-math/options/monte-carlo")]
    public static object McAmerican(
        [ExcelArgument(Name = "S",     Description = "Current asset price")]                          object s,
        [ExcelArgument(Name = "K",     Description = "Strike price")]                                  object k,
        [ExcelArgument(Name = "T",     Description = "Time to expiry in years")]                       object t,
        [ExcelArgument(Name = "r",     Description = "Continuous risk-free rate")]                     object r,
        [ExcelArgument(Name = "sigma", Description = "Annualised volatility")]                         object sigma,
        [ExcelArgument(Name = "paths", Description = "Number of simulated paths (default 10000)")]     object paths,
        [ExcelArgument(Name = "steps", Description = "Backward-induction time steps (default 50)")]    object steps,
        [ExcelArgument(Name = "isPut", Description = "TRUE for put, FALSE for call")]                  object isPut,
        [ExcelArgument(Name = "seed",  Description = "Random seed (default 42)")]                      object seed)
        => Enabled ? MonteCarlo.AmericanPrice(
                         RangeHelper.Scalar(s), RangeHelper.Scalar(k), RangeHelper.Scalar(t),
                         RangeHelper.Scalar(r), RangeHelper.Scalar(sigma),
                         RangeHelper.IsMissing(paths) ? 10_000 : RangeHelper.ScalarInt(paths),
                         RangeHelper.IsMissing(steps) ? 50     : RangeHelper.ScalarInt(steps),
                         RangeHelper.ScalarBool(isPut),
                         RangeHelper.IsMissing(seed)  ? 42     : RangeHelper.ScalarInt(seed))
                   : (object)Off;
}
