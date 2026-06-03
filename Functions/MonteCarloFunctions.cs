using ExcelDna.Integration;
using Aleksej.Finance.Options;
using Aleksej.Finance.Excel.Constants;
using Aleksej.Finance.Excel.Helpers;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>Monte Carlo option pricing: European via GBM, American via Longstaff-Schwartz LSM.</summary>
public static class MonteCarloFunctions
{
    [ExcelFunction(Name = MonteCarloConstants.EuropeanName, Category = Cat.Options,
        Description = MonteCarloConstants.EuropeanDesc, HelpTopic = MonteCarloConstants.Help)]
    public static object McEuropean(
        [ExcelArgument(Name = "S",     Description = Arg.S)]     object s,
        [ExcelArgument(Name = "K",     Description = Arg.K)]     object k,
        [ExcelArgument(Name = "T",     Description = Arg.T)]     object t,
        [ExcelArgument(Name = "r",     Description = Arg.R)]     object r,
        [ExcelArgument(Name = "sigma", Description = Arg.Sigma)] object sigma,
        [ExcelArgument(Name = "paths", Description = MonteCarloConstants.PathsFull)] object paths,
        [ExcelArgument(Name = "steps", Description = MonteCarloConstants.StepsFull)] object steps,
        [ExcelArgument(Name = "isPut", Description = Arg.IsPut)]  object isPut,
        [ExcelArgument(Name = "seed",  Description = Arg.Seed)]   object seed)
        => Fn.RunAsync(MonteCarloConstants.EuropeanName,
               new object[] { s, k, t, r, sigma, paths, steps, isPut, seed },
               Category.Options, () => MonteCarlo.EuropeanPrice(
                   In.Price("S", s), In.Price("K", k), In.Years("T", t),
                   In.Rate("r", r), In.Vol("sigma", sigma),
                   In.PosInt("paths", paths, MonteCarloConstants.DefaultPaths),
                   In.PosInt("steps", steps, MonteCarloConstants.DefaultSteps),
                   In.Flag("isPut", isPut),
                   In.PosInt("seed", seed, MonteCarloConstants.DefaultSeed)));

    [ExcelFunction(Name = MonteCarloConstants.AmericanName, Category = Cat.Options,
        Description = MonteCarloConstants.AmericanDesc, HelpTopic = MonteCarloConstants.Help)]
    public static object McAmerican(
        [ExcelArgument(Name = "S",     Description = Arg.S)]     object s,
        [ExcelArgument(Name = "K",     Description = Arg.K)]     object k,
        [ExcelArgument(Name = "T",     Description = Arg.T)]     object t,
        [ExcelArgument(Name = "r",     Description = Arg.R)]     object r,
        [ExcelArgument(Name = "sigma", Description = Arg.Sigma)] object sigma,
        [ExcelArgument(Name = "paths", Description = MonteCarloConstants.Paths)] object paths,
        [ExcelArgument(Name = "steps", Description = MonteCarloConstants.Steps)] object steps,
        [ExcelArgument(Name = "isPut", Description = Arg.IsPut)]  object isPut,
        [ExcelArgument(Name = "seed",  Description = MonteCarloConstants.SeedShort)] object seed)
        => Fn.RunAsync(MonteCarloConstants.AmericanName,
               new object[] { s, k, t, r, sigma, paths, steps, isPut, seed },
               Category.Options, () => MonteCarlo.AmericanPrice(
                   In.Price("S", s), In.Price("K", k), In.Years("T", t),
                   In.Rate("r", r), In.Vol("sigma", sigma),
                   In.PosInt("paths", paths, MonteCarloConstants.DefaultPaths),
                   In.PosInt("steps", steps, MonteCarloConstants.DefaultSteps),
                   In.Flag("isPut", isPut),
                   In.PosInt("seed", seed, MonteCarloConstants.DefaultSeed)));
}
