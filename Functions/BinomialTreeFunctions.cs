using ExcelDna.Integration;
using Aleksej.Finance.Options;
using Aleksej.Finance.Excel.Constants;
using Aleksej.Finance.Excel.Helpers;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>CRR binomial tree pricing for European and American options.</summary>
public static class BinomialTreeFunctions
{
    [ExcelFunction(Name = BinomialTreeConstants.PriceName, Category = Cat.Options, IsThreadSafe = true,
        Description = BinomialTreeConstants.PriceDesc, HelpTopic = BinomialTreeConstants.Help)]
    public static object BtPrice(
        [ExcelArgument(Name = "S",          Description = Arg.S)]     object s,
        [ExcelArgument(Name = "K",          Description = Arg.K)]     object k,
        [ExcelArgument(Name = "T",          Description = Arg.T)]     object t,
        [ExcelArgument(Name = "r",          Description = Arg.R)]     object r,
        [ExcelArgument(Name = "sigma",      Description = Arg.Sigma)] object sigma,
        [ExcelArgument(Name = "steps",      Description = BinomialTreeConstants.StepsDefaultDesc)] object steps,
        [ExcelArgument(Name = "isPut",      Description = BinomialTreeConstants.IsPutDefaultDesc)] object isPut,
        [ExcelArgument(Name = "isAmerican", Description = BinomialTreeConstants.IsAmericanDefaultDesc)] object isAmerican)
        => Fn.Run(Category.Options, () => BinomialTree.Price(
               In.Price("S", s), In.Price("K", k), In.Years("T", t),
               In.Rate("r", r), In.Vol("sigma", sigma),
               In.PosInt("steps", steps, BinomialTreeConstants.DefaultSteps),
               In.Flag("isPut", isPut), In.Flag("isAmerican", isAmerican)));

    [ExcelFunction(Name = BinomialTreeConstants.DeltaName, Category = Cat.Options, IsThreadSafe = true,
        Description = BinomialTreeConstants.DeltaDesc)]
    public static object BtDelta(
        [ExcelArgument(Name = "S",          Description = Arg.S)]     object s,
        [ExcelArgument(Name = "K",          Description = Arg.K)]     object k,
        [ExcelArgument(Name = "T",          Description = Arg.T)]     object t,
        [ExcelArgument(Name = "r",          Description = Arg.R)]     object r,
        [ExcelArgument(Name = "sigma",      Description = Arg.Sigma)] object sigma,
        [ExcelArgument(Name = "steps",      Description = BinomialTreeConstants.StepsDesc)]      object steps,
        [ExcelArgument(Name = "isPut",      Description = BinomialTreeConstants.IsPutDesc)]      object isPut,
        [ExcelArgument(Name = "isAmerican", Description = BinomialTreeConstants.IsAmericanDesc)] object isAmerican)
        => Fn.Run(Category.Options, () => BinomialTree.Delta(
               In.Price("S", s), In.Price("K", k), In.Years("T", t),
               In.Rate("r", r), In.Vol("sigma", sigma),
               In.PosInt("steps", steps, BinomialTreeConstants.DefaultSteps),
               In.Flag("isPut", isPut), In.Flag("isAmerican", isAmerican)));

    [ExcelFunction(Name = BinomialTreeConstants.GammaName, Category = Cat.Options, IsThreadSafe = true,
        Description = BinomialTreeConstants.GammaDesc)]
    public static object BtGamma(
        [ExcelArgument(Name = "S",          Description = Arg.S)]     object s,
        [ExcelArgument(Name = "K",          Description = Arg.K)]     object k,
        [ExcelArgument(Name = "T",          Description = Arg.T)]     object t,
        [ExcelArgument(Name = "r",          Description = Arg.R)]     object r,
        [ExcelArgument(Name = "sigma",      Description = Arg.Sigma)] object sigma,
        [ExcelArgument(Name = "steps",      Description = BinomialTreeConstants.StepsDesc)]      object steps,
        [ExcelArgument(Name = "isPut",      Description = BinomialTreeConstants.IsPutDesc)]      object isPut,
        [ExcelArgument(Name = "isAmerican", Description = BinomialTreeConstants.IsAmericanDesc)] object isAmerican)
        => Fn.Run(Category.Options, () => BinomialTree.Gamma(
               In.Price("S", s), In.Price("K", k), In.Years("T", t),
               In.Rate("r", r), In.Vol("sigma", sigma),
               In.PosInt("steps", steps, BinomialTreeConstants.DefaultSteps),
               In.Flag("isPut", isPut), In.Flag("isAmerican", isAmerican)));
}
