using ExcelDna.Integration;
using Aleksej.Finance.Options;
using Aleksej.Finance.Excel.Constants;
using Aleksej.Finance.Excel.Helpers;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>Options on futures priced via Black's (1976) model.</summary>
public static class OptionsOnFuturesFunctions
{
    [ExcelFunction(Name = OptionsOnFuturesConstants.CallName, Category = Cat.Options, IsThreadSafe = true,
        Description = OptionsOnFuturesConstants.CallDesc, HelpTopic = OptionsOnFuturesConstants.Help)]
    public static object OfCall(
        [ExcelArgument(Name = "F",     Description = OptionsOnFuturesConstants.Futures)]      object f,
        [ExcelArgument(Name = "K",     Description = Arg.K)]                                  object k,
        [ExcelArgument(Name = "T",     Description = OptionsOnFuturesConstants.OptionExpiry)] object t,
        [ExcelArgument(Name = "r",     Description = Arg.R)]                                  object r,
        [ExcelArgument(Name = "sigma", Description = OptionsOnFuturesConstants.FuturesVol)]   object sigma)
        => Fn.Run(Category.Options, () => OptionsOnFutures.Call(
               In.Price("F", f), In.Price("K", k), In.Years("T", t),
               In.Rate("r", r), In.Vol("sigma", sigma)));

    [ExcelFunction(Name = OptionsOnFuturesConstants.PutName, Category = Cat.Options, IsThreadSafe = true,
        Description = OptionsOnFuturesConstants.PutDesc)]
    public static object OfPut(
        [ExcelArgument(Name = "F",     Description = OptionsOnFuturesConstants.Futures)]      object f,
        [ExcelArgument(Name = "K",     Description = Arg.K)]                                  object k,
        [ExcelArgument(Name = "T",     Description = OptionsOnFuturesConstants.OptionExpiry)] object t,
        [ExcelArgument(Name = "r",     Description = Arg.R)]                                  object r,
        [ExcelArgument(Name = "sigma", Description = OptionsOnFuturesConstants.FuturesVolShort)] object sigma)
        => Fn.Run(Category.Options, () => OptionsOnFutures.Put(
               In.Price("F", f), In.Price("K", k), In.Years("T", t),
               In.Rate("r", r), In.Vol("sigma", sigma)));

    [ExcelFunction(Name = OptionsOnFuturesConstants.CallFromPutName, Category = Cat.Options, IsThreadSafe = true,
        Description = OptionsOnFuturesConstants.CallFromPutDesc)]
    public static object OfCallFromPut(
        [ExcelArgument(Name = "putPrice", Description = OptionsOnFuturesConstants.PutPrice)]     object putPrice,
        [ExcelArgument(Name = "F",        Description = OptionsOnFuturesConstants.FuturesShort)] object f,
        [ExcelArgument(Name = "K",        Description = OptionsOnFuturesConstants.Strike)]       object k,
        [ExcelArgument(Name = "T",        Description = OptionsOnFuturesConstants.Time)]         object t,
        [ExcelArgument(Name = "r",        Description = OptionsOnFuturesConstants.RiskFree)]     object r)
        => Fn.Run(Category.Options, () => OptionsOnFutures.CallFromPut(
               In.Price("putPrice", putPrice), In.Price("F", f),
               In.Price("K", k), In.Years("T", t), In.Rate("r", r)));

    [ExcelFunction(Name = OptionsOnFuturesConstants.DeltaName, Category = Cat.Options, IsThreadSafe = true,
        Description = OptionsOnFuturesConstants.DeltaDesc)]
    public static object OfDelta(
        [ExcelArgument(Name = "F",     Description = OptionsOnFuturesConstants.Futures)]  object f,
        [ExcelArgument(Name = "K",     Description = OptionsOnFuturesConstants.Strike)]   object k,
        [ExcelArgument(Name = "T",     Description = OptionsOnFuturesConstants.Time)]     object t,
        [ExcelArgument(Name = "r",     Description = OptionsOnFuturesConstants.RiskFree)] object r,
        [ExcelArgument(Name = "sigma", Description = OptionsOnFuturesConstants.Vol)]      object sigma,
        [ExcelArgument(Name = "isPut", Description = OptionsOnFuturesConstants.IsPut)]    object isPut)
        => Fn.Run(Category.Options, () => OptionsOnFutures.Delta(
               In.Price("F", f), In.Price("K", k), In.Years("T", t),
               In.Rate("r", r), In.Vol("sigma", sigma), In.Flag("isPut", isPut)));

    [ExcelFunction(Name = OptionsOnFuturesConstants.GammaName, Category = Cat.Options, IsThreadSafe = true,
        Description = OptionsOnFuturesConstants.GammaDesc)]
    public static object OfGamma(
        [ExcelArgument(Name = "F",     Description = OptionsOnFuturesConstants.Futures)]  object f,
        [ExcelArgument(Name = "K",     Description = OptionsOnFuturesConstants.Strike)]   object k,
        [ExcelArgument(Name = "T",     Description = OptionsOnFuturesConstants.Time)]     object t,
        [ExcelArgument(Name = "r",     Description = OptionsOnFuturesConstants.RiskFree)] object r,
        [ExcelArgument(Name = "sigma", Description = OptionsOnFuturesConstants.Vol)]      object sigma)
        => Fn.Run(Category.Options, () => OptionsOnFutures.Gamma(
               In.Price("F", f), In.Price("K", k), In.Years("T", t),
               In.Rate("r", r), In.Vol("sigma", sigma)));

    [ExcelFunction(Name = OptionsOnFuturesConstants.VegaName, Category = Cat.Options, IsThreadSafe = true,
        Description = OptionsOnFuturesConstants.VegaDesc)]
    public static object OfVega(
        [ExcelArgument(Name = "F",     Description = OptionsOnFuturesConstants.Futures)]  object f,
        [ExcelArgument(Name = "K",     Description = OptionsOnFuturesConstants.Strike)]   object k,
        [ExcelArgument(Name = "T",     Description = OptionsOnFuturesConstants.Time)]     object t,
        [ExcelArgument(Name = "r",     Description = OptionsOnFuturesConstants.RiskFree)] object r,
        [ExcelArgument(Name = "sigma", Description = OptionsOnFuturesConstants.Vol)]      object sigma)
        => Fn.Run(Category.Options, () => OptionsOnFutures.Vega(
               In.Price("F", f), In.Price("K", k), In.Years("T", t),
               In.Rate("r", r), In.Vol("sigma", sigma)));

    [ExcelFunction(Name = OptionsOnFuturesConstants.IvName, Category = Cat.Options, IsThreadSafe = true,
        Description = OptionsOnFuturesConstants.IvDesc)]
    public static object OfIv(
        [ExcelArgument(Name = "marketPrice", Description = OptionsOnFuturesConstants.MarketPrice)]  object marketPrice,
        [ExcelArgument(Name = "F",           Description = OptionsOnFuturesConstants.FuturesShort)] object f,
        [ExcelArgument(Name = "K",           Description = OptionsOnFuturesConstants.Strike)]       object k,
        [ExcelArgument(Name = "T",           Description = OptionsOnFuturesConstants.Time)]         object t,
        [ExcelArgument(Name = "r",           Description = OptionsOnFuturesConstants.RiskFree)]     object r,
        [ExcelArgument(Name = "isPut",       Description = OptionsOnFuturesConstants.IsPut)]        object isPut)
        => Fn.Run(Category.Options, () => OptionsOnFutures.ImpliedVolatility(
               In.Price("marketPrice", marketPrice), In.Price("F", f), In.Price("K", k),
               In.Years("T", t), In.Rate("r", r), In.Flag("isPut", isPut)));
}
