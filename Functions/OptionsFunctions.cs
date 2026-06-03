using ExcelDna.Integration;
using Aleksej.Finance.Options;
using Aleksej.Finance.Excel.Constants;
using Aleksej.Finance.Excel.Helpers;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>Black-Scholes European option pricing, Greeks, and higher-order Greeks.</summary>
public static class OptionsFunctions
{
    [ExcelFunction(Name = OptionsConstants.CallName, Category = Cat.Options, IsThreadSafe = true,
        Description = OptionsConstants.CallDesc, HelpTopic = OptionsConstants.Help)]
    public static object BsCall(
        [ExcelArgument(Name = "S",     Description = Arg.S)]     object s,
        [ExcelArgument(Name = "K",     Description = Arg.K)]     object k,
        [ExcelArgument(Name = "T",     Description = Arg.T)]     object t,
        [ExcelArgument(Name = "r",     Description = Arg.R)]     object r,
        [ExcelArgument(Name = "sigma", Description = Arg.Sigma)] object sigma)
        => Fn.Run(Category.Options, () => BlackScholes.Call(
               In.Price("S", s), In.Price("K", k), In.Years("T", t),
               In.Rate("r", r), In.Vol("sigma", sigma)));

    [ExcelFunction(Name = OptionsConstants.PutName, Category = Cat.Options, IsThreadSafe = true,
        Description = OptionsConstants.PutDesc, HelpTopic = OptionsConstants.Help)]
    public static object BsPut(
        [ExcelArgument(Name = "S",     Description = Arg.S)]     object s,
        [ExcelArgument(Name = "K",     Description = Arg.K)]     object k,
        [ExcelArgument(Name = "T",     Description = Arg.T)]     object t,
        [ExcelArgument(Name = "r",     Description = Arg.R)]     object r,
        [ExcelArgument(Name = "sigma", Description = Arg.Sigma)] object sigma)
        => Fn.Run(Category.Options, () => BlackScholes.Put(
               In.Price("S", s), In.Price("K", k), In.Years("T", t),
               In.Rate("r", r), In.Vol("sigma", sigma)));

    [ExcelFunction(Name = OptionsConstants.DeltaName, Category = Cat.Options, IsThreadSafe = true,
        Description = OptionsConstants.DeltaDesc)]
    public static object BsDelta(
        [ExcelArgument(Name = "S",     Description = Arg.S)]     object s,
        [ExcelArgument(Name = "K",     Description = Arg.K)]     object k,
        [ExcelArgument(Name = "T",     Description = Arg.T)]     object t,
        [ExcelArgument(Name = "r",     Description = Arg.R)]     object r,
        [ExcelArgument(Name = "sigma", Description = Arg.Sigma)] object sigma,
        [ExcelArgument(Name = "isPut", Description = Arg.IsPut)] object isPut)
        => Fn.Run(Category.Options, () => BlackScholes.Delta(
               In.Price("S", s), In.Price("K", k), In.Years("T", t),
               In.Rate("r", r), In.Vol("sigma", sigma), In.Flag("isPut", isPut)));

    [ExcelFunction(Name = OptionsConstants.GammaName, Category = Cat.Options, IsThreadSafe = true,
        Description = OptionsConstants.GammaDesc)]
    public static object BsGamma(
        [ExcelArgument(Name = "S",     Description = Arg.S)]     object s,
        [ExcelArgument(Name = "K",     Description = Arg.K)]     object k,
        [ExcelArgument(Name = "T",     Description = Arg.T)]     object t,
        [ExcelArgument(Name = "r",     Description = Arg.R)]     object r,
        [ExcelArgument(Name = "sigma", Description = Arg.Sigma)] object sigma)
        => Fn.Run(Category.Options, () => BlackScholes.Gamma(
               In.Price("S", s), In.Price("K", k), In.Years("T", t),
               In.Rate("r", r), In.Vol("sigma", sigma)));

    [ExcelFunction(Name = OptionsConstants.VegaName, Category = Cat.Options, IsThreadSafe = true,
        Description = OptionsConstants.VegaDesc)]
    public static object BsVega(
        [ExcelArgument(Name = "S",     Description = Arg.S)]     object s,
        [ExcelArgument(Name = "K",     Description = Arg.K)]     object k,
        [ExcelArgument(Name = "T",     Description = Arg.T)]     object t,
        [ExcelArgument(Name = "r",     Description = Arg.R)]     object r,
        [ExcelArgument(Name = "sigma", Description = Arg.Sigma)] object sigma)
        => Fn.Run(Category.Options, () => BlackScholes.Vega(
               In.Price("S", s), In.Price("K", k), In.Years("T", t),
               In.Rate("r", r), In.Vol("sigma", sigma)));

    [ExcelFunction(Name = OptionsConstants.ThetaName, Category = Cat.Options, IsThreadSafe = true,
        Description = OptionsConstants.ThetaDesc)]
    public static object BsTheta(
        [ExcelArgument(Name = "S",     Description = Arg.S)]     object s,
        [ExcelArgument(Name = "K",     Description = Arg.K)]     object k,
        [ExcelArgument(Name = "T",     Description = Arg.T)]     object t,
        [ExcelArgument(Name = "r",     Description = Arg.R)]     object r,
        [ExcelArgument(Name = "sigma", Description = Arg.Sigma)] object sigma,
        [ExcelArgument(Name = "isPut", Description = Arg.IsPut)] object isPut)
        => Fn.Run(Category.Options, () => BlackScholes.Theta(
               In.Price("S", s), In.Price("K", k), In.Years("T", t),
               In.Rate("r", r), In.Vol("sigma", sigma), In.Flag("isPut", isPut)));

    [ExcelFunction(Name = OptionsConstants.RhoName, Category = Cat.Options, IsThreadSafe = true,
        Description = OptionsConstants.RhoDesc)]
    public static object BsRho(
        [ExcelArgument(Name = "S",     Description = Arg.S)]     object s,
        [ExcelArgument(Name = "K",     Description = Arg.K)]     object k,
        [ExcelArgument(Name = "T",     Description = Arg.T)]     object t,
        [ExcelArgument(Name = "r",     Description = Arg.R)]     object r,
        [ExcelArgument(Name = "sigma", Description = Arg.Sigma)] object sigma,
        [ExcelArgument(Name = "isPut", Description = Arg.IsPut)] object isPut)
        => Fn.Run(Category.Options, () => BlackScholes.Rho(
               In.Price("S", s), In.Price("K", k), In.Years("T", t),
               In.Rate("r", r), In.Vol("sigma", sigma), In.Flag("isPut", isPut)));

    [ExcelFunction(Name = OptionsConstants.IvName, Category = Cat.Options, IsThreadSafe = true,
        Description = OptionsConstants.IvDesc)]
    public static object BsIv(
        [ExcelArgument(Name = "marketPrice", Description = Arg.MarketPrice)] object marketPrice,
        [ExcelArgument(Name = "S",           Description = Arg.S)]           object s,
        [ExcelArgument(Name = "K",           Description = Arg.K)]           object k,
        [ExcelArgument(Name = "T",           Description = Arg.T)]           object t,
        [ExcelArgument(Name = "r",           Description = Arg.R)]           object r,
        [ExcelArgument(Name = "isPut",       Description = Arg.IsPut)]       object isPut)
        => Fn.Run(Category.Options, () => BlackScholes.ImpliedVolatility(
               In.Price("marketPrice", marketPrice), In.Price("S", s), In.Price("K", k),
               In.Years("T", t), In.Rate("r", r), In.Flag("isPut", isPut)));

    [ExcelFunction(Name = OptionsConstants.VannaName, Category = Cat.Options, IsThreadSafe = true,
        Description = OptionsConstants.VannaDesc)]
    public static object BsVanna(
        [ExcelArgument(Name = "S",     Description = Arg.S)]     object s,
        [ExcelArgument(Name = "K",     Description = Arg.K)]     object k,
        [ExcelArgument(Name = "T",     Description = Arg.T)]     object t,
        [ExcelArgument(Name = "r",     Description = Arg.R)]     object r,
        [ExcelArgument(Name = "sigma", Description = Arg.Sigma)] object sigma)
        => Fn.Run(Category.Options, () => BlackScholes.Vanna(
               In.Price("S", s), In.Price("K", k), In.Years("T", t),
               In.Rate("r", r), In.Vol("sigma", sigma)));

    [ExcelFunction(Name = OptionsConstants.CharmName, Category = Cat.Options, IsThreadSafe = true,
        Description = OptionsConstants.CharmDesc)]
    public static object BsCharm(
        [ExcelArgument(Name = "S",     Description = Arg.S)]     object s,
        [ExcelArgument(Name = "K",     Description = Arg.K)]     object k,
        [ExcelArgument(Name = "T",     Description = Arg.T)]     object t,
        [ExcelArgument(Name = "r",     Description = Arg.R)]     object r,
        [ExcelArgument(Name = "sigma", Description = Arg.Sigma)] object sigma)
        => Fn.Run(Category.Options, () => BlackScholes.Charm(
               In.Price("S", s), In.Price("K", k), In.Years("T", t),
               In.Rate("r", r), In.Vol("sigma", sigma)));

    [ExcelFunction(Name = OptionsConstants.VolgaName, Category = Cat.Options, IsThreadSafe = true,
        Description = OptionsConstants.VolgaDesc)]
    public static object BsVolga(
        [ExcelArgument(Name = "S",     Description = Arg.S)]     object s,
        [ExcelArgument(Name = "K",     Description = Arg.K)]     object k,
        [ExcelArgument(Name = "T",     Description = Arg.T)]     object t,
        [ExcelArgument(Name = "r",     Description = Arg.R)]     object r,
        [ExcelArgument(Name = "sigma", Description = Arg.Sigma)] object sigma)
        => Fn.Run(Category.Options, () => BlackScholes.Volga(
               In.Price("S", s), In.Price("K", k), In.Years("T", t),
               In.Rate("r", r), In.Vol("sigma", sigma)));

    [ExcelFunction(Name = OptionsConstants.SpeedName, Category = Cat.Options, IsThreadSafe = true,
        Description = OptionsConstants.SpeedDesc)]
    public static object BsSpeed(
        [ExcelArgument(Name = "S",     Description = Arg.S)]     object s,
        [ExcelArgument(Name = "K",     Description = Arg.K)]     object k,
        [ExcelArgument(Name = "T",     Description = Arg.T)]     object t,
        [ExcelArgument(Name = "r",     Description = Arg.R)]     object r,
        [ExcelArgument(Name = "sigma", Description = Arg.Sigma)] object sigma)
        => Fn.Run(Category.Options, () => BlackScholes.Speed(
               In.Price("S", s), In.Price("K", k), In.Years("T", t),
               In.Rate("r", r), In.Vol("sigma", sigma)));

    [ExcelFunction(Name = OptionsConstants.ZommaName, Category = Cat.Options, IsThreadSafe = true,
        Description = OptionsConstants.ZommaDesc)]
    public static object BsZomma(
        [ExcelArgument(Name = "S",     Description = Arg.S)]     object s,
        [ExcelArgument(Name = "K",     Description = Arg.K)]     object k,
        [ExcelArgument(Name = "T",     Description = Arg.T)]     object t,
        [ExcelArgument(Name = "r",     Description = Arg.R)]     object r,
        [ExcelArgument(Name = "sigma", Description = Arg.Sigma)] object sigma)
        => Fn.Run(Category.Options, () => BlackScholes.Zomma(
               In.Price("S", s), In.Price("K", k), In.Years("T", t),
               In.Rate("r", r), In.Vol("sigma", sigma)));
}
