using ExcelDna.Integration;
using Aleksej.Finance.Options;
using Aleksej.Finance.Excel.Constants;
using Aleksej.Finance.Excel.Helpers;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>Garman-Kohlhagen FX option pricing and Greeks.</summary>
public static class GarmanKohlhagenFunctions
{
    [ExcelFunction(Name = GarmanKohlhagenConstants.CallName, Category = Cat.Options, IsThreadSafe = true,
        Description = GarmanKohlhagenConstants.CallDesc)]
    public static object GkCall(
        [ExcelArgument(Name = "S",     Description = GarmanKohlhagenConstants.SpotFull)]     object s,
        [ExcelArgument(Name = "K",     Description = GarmanKohlhagenConstants.StrikeFx)]     object k,
        [ExcelArgument(Name = "T",     Description = Arg.T)]                                 object t,
        [ExcelArgument(Name = "r",     Description = GarmanKohlhagenConstants.DomesticFull)] object r,
        [ExcelArgument(Name = "rf",    Description = GarmanKohlhagenConstants.ForeignFull)]  object rf,
        [ExcelArgument(Name = "sigma", Description = GarmanKohlhagenConstants.SigmaFx)]      object sigma)
        => Fn.Run(Category.Options, () => GarmanKohlhagen.Call(
               In.Price("S", s), In.Price("K", k), In.Years("T", t),
               In.Rate("r", r), In.Rate("rf", rf), In.Vol("sigma", sigma)));

    [ExcelFunction(Name = GarmanKohlhagenConstants.PutName, Category = Cat.Options, IsThreadSafe = true,
        Description = GarmanKohlhagenConstants.PutDesc)]
    public static object GkPut(
        [ExcelArgument(Name = "S",     Description = GarmanKohlhagenConstants.Spot)]     object s,
        [ExcelArgument(Name = "K",     Description = GarmanKohlhagenConstants.StrikeFx)] object k,
        [ExcelArgument(Name = "T",     Description = Arg.T)]                             object t,
        [ExcelArgument(Name = "r",     Description = GarmanKohlhagenConstants.Domestic)] object r,
        [ExcelArgument(Name = "rf",    Description = GarmanKohlhagenConstants.Foreign)]  object rf,
        [ExcelArgument(Name = "sigma", Description = Arg.Sigma)]                         object sigma)
        => Fn.Run(Category.Options, () => GarmanKohlhagen.Put(
               In.Price("S", s), In.Price("K", k), In.Years("T", t),
               In.Rate("r", r), In.Rate("rf", rf), In.Vol("sigma", sigma)));

    [ExcelFunction(Name = GarmanKohlhagenConstants.DeltaName, Category = Cat.Options, IsThreadSafe = true,
        Description = GarmanKohlhagenConstants.DeltaDesc)]
    public static object GkDelta(
        [ExcelArgument(Name = "S",     Description = GarmanKohlhagenConstants.Spot)]     object s,
        [ExcelArgument(Name = "K",     Description = GarmanKohlhagenConstants.StrikeFx)] object k,
        [ExcelArgument(Name = "T",     Description = Arg.T)]                             object t,
        [ExcelArgument(Name = "r",     Description = GarmanKohlhagenConstants.Domestic)] object r,
        [ExcelArgument(Name = "rf",    Description = GarmanKohlhagenConstants.Foreign)]  object rf,
        [ExcelArgument(Name = "sigma", Description = Arg.Sigma)]                         object sigma,
        [ExcelArgument(Name = "isPut", Description = Arg.IsPut)]                         object isPut)
        => Fn.Run(Category.Options, () => GarmanKohlhagen.Delta(
               In.Price("S", s), In.Price("K", k), In.Years("T", t),
               In.Rate("r", r), In.Rate("rf", rf), In.Vol("sigma", sigma),
               In.Flag("isPut", isPut)));

    [ExcelFunction(Name = GarmanKohlhagenConstants.GammaName, Category = Cat.Options, IsThreadSafe = true,
        Description = GarmanKohlhagenConstants.GammaDesc)]
    public static object GkGamma(
        [ExcelArgument(Name = "S",     Description = GarmanKohlhagenConstants.Spot)]     object s,
        [ExcelArgument(Name = "K",     Description = GarmanKohlhagenConstants.StrikeFx)] object k,
        [ExcelArgument(Name = "T",     Description = Arg.T)]                             object t,
        [ExcelArgument(Name = "r",     Description = GarmanKohlhagenConstants.Domestic)] object r,
        [ExcelArgument(Name = "rf",    Description = GarmanKohlhagenConstants.Foreign)]  object rf,
        [ExcelArgument(Name = "sigma", Description = Arg.Sigma)]                         object sigma)
        => Fn.Run(Category.Options, () => GarmanKohlhagen.Gamma(
               In.Price("S", s), In.Price("K", k), In.Years("T", t),
               In.Rate("r", r), In.Rate("rf", rf), In.Vol("sigma", sigma)));

    [ExcelFunction(Name = GarmanKohlhagenConstants.VegaName, Category = Cat.Options, IsThreadSafe = true,
        Description = GarmanKohlhagenConstants.VegaDesc)]
    public static object GkVega(
        [ExcelArgument(Name = "S",     Description = GarmanKohlhagenConstants.Spot)]     object s,
        [ExcelArgument(Name = "K",     Description = GarmanKohlhagenConstants.StrikeFx)] object k,
        [ExcelArgument(Name = "T",     Description = Arg.T)]                             object t,
        [ExcelArgument(Name = "r",     Description = GarmanKohlhagenConstants.Domestic)] object r,
        [ExcelArgument(Name = "rf",    Description = GarmanKohlhagenConstants.Foreign)]  object rf,
        [ExcelArgument(Name = "sigma", Description = Arg.Sigma)]                         object sigma)
        => Fn.Run(Category.Options, () => GarmanKohlhagen.Vega(
               In.Price("S", s), In.Price("K", k), In.Years("T", t),
               In.Rate("r", r), In.Rate("rf", rf), In.Vol("sigma", sigma)));

    [ExcelFunction(Name = GarmanKohlhagenConstants.ThetaName, Category = Cat.Options, IsThreadSafe = true,
        Description = GarmanKohlhagenConstants.ThetaDesc)]
    public static object GkTheta(
        [ExcelArgument(Name = "S",     Description = GarmanKohlhagenConstants.Spot)]     object s,
        [ExcelArgument(Name = "K",     Description = GarmanKohlhagenConstants.StrikeFx)] object k,
        [ExcelArgument(Name = "T",     Description = Arg.T)]                             object t,
        [ExcelArgument(Name = "r",     Description = GarmanKohlhagenConstants.Domestic)] object r,
        [ExcelArgument(Name = "rf",    Description = GarmanKohlhagenConstants.Foreign)]  object rf,
        [ExcelArgument(Name = "sigma", Description = Arg.Sigma)]                         object sigma,
        [ExcelArgument(Name = "isPut", Description = Arg.IsPut)]                         object isPut)
        => Fn.Run(Category.Options, () => GarmanKohlhagen.Theta(
               In.Price("S", s), In.Price("K", k), In.Years("T", t),
               In.Rate("r", r), In.Rate("rf", rf), In.Vol("sigma", sigma),
               In.Flag("isPut", isPut)));

    [ExcelFunction(Name = GarmanKohlhagenConstants.RhoName, Category = Cat.Options, IsThreadSafe = true,
        Description = GarmanKohlhagenConstants.RhoDesc)]
    public static object GkRho(
        [ExcelArgument(Name = "S",     Description = GarmanKohlhagenConstants.Spot)]     object s,
        [ExcelArgument(Name = "K",     Description = GarmanKohlhagenConstants.StrikeFx)] object k,
        [ExcelArgument(Name = "T",     Description = Arg.T)]                             object t,
        [ExcelArgument(Name = "r",     Description = GarmanKohlhagenConstants.Domestic)] object r,
        [ExcelArgument(Name = "rf",    Description = GarmanKohlhagenConstants.Foreign)]  object rf,
        [ExcelArgument(Name = "sigma", Description = Arg.Sigma)]                         object sigma,
        [ExcelArgument(Name = "isPut", Description = Arg.IsPut)]                         object isPut)
        => Fn.Run(Category.Options, () => GarmanKohlhagen.Rho(
               In.Price("S", s), In.Price("K", k), In.Years("T", t),
               In.Rate("r", r), In.Rate("rf", rf), In.Vol("sigma", sigma),
               In.Flag("isPut", isPut)));

    [ExcelFunction(Name = GarmanKohlhagenConstants.RhoForeignName, Category = Cat.Options, IsThreadSafe = true,
        Description = GarmanKohlhagenConstants.RhoForeignDesc)]
    public static object GkRhoForeign(
        [ExcelArgument(Name = "S",     Description = GarmanKohlhagenConstants.Spot)]     object s,
        [ExcelArgument(Name = "K",     Description = GarmanKohlhagenConstants.StrikeFx)] object k,
        [ExcelArgument(Name = "T",     Description = Arg.T)]                             object t,
        [ExcelArgument(Name = "r",     Description = GarmanKohlhagenConstants.Domestic)] object r,
        [ExcelArgument(Name = "rf",    Description = GarmanKohlhagenConstants.Foreign)]  object rf,
        [ExcelArgument(Name = "sigma", Description = Arg.Sigma)]                         object sigma,
        [ExcelArgument(Name = "isPut", Description = Arg.IsPut)]                         object isPut)
        => Fn.Run(Category.Options, () => GarmanKohlhagen.RhoForeign(
               In.Price("S", s), In.Price("K", k), In.Years("T", t),
               In.Rate("r", r), In.Rate("rf", rf), In.Vol("sigma", sigma),
               In.Flag("isPut", isPut)));

    [ExcelFunction(Name = GarmanKohlhagenConstants.IvName, Category = Cat.Options, IsThreadSafe = true,
        Description = GarmanKohlhagenConstants.IvDesc)]
    public static object GkIv(
        [ExcelArgument(Name = "marketPrice", Description = GarmanKohlhagenConstants.MarketPrice)]  object marketPrice,
        [ExcelArgument(Name = "S",           Description = GarmanKohlhagenConstants.Spot)]         object s,
        [ExcelArgument(Name = "K",           Description = GarmanKohlhagenConstants.StrikeFx)]     object k,
        [ExcelArgument(Name = "T",           Description = GarmanKohlhagenConstants.Time)]         object t,
        [ExcelArgument(Name = "r",           Description = GarmanKohlhagenConstants.DomesticShort)] object r,
        [ExcelArgument(Name = "rf",          Description = GarmanKohlhagenConstants.ForeignShort)] object rf,
        [ExcelArgument(Name = "isPut",       Description = GarmanKohlhagenConstants.IsPut)]        object isPut)
        => Fn.Run(Category.Options, () => GarmanKohlhagen.ImpliedVolatility(
               In.Price("marketPrice", marketPrice), In.Price("S", s), In.Price("K", k),
               In.Years("T", t), In.Rate("r", r), In.Rate("rf", rf),
               In.Flag("isPut", isPut)));
}
