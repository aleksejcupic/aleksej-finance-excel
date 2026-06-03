using ExcelDna.Integration;
using Aleksej.Finance.Options;
using Aleksej.Finance.Excel.Constants;
using Aleksej.Finance.Excel.Helpers;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>Exotic option pricing: binary, barrier, Asian, and lookback options.</summary>
public static class ExoticOptionsFunctions
{
    [ExcelFunction(Name = ExoticOptionsConstants.BinaryCashName, Category = Cat.Options, IsThreadSafe = true,
        Description = ExoticOptionsConstants.BinaryCashDesc)]
    public static object ExBinaryCash(
        [ExcelArgument(Name = "S",          Description = Arg.S)]     object s,
        [ExcelArgument(Name = "K",          Description = Arg.K)]     object k,
        [ExcelArgument(Name = "T",          Description = Arg.T)]     object t,
        [ExcelArgument(Name = "r",          Description = Arg.R)]     object r,
        [ExcelArgument(Name = "sigma",      Description = Arg.Sigma)] object sigma,
        [ExcelArgument(Name = "cashPayoff", Description = ExoticOptionsConstants.CashPayoff)]      object cashPayoff,
        [ExcelArgument(Name = "isPut",      Description = ExoticOptionsConstants.BinaryCashIsPut)] object isPut)
        => Fn.Run(Category.Options, () => ExoticOptions.CashOrNothing(
               In.Price("S", s), In.Price("K", k), In.Years("T", t),
               In.Rate("r", r), In.Vol("sigma", sigma),
               In.Num("cashPayoff", cashPayoff, ExoticOptionsConstants.DefaultCashPayoff),
               In.Flag("isPut", isPut)));

    [ExcelFunction(Name = ExoticOptionsConstants.BinaryAssetName, Category = Cat.Options, IsThreadSafe = true,
        Description = ExoticOptionsConstants.BinaryAssetDesc)]
    public static object ExBinaryAsset(
        [ExcelArgument(Name = "S",     Description = Arg.S)]                        object s,
        [ExcelArgument(Name = "K",     Description = Arg.K)]                        object k,
        [ExcelArgument(Name = "T",     Description = ExoticOptionsConstants.Time)]  object t,
        [ExcelArgument(Name = "r",     Description = ExoticOptionsConstants.RiskFree)] object r,
        [ExcelArgument(Name = "sigma", Description = Arg.Sigma)]                    object sigma,
        [ExcelArgument(Name = "isPut", Description = ExoticOptionsConstants.IsPut)] object isPut)
        => Fn.Run(Category.Options, () => ExoticOptions.AssetOrNothing(
               In.Price("S", s), In.Price("K", k), In.Years("T", t),
               In.Rate("r", r), In.Vol("sigma", sigma), In.Flag("isPut", isPut)));

    [ExcelFunction(Name = ExoticOptionsConstants.BarrierCallName, Category = Cat.Options, IsThreadSafe = true,
        Description = ExoticOptionsConstants.BarrierCallDesc)]
    public static object ExBarrierCall(
        [ExcelArgument(Name = "S",       Description = Arg.S)]     object s,
        [ExcelArgument(Name = "K",       Description = Arg.K)]     object k,
        [ExcelArgument(Name = "H",       Description = ExoticOptionsConstants.Barrier)]        object h,
        [ExcelArgument(Name = "T",       Description = Arg.T)]     object t,
        [ExcelArgument(Name = "r",       Description = Arg.R)]     object r,
        [ExcelArgument(Name = "sigma",   Description = Arg.Sigma)] object sigma,
        [ExcelArgument(Name = "knockIn", Description = ExoticOptionsConstants.KnockInDefault)] object knockIn,
        [ExcelArgument(Name = "isUp",    Description = ExoticOptionsConstants.IsUp)]           object isUp)
        => Fn.Run(Category.Options, () => ExoticOptions.BarrierCall(
               In.Price("S", s), In.Price("K", k), In.Price("H", h),
               In.Years("T", t), In.Rate("r", r), In.Vol("sigma", sigma),
               In.Flag("knockIn", knockIn), In.Flag("isUp", isUp)));

    [ExcelFunction(Name = ExoticOptionsConstants.BarrierPutName, Category = Cat.Options, IsThreadSafe = true,
        Description = ExoticOptionsConstants.BarrierPutDesc)]
    public static object ExBarrierPut(
        [ExcelArgument(Name = "S",       Description = Arg.S)]                          object s,
        [ExcelArgument(Name = "K",       Description = Arg.K)]                          object k,
        [ExcelArgument(Name = "H",       Description = ExoticOptionsConstants.Barrier)] object h,
        [ExcelArgument(Name = "T",       Description = ExoticOptionsConstants.Time)]    object t,
        [ExcelArgument(Name = "r",       Description = ExoticOptionsConstants.RiskFree)] object r,
        [ExcelArgument(Name = "sigma",   Description = Arg.Sigma)]                      object sigma,
        [ExcelArgument(Name = "knockIn", Description = ExoticOptionsConstants.KnockIn)] object knockIn,
        [ExcelArgument(Name = "isUp",    Description = ExoticOptionsConstants.IsUpShort)] object isUp)
        => Fn.Run(Category.Options, () => ExoticOptions.BarrierPut(
               In.Price("S", s), In.Price("K", k), In.Price("H", h),
               In.Years("T", t), In.Rate("r", r), In.Vol("sigma", sigma),
               In.Flag("knockIn", knockIn), In.Flag("isUp", isUp)));

    [ExcelFunction(Name = ExoticOptionsConstants.AsianGeoName, Category = Cat.Options, IsThreadSafe = true,
        Description = ExoticOptionsConstants.AsianGeoDesc)]
    public static object ExAsianGeo(
        [ExcelArgument(Name = "S",     Description = Arg.S)]                        object s,
        [ExcelArgument(Name = "K",     Description = Arg.K)]                        object k,
        [ExcelArgument(Name = "T",     Description = ExoticOptionsConstants.Time)]  object t,
        [ExcelArgument(Name = "r",     Description = ExoticOptionsConstants.RiskFree)] object r,
        [ExcelArgument(Name = "sigma", Description = Arg.Sigma)]                    object sigma,
        [ExcelArgument(Name = "isPut", Description = ExoticOptionsConstants.IsPut)] object isPut)
        => Fn.Run(Category.Options, () => ExoticOptions.GeometricAsian(
               In.Price("S", s), In.Price("K", k), In.Years("T", t),
               In.Rate("r", r), In.Vol("sigma", sigma), In.Flag("isPut", isPut)));

    [ExcelFunction(Name = ExoticOptionsConstants.AsianArithName, Category = Cat.Options, IsThreadSafe = true,
        Description = ExoticOptionsConstants.AsianArithDesc)]
    public static object ExAsianArith(
        [ExcelArgument(Name = "S",          Description = Arg.S)]                        object s,
        [ExcelArgument(Name = "K",          Description = Arg.K)]                        object k,
        [ExcelArgument(Name = "T",          Description = ExoticOptionsConstants.Time)]  object t,
        [ExcelArgument(Name = "r",          Description = ExoticOptionsConstants.RiskFree)] object r,
        [ExcelArgument(Name = "sigma",      Description = Arg.Sigma)]                    object sigma,
        [ExcelArgument(Name = "monitoring", Description = ExoticOptionsConstants.Monitoring)] object monitoring,
        [ExcelArgument(Name = "paths",      Description = ExoticOptionsConstants.Paths)]      object paths,
        [ExcelArgument(Name = "isPut",      Description = ExoticOptionsConstants.IsPut)]      object isPut,
        [ExcelArgument(Name = "seed",       Description = ExoticOptionsConstants.Seed)]       object seed)
        => Fn.Run(Category.Options, () => ExoticOptions.ArithmeticAsian(
               In.Price("S", s), In.Price("K", k), In.Years("T", t),
               In.Rate("r", r), In.Vol("sigma", sigma),
               In.PosInt("monitoring", monitoring, ExoticOptionsConstants.DefaultMonitoring),
               In.PosInt("paths", paths, ExoticOptionsConstants.DefaultPaths),
               In.Flag("isPut", isPut),
               In.PosInt("seed", seed, ExoticOptionsConstants.DefaultSeed)));

    [ExcelFunction(Name = ExoticOptionsConstants.LookbackCallName, Category = Cat.Options, IsThreadSafe = true,
        Description = ExoticOptionsConstants.LookbackCallDesc)]
    public static object ExLookbackCall(
        [ExcelArgument(Name = "S",    Description = Arg.S)]                            object s,
        [ExcelArgument(Name = "sMin", Description = ExoticOptionsConstants.SMin)]      object sMin,
        [ExcelArgument(Name = "T",    Description = ExoticOptionsConstants.LookbackTime)] object t,
        [ExcelArgument(Name = "r",    Description = Arg.R)]                            object r,
        [ExcelArgument(Name = "sigma",Description = Arg.Sigma)]                        object sigma)
        => Fn.Run(Category.Options, () => ExoticOptions.LookbackCall(
               In.Price("S", s), In.Price("sMin", sMin), In.Years("T", t),
               In.Rate("r", r), In.Vol("sigma", sigma)));

    [ExcelFunction(Name = ExoticOptionsConstants.LookbackPutName, Category = Cat.Options, IsThreadSafe = true,
        Description = ExoticOptionsConstants.LookbackPutDesc)]
    public static object ExLookbackPut(
        [ExcelArgument(Name = "S",    Description = Arg.S)]                            object s,
        [ExcelArgument(Name = "sMax", Description = ExoticOptionsConstants.SMax)]      object sMax,
        [ExcelArgument(Name = "T",    Description = ExoticOptionsConstants.LookbackTime)] object t,
        [ExcelArgument(Name = "r",    Description = Arg.R)]                            object r,
        [ExcelArgument(Name = "sigma",Description = Arg.Sigma)]                        object sigma)
        => Fn.Run(Category.Options, () => ExoticOptions.LookbackPut(
               In.Price("S", s), In.Price("sMax", sMax), In.Years("T", t),
               In.Rate("r", r), In.Vol("sigma", sigma)));
}
