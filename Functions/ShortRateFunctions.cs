using ExcelDna.Integration;
using Aleksej.Finance.Derivatives;
using Aleksej.Finance.Excel.Constants;
using Aleksej.Finance.Excel.Helpers;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>Vasicek and CIR short-rate models for term structure and bond pricing (Hull Ch. 31-32).</summary>
public static class ShortRateFunctions
{
    [ExcelFunction(Name = ShortRateConstants.VasicekPriceName, Category = Cat.Derivatives, IsThreadSafe = true,
        Description = ShortRateConstants.VasicekPriceDesc, HelpTopic = ShortRateConstants.Help)]
    public static object SrVasicekPrice(
        [ExcelArgument(Name = "r",     Description = ShortRateConstants.ShortRate)]   object r,
        [ExcelArgument(Name = "tau",   Description = ShortRateConstants.TauMaturity)] object tau,
        [ExcelArgument(Name = "kappa", Description = ShortRateConstants.Kappa)]       object kappa,
        [ExcelArgument(Name = "theta", Description = ShortRateConstants.Theta)]       object theta,
        [ExcelArgument(Name = "sigma", Description = ShortRateConstants.SigmaVol)]    object sigma)
        => Fn.Run(Category.Derivatives, () => ShortRateModels.VasicekBondPrice(
               In.Num("r", r), In.Years("tau", tau), In.Num("kappa", kappa),
               In.Num("theta", theta), In.Vol("sigma", sigma)));

    [ExcelFunction(Name = ShortRateConstants.VasicekYieldName, Category = Cat.Derivatives, IsThreadSafe = true,
        Description = ShortRateConstants.VasicekYieldDesc, HelpTopic = ShortRateConstants.Help)]
    public static object SrVasicekYield(
        [ExcelArgument(Name = "r",     Description = ShortRateConstants.ShortRate)] object r,
        [ExcelArgument(Name = "tau",   Description = ShortRateConstants.Maturity)]  object tau,
        [ExcelArgument(Name = "kappa", Description = ShortRateConstants.Kappa)]     object kappa,
        [ExcelArgument(Name = "theta", Description = ShortRateConstants.Theta)]     object theta,
        [ExcelArgument(Name = "sigma", Description = ShortRateConstants.SigmaVol)]  object sigma)
        => Fn.Run(Category.Derivatives, () => ShortRateModels.VasicekYield(
               In.Num("r", r), In.Years("tau", tau), In.Num("kappa", kappa),
               In.Num("theta", theta), In.Vol("sigma", sigma)));

    [ExcelFunction(Name = ShortRateConstants.VasicekLrYieldName, Category = Cat.Derivatives, IsThreadSafe = true,
        Description = ShortRateConstants.VasicekLrYieldDesc, HelpTopic = ShortRateConstants.Help)]
    public static object SrVasicekLrYield(
        [ExcelArgument(Name = "kappa", Description = ShortRateConstants.Kappa)]    object kappa,
        [ExcelArgument(Name = "theta", Description = ShortRateConstants.Theta)]    object theta,
        [ExcelArgument(Name = "sigma", Description = ShortRateConstants.SigmaVol)] object sigma)
        => Fn.Run(Category.Derivatives, () => ShortRateModels.VasicekLongRunYield(
               In.Num("kappa", kappa), In.Num("theta", theta), In.Vol("sigma", sigma)));

    [ExcelFunction(Name = ShortRateConstants.VasicekOptionName, Category = Cat.Derivatives, IsThreadSafe = true,
        Description = ShortRateConstants.VasicekOptionDesc, HelpTopic = ShortRateConstants.Help)]
    public static object SrVasicekOption(
        [ExcelArgument(Name = "r",        Description = ShortRateConstants.ShortRate)]    object r,
        [ExcelArgument(Name = "T",        Description = ShortRateConstants.OptionExpiry)] object t,
        [ExcelArgument(Name = "maturity", Description = ShortRateConstants.BondMaturity)] object maturity,
        [ExcelArgument(Name = "K",        Description = ShortRateConstants.Strike)]       object k,
        [ExcelArgument(Name = "kappa",    Description = ShortRateConstants.Kappa)]        object kappa,
        [ExcelArgument(Name = "theta",    Description = ShortRateConstants.Theta)]        object theta,
        [ExcelArgument(Name = "sigma",    Description = ShortRateConstants.SigmaVol)]     object sigma,
        [ExcelArgument(Name = "isPut",    Description = ShortRateConstants.IsPut)]        object isPut)
        => Fn.Run(Category.Derivatives, () => ShortRateModels.VasicekBondOption(
               In.Num("r", r), In.Years("T", t), In.Years("maturity", maturity),
               In.Price("K", k), In.Num("kappa", kappa), In.Num("theta", theta),
               In.Vol("sigma", sigma), In.Flag("isPut", isPut)));

    [ExcelFunction(Name = ShortRateConstants.CirPriceName, Category = Cat.Derivatives, IsThreadSafe = true,
        Description = ShortRateConstants.CirPriceDesc, HelpTopic = ShortRateConstants.Help)]
    public static object SrCirPrice(
        [ExcelArgument(Name = "r",     Description = ShortRateConstants.ShortRate)]  object r,
        [ExcelArgument(Name = "tau",   Description = ShortRateConstants.Maturity)]   object tau,
        [ExcelArgument(Name = "kappa", Description = ShortRateConstants.Kappa)]      object kappa,
        [ExcelArgument(Name = "theta", Description = ShortRateConstants.Theta)]      object theta,
        [ExcelArgument(Name = "sigma", Description = ShortRateConstants.SigmaCoeff)] object sigma)
        => Fn.Run(Category.Derivatives, () => ShortRateModels.CirBondPrice(
               In.Num("r", r), In.Years("tau", tau), In.Num("kappa", kappa),
               In.Num("theta", theta), In.Vol("sigma", sigma)));

    [ExcelFunction(Name = ShortRateConstants.CirYieldName, Category = Cat.Derivatives, IsThreadSafe = true,
        Description = ShortRateConstants.CirYieldDesc, HelpTopic = ShortRateConstants.Help)]
    public static object SrCirYield(
        [ExcelArgument(Name = "r",     Description = ShortRateConstants.ShortRate)]  object r,
        [ExcelArgument(Name = "tau",   Description = ShortRateConstants.Maturity)]   object tau,
        [ExcelArgument(Name = "kappa", Description = ShortRateConstants.Kappa)]      object kappa,
        [ExcelArgument(Name = "theta", Description = ShortRateConstants.Theta)]      object theta,
        [ExcelArgument(Name = "sigma", Description = ShortRateConstants.SigmaCoeff)] object sigma)
        => Fn.Run(Category.Derivatives, () => ShortRateModels.CirYield(
               In.Num("r", r), In.Years("tau", tau), In.Num("kappa", kappa),
               In.Num("theta", theta), In.Vol("sigma", sigma)));

    [ExcelFunction(Name = ShortRateConstants.CirLrYieldName, Category = Cat.Derivatives, IsThreadSafe = true,
        Description = ShortRateConstants.CirLrYieldDesc, HelpTopic = ShortRateConstants.Help)]
    public static object SrCirLrYield(
        [ExcelArgument(Name = "kappa", Description = ShortRateConstants.Kappa)]      object kappa,
        [ExcelArgument(Name = "theta", Description = ShortRateConstants.Theta)]      object theta,
        [ExcelArgument(Name = "sigma", Description = ShortRateConstants.SigmaCoeff)] object sigma)
        => Fn.Run(Category.Derivatives, () => ShortRateModels.CirLongRunYield(
               In.Num("kappa", kappa), In.Num("theta", theta), In.Vol("sigma", sigma)));
}
