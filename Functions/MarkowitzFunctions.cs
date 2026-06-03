using ExcelDna.Integration;
using Aleksej.Finance.Portfolio;
using Aleksej.Finance.Excel.Constants;
using Aleksej.Finance.Excel.Helpers;
using Aleksej.Finance.Excel.Settings;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>Markowitz mean-variance portfolio optimisation functions.</summary>
public static class MarkowitzFunctions
{
    [ExcelFunction(Name = MarkowitzConstants.ReturnName, Category = Cat.Portfolio, IsThreadSafe = true,
        Description = MarkowitzConstants.ReturnDesc, HelpTopic = MarkowitzConstants.Help)]
    public static object PortReturn(
        [ExcelArgument(Name = "weights", Description = MarkowitzConstants.WeightsSum1)] object weights,
        [ExcelArgument(Name = "mu",      Description = MarkowitzConstants.MuAnn)]       object mu)
        => Fn.Run(Category.PortfolioRisk, () => Markowitz.PortfolioReturn(
               In.Vector("weights", weights), In.Vector("mu", mu)));

    [ExcelFunction(Name = MarkowitzConstants.VolName, Category = Cat.Portfolio, IsThreadSafe = true,
        Description = MarkowitzConstants.VolDesc)]
    public static object PortVol(
        [ExcelArgument(Name = "weights", Description = MarkowitzConstants.Weights)] object weights,
        [ExcelArgument(Name = "cov",     Description = MarkowitzConstants.CovAnn)]  object cov)
        => Fn.Run(Category.PortfolioRisk, () => Markowitz.PortfolioVolatility(
               In.Vector("weights", weights), In.Matrix("cov", cov)));

    [ExcelFunction(Name = MarkowitzConstants.SharpeName, Category = Cat.Portfolio, IsThreadSafe = true,
        Description = MarkowitzConstants.SharpeDesc)]
    public static object PortSharpe(
        [ExcelArgument(Name = "weights", Description = MarkowitzConstants.Weights)] object weights,
        [ExcelArgument(Name = "mu",      Description = MarkowitzConstants.Mu)]      object mu,
        [ExcelArgument(Name = "cov",     Description = MarkowitzConstants.Cov)]     object cov,
        [ExcelArgument(Name = "rf",      Description = MarkowitzConstants.Rf)]      object rf)
        => Fn.Run(Category.PortfolioRisk, () => Markowitz.PortfolioSharpe(
               In.Vector("weights", weights), In.Vector("mu", mu),
               In.Matrix("cov", cov),
               In.Rate("rf", rf, UserSettings.Current.DefaultRiskFreeRate)));

    [ExcelFunction(Name = MarkowitzConstants.MinVarName, Category = Cat.Portfolio, IsThreadSafe = true,
        Description = MarkowitzConstants.MinVarDesc)]
    public static object PortMinVar(
        [ExcelArgument(Name = "cov", Description = MarkowitzConstants.CovAnn)] object cov)
        => Fn.Run(Category.PortfolioRisk, () =>
        {
            var w = Markowitz.MinVariancePortfolio(In.Matrix("cov", cov));
            var result = new object[w.Length, 1];
            for (int i = 0; i < w.Length; i++) result[i, 0] = w[i];
            return result;
        });

    [ExcelFunction(Name = MarkowitzConstants.MaxSharpeName, Category = Cat.Portfolio, IsThreadSafe = true,
        Description = MarkowitzConstants.MaxSharpeDesc)]
    public static object PortMaxSharpe(
        [ExcelArgument(Name = "mu",  Description = MarkowitzConstants.MuAnnShort)] object mu,
        [ExcelArgument(Name = "cov", Description = MarkowitzConstants.Cov)]        object cov,
        [ExcelArgument(Name = "rf",  Description = MarkowitzConstants.Rf)]         object rf)
        => Fn.Run(Category.PortfolioRisk, () =>
        {
            var w = Markowitz.MaxSharpePortfolioConstrained(
                In.Vector("mu", mu), In.Matrix("cov", cov),
                In.Rate("rf", rf, UserSettings.Current.DefaultRiskFreeRate));
            var result = new object[w.Length, 1];
            for (int i = 0; i < w.Length; i++) result[i, 0] = w[i];
            return result;
        });

    [ExcelFunction(Name = MarkowitzConstants.RiskParityName, Category = Cat.Portfolio, IsThreadSafe = true,
        Description = MarkowitzConstants.RiskParityDesc)]
    public static object PortRiskParity(
        [ExcelArgument(Name = "cov", Description = MarkowitzConstants.CovAnn)] object cov)
        => Fn.Run(Category.PortfolioRisk, () =>
        {
            var w = Markowitz.RiskParityPortfolio(In.Matrix("cov", cov));
            var result = new object[w.Length, 1];
            for (int i = 0; i < w.Length; i++) result[i, 0] = w[i];
            return result;
        });

    [ExcelFunction(Name = MarkowitzConstants.RiskContribName, Category = Cat.Portfolio, IsThreadSafe = true,
        Description = MarkowitzConstants.RiskContribDesc)]
    public static object PortRiskContrib(
        [ExcelArgument(Name = "weights", Description = MarkowitzConstants.Weights)] object weights,
        [ExcelArgument(Name = "cov",     Description = MarkowitzConstants.Cov)]     object cov)
        => Fn.Run(Category.PortfolioRisk, () =>
        {
            var rc = Markowitz.RiskContributions(In.Vector("weights", weights), In.Matrix("cov", cov));
            var result = new object[rc.Length, 1];
            for (int i = 0; i < rc.Length; i++) result[i, 0] = rc[i];
            return result;
        });
}
