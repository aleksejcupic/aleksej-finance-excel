using ExcelDna.Integration;
using AleksejCupic.FinancialMath.Derivatives;
using Aleksej.Finance.Excel.Helpers;
using Aleksej.Finance.Excel.Settings;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>Vasicek and CIR short-rate models for term structure and bond pricing (Hull Ch. 31-32).</summary>
public static class ShortRateFunctions
{
    private static bool Enabled => UserSettings.Load().EnableDerivatives;
    private static string Off   => RangeHelper.DisabledMessage("Derivatives");

    [ExcelFunction(Name = "SR_VASICEK_PRICE", Category = "Finance | Derivatives", IsThreadSafe = true,
        Description = "Zero-coupon bond price under the Vasicek model. dr = kappa*(theta-r)*dt + sigma*dW. P = A(tau)*exp(-B(tau)*r).",
        HelpTopic = "https://aleksejcupic.github.io/financial-math/derivatives/short-rate-models")]
    public static object SrVasicekPrice(
        [ExcelArgument(Name = "r",     Description = "Current short rate")]          object r,
        [ExcelArgument(Name = "tau",   Description = "Time to maturity in years")]   object tau,
        [ExcelArgument(Name = "kappa", Description = "Mean-reversion speed")]        object kappa,
        [ExcelArgument(Name = "theta", Description = "Long-run mean rate")]          object theta,
        [ExcelArgument(Name = "sigma", Description = "Short-rate volatility")]       object sigma)
        => Enabled ? ShortRateModels.VasicekBondPrice(
                         RangeHelper.Scalar(r), RangeHelper.Scalar(tau), RangeHelper.Scalar(kappa),
                         RangeHelper.Scalar(theta), RangeHelper.Scalar(sigma))
                   : (object)Off;

    [ExcelFunction(Name = "SR_VASICEK_YIELD", Category = "Finance | Derivatives", IsThreadSafe = true,
        Description = "Continuously compounded zero yield under the Vasicek model. R = -ln(P)/tau.")]
    public static object SrVasicekYield(
        [ExcelArgument(Name = "r",     Description = "Current short rate")]     object r,
        [ExcelArgument(Name = "tau",   Description = "Maturity in years")]      object tau,
        [ExcelArgument(Name = "kappa", Description = "Mean-reversion speed")]   object kappa,
        [ExcelArgument(Name = "theta", Description = "Long-run mean rate")]     object theta,
        [ExcelArgument(Name = "sigma", Description = "Short-rate volatility")]  object sigma)
        => Enabled ? ShortRateModels.VasicekYield(
                         RangeHelper.Scalar(r), RangeHelper.Scalar(tau), RangeHelper.Scalar(kappa),
                         RangeHelper.Scalar(theta), RangeHelper.Scalar(sigma))
                   : (object)Off;

    [ExcelFunction(Name = "SR_VASICEK_LRYIELD", Category = "Finance | Derivatives", IsThreadSafe = true,
        Description = "Vasicek long-run yield as tau → infinity. R(∞) = theta - sigma²/(2*kappa²).")]
    public static object SrVasicekLrYield(
        [ExcelArgument(Name = "kappa", Description = "Mean-reversion speed")]   object kappa,
        [ExcelArgument(Name = "theta", Description = "Long-run mean rate")]     object theta,
        [ExcelArgument(Name = "sigma", Description = "Short-rate volatility")]  object sigma)
        => Enabled ? ShortRateModels.VasicekLongRunYield(
                         RangeHelper.Scalar(kappa), RangeHelper.Scalar(theta), RangeHelper.Scalar(sigma))
                   : (object)Off;

    [ExcelFunction(Name = "SR_VASICEK_OPTION", Category = "Finance | Derivatives", IsThreadSafe = true,
        Description = "European call or put option on a zero-coupon bond under the Vasicek model (Jamshidian 1989).")]
    public static object SrVasicekOption(
        [ExcelArgument(Name = "r",        Description = "Current short rate")]              object r,
        [ExcelArgument(Name = "T",        Description = "Option expiry in years")]          object t,
        [ExcelArgument(Name = "maturity", Description = "Bond maturity in years (> T)")]   object maturity,
        [ExcelArgument(Name = "K",        Description = "Option strike price")]             object k,
        [ExcelArgument(Name = "kappa",    Description = "Mean-reversion speed")]            object kappa,
        [ExcelArgument(Name = "theta",    Description = "Long-run mean rate")]              object theta,
        [ExcelArgument(Name = "sigma",    Description = "Short-rate volatility")]           object sigma,
        [ExcelArgument(Name = "isPut",    Description = "TRUE for put, FALSE for call")]    object isPut)
        => Enabled ? ShortRateModels.VasicekBondOption(
                         RangeHelper.Scalar(r), RangeHelper.Scalar(t), RangeHelper.Scalar(maturity),
                         RangeHelper.Scalar(k), RangeHelper.Scalar(kappa), RangeHelper.Scalar(theta),
                         RangeHelper.Scalar(sigma), RangeHelper.ScalarBool(isPut))
                   : (object)Off;

    [ExcelFunction(Name = "SR_CIR_PRICE", Category = "Finance | Derivatives", IsThreadSafe = true,
        Description = "Zero-coupon bond price under the CIR model. dr = kappa*(theta-r)*dt + sigma*sqrt(r)*dW. Square-root keeps r non-negative.")]
    public static object SrCirPrice(
        [ExcelArgument(Name = "r",     Description = "Current short rate")]     object r,
        [ExcelArgument(Name = "tau",   Description = "Maturity in years")]      object tau,
        [ExcelArgument(Name = "kappa", Description = "Mean-reversion speed")]   object kappa,
        [ExcelArgument(Name = "theta", Description = "Long-run mean rate")]     object theta,
        [ExcelArgument(Name = "sigma", Description = "Volatility coefficient")] object sigma)
        => Enabled ? ShortRateModels.CirBondPrice(
                         RangeHelper.Scalar(r), RangeHelper.Scalar(tau), RangeHelper.Scalar(kappa),
                         RangeHelper.Scalar(theta), RangeHelper.Scalar(sigma))
                   : (object)Off;

    [ExcelFunction(Name = "SR_CIR_YIELD", Category = "Finance | Derivatives", IsThreadSafe = true,
        Description = "Continuously compounded zero yield under the CIR model. R = -ln(P)/tau.")]
    public static object SrCirYield(
        [ExcelArgument(Name = "r",     Description = "Current short rate")]     object r,
        [ExcelArgument(Name = "tau",   Description = "Maturity in years")]      object tau,
        [ExcelArgument(Name = "kappa", Description = "Mean-reversion speed")]   object kappa,
        [ExcelArgument(Name = "theta", Description = "Long-run mean rate")]     object theta,
        [ExcelArgument(Name = "sigma", Description = "Volatility coefficient")] object sigma)
        => Enabled ? ShortRateModels.CirYield(
                         RangeHelper.Scalar(r), RangeHelper.Scalar(tau), RangeHelper.Scalar(kappa),
                         RangeHelper.Scalar(theta), RangeHelper.Scalar(sigma))
                   : (object)Off;

    [ExcelFunction(Name = "SR_CIR_LRYIELD", Category = "Finance | Derivatives", IsThreadSafe = true,
        Description = "CIR long-run yield. R(∞) = 2*kappa*theta / (kappa + gamma), where gamma = sqrt(kappa² + 2*sigma²).")]
    public static object SrCirLrYield(
        [ExcelArgument(Name = "kappa", Description = "Mean-reversion speed")]   object kappa,
        [ExcelArgument(Name = "theta", Description = "Long-run mean rate")]     object theta,
        [ExcelArgument(Name = "sigma", Description = "Volatility coefficient")] object sigma)
        => Enabled ? ShortRateModels.CirLongRunYield(
                         RangeHelper.Scalar(kappa), RangeHelper.Scalar(theta), RangeHelper.Scalar(sigma))
                   : (object)Off;
}
