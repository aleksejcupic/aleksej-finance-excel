using ExcelDna.Integration;
using Aleksej.Finance.Bonds;
using Aleksej.Finance.Excel.Constants;
using Aleksej.Finance.Excel.Helpers;
using Aleksej.Finance.Excel.Settings;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>Fixed-income functions: bond math, yield curve, and mortgage calculations.</summary>
public static class BondFunctions
{
    // ── BondMath ──────────────────────────────────────────────────────────────

    [ExcelFunction(Name = BondConstants.PriceName, Category = Cat.Bonds, IsThreadSafe = true,
        Description = BondConstants.PriceDesc, HelpTopic = BondConstants.HelpBondMath)]
    public static object BondPrice(
        [ExcelArgument(Name = "face",       Description = Arg.Face)]                              object face,
        [ExcelArgument(Name = "couponRate", Description = Arg.CouponRate)]                        object couponRate,
        [ExcelArgument(Name = "ytm",        Description = BondConstants.ArgYtm)]                  object ytm,
        [ExcelArgument(Name = "years",      Description = BondConstants.ArgYears)]                object years,
        [ExcelArgument(Name = "frequency",  Description = BondConstants.ArgFrequencyDefault)]    object frequency)
        => Fn.Run(Category.Bonds, () => BondMath.Price(
               In.Price("face", face), In.Rate("couponRate", couponRate),
               In.Rate("ytm", ytm), In.Years("years", years),
               In.PosInt("frequency", frequency, UserSettings.Current.DefaultFrequency)));

    [ExcelFunction(Name = BondConstants.YtmName, Category = Cat.Bonds, IsThreadSafe = true,
        Description = BondConstants.YtmDesc)]
    public static object BondYtm(
        [ExcelArgument(Name = "price",      Description = BondConstants.ArgPrice)]               object price,
        [ExcelArgument(Name = "face",       Description = Arg.Face)]                              object face,
        [ExcelArgument(Name = "couponRate", Description = BondConstants.ArgCouponRatePlain)]      object couponRate,
        [ExcelArgument(Name = "years",      Description = BondConstants.ArgYears)]                object years,
        [ExcelArgument(Name = "frequency",  Description = BondConstants.ArgFrequencyDefault)]    object frequency)
        => Fn.Run(Category.Bonds, () => BondMath.YieldToMaturity(
               In.Price("price", price), In.Price("face", face),
               In.Rate("couponRate", couponRate), In.Years("years", years),
               In.PosInt("frequency", frequency, UserSettings.Current.DefaultFrequency)));

    [ExcelFunction(Name = BondConstants.DurationName, Category = Cat.Bonds, IsThreadSafe = true,
        Description = BondConstants.DurationDesc)]
    public static object BondDuration(
        [ExcelArgument(Name = "face",       Description = Arg.Face)]                          object face,
        [ExcelArgument(Name = "couponRate", Description = BondConstants.ArgCouponRatePlain)]   object couponRate,
        [ExcelArgument(Name = "ytm",        Description = BondConstants.ArgYtm)]              object ytm,
        [ExcelArgument(Name = "years",      Description = BondConstants.ArgYears)]            object years,
        [ExcelArgument(Name = "frequency",  Description = BondConstants.ArgFrequencyPlain)]   object frequency)
        => Fn.Run(Category.Bonds, () => BondMath.MacaulayDuration(
               In.Price("face", face), In.Rate("couponRate", couponRate),
               In.Rate("ytm", ytm), In.Years("years", years),
               In.PosInt("frequency", frequency, UserSettings.Current.DefaultFrequency)));

    [ExcelFunction(Name = BondConstants.ModDurationName, Category = Cat.Bonds, IsThreadSafe = true,
        Description = BondConstants.ModDurationDesc)]
    public static object BondModDuration(
        [ExcelArgument(Name = "face",       Description = Arg.Face)]                          object face,
        [ExcelArgument(Name = "couponRate", Description = BondConstants.ArgCouponRatePlain)]   object couponRate,
        [ExcelArgument(Name = "ytm",        Description = BondConstants.ArgYtm)]              object ytm,
        [ExcelArgument(Name = "years",      Description = BondConstants.ArgYears)]            object years,
        [ExcelArgument(Name = "frequency",  Description = BondConstants.ArgFrequencyPlain)]   object frequency)
        => Fn.Run(Category.Bonds, () => BondMath.ModifiedDuration(
               In.Price("face", face), In.Rate("couponRate", couponRate),
               In.Rate("ytm", ytm), In.Years("years", years),
               In.PosInt("frequency", frequency, UserSettings.Current.DefaultFrequency)));

    [ExcelFunction(Name = BondConstants.ConvexityName, Category = Cat.Bonds, IsThreadSafe = true,
        Description = BondConstants.ConvexityDesc)]
    public static object BondConvexity(
        [ExcelArgument(Name = "face",       Description = Arg.Face)]                          object face,
        [ExcelArgument(Name = "couponRate", Description = BondConstants.ArgCouponRatePlain)]   object couponRate,
        [ExcelArgument(Name = "ytm",        Description = BondConstants.ArgYtm)]              object ytm,
        [ExcelArgument(Name = "years",      Description = BondConstants.ArgYears)]            object years,
        [ExcelArgument(Name = "frequency",  Description = BondConstants.ArgFrequencyPlain)]   object frequency)
        => Fn.Run(Category.Bonds, () => BondMath.Convexity(
               In.Price("face", face), In.Rate("couponRate", couponRate),
               In.Rate("ytm", ytm), In.Years("years", years),
               In.PosInt("frequency", frequency, UserSettings.Current.DefaultFrequency)));

    [ExcelFunction(Name = BondConstants.Dv01Name, Category = Cat.Bonds, IsThreadSafe = true,
        Description = BondConstants.Dv01Desc)]
    public static object BondDv01(
        [ExcelArgument(Name = "face",       Description = Arg.Face)]                          object face,
        [ExcelArgument(Name = "couponRate", Description = BondConstants.ArgCouponRatePlain)]   object couponRate,
        [ExcelArgument(Name = "ytm",        Description = BondConstants.ArgYtm)]              object ytm,
        [ExcelArgument(Name = "years",      Description = BondConstants.ArgYears)]            object years,
        [ExcelArgument(Name = "frequency",  Description = BondConstants.ArgFrequencyPlain)]   object frequency)
        => Fn.Run(Category.Bonds, () => BondMath.DV01(
               In.Price("face", face), In.Rate("couponRate", couponRate),
               In.Rate("ytm", ytm), In.Years("years", years),
               In.PosInt("frequency", frequency, UserSettings.Current.DefaultFrequency)));

    [ExcelFunction(Name = BondConstants.PriceChangeName, Category = Cat.Bonds, IsThreadSafe = true,
        Description = BondConstants.PriceChangeDesc)]
    public static object BondPriceChange(
        [ExcelArgument(Name = "face",       Description = Arg.Face)]                          object face,
        [ExcelArgument(Name = "couponRate", Description = BondConstants.ArgCouponRatePlain)]   object couponRate,
        [ExcelArgument(Name = "ytm",        Description = BondConstants.ArgCurrentYtm)]       object ytm,
        [ExcelArgument(Name = "years",      Description = BondConstants.ArgYears)]            object years,
        [ExcelArgument(Name = "deltaYtm",   Description = BondConstants.ArgDeltaYtm)]         object deltaYtm,
        [ExcelArgument(Name = "frequency",  Description = BondConstants.ArgFrequencyPlain)]   object frequency)
        => Fn.Run(Category.Bonds, () => BondMath.ApproximatePriceChange(
               In.Price("face", face), In.Rate("couponRate", couponRate),
               In.Rate("ytm", ytm), In.Years("years", years), In.Rate("deltaYtm", deltaYtm),
               In.PosInt("frequency", frequency, UserSettings.Current.DefaultFrequency)));

    // ── YieldCurve ────────────────────────────────────────────────────────────

    [ExcelFunction(Name = BondConstants.YcDfName, Category = Cat.Bonds, IsThreadSafe = true,
        Description = BondConstants.YcDfDesc)]
    public static object YcDf(
        [ExcelArgument(Name = "zeroRate", Description = BondConstants.ArgZeroRate)]      object zeroRate,
        [ExcelArgument(Name = "T",        Description = BondConstants.ArgMaturityYears)] object t)
        => Fn.Run(Category.Bonds, () => YieldCurve.DiscountFactor(
               In.Rate("zeroRate", zeroRate), In.Years("T", t)));

    [ExcelFunction(Name = BondConstants.YcToContName, Category = Cat.Bonds, IsThreadSafe = true,
        Description = BondConstants.YcToContDesc)]
    public static object YcToCont(
        [ExcelArgument(Name = "rate",      Description = BondConstants.ArgRatePeriodic)]       object rate,
        [ExcelArgument(Name = "frequency", Description = BondConstants.ArgFreqToContDefault)]  object frequency)
        => Fn.Run(Category.Bonds, () => YieldCurve.ToContinuous(
               In.Rate("rate", rate),
               In.PosInt("frequency", frequency, UserSettings.Current.DefaultFrequency)));

    [ExcelFunction(Name = BondConstants.YcFromContName, Category = Cat.Bonds, IsThreadSafe = true,
        Description = BondConstants.YcFromContDesc)]
    public static object YcFromCont(
        [ExcelArgument(Name = "rate",      Description = BondConstants.ArgRateContinuous)] object rate,
        [ExcelArgument(Name = "frequency", Description = BondConstants.ArgFreqTarget)]     object frequency)
        => Fn.Run(Category.Bonds, () => YieldCurve.FromContinuous(
               In.Rate("rate", rate),
               In.PosInt("frequency", frequency, UserSettings.Current.DefaultFrequency)));

    [ExcelFunction(Name = BondConstants.YcFwdRateName, Category = Cat.Bonds, IsThreadSafe = true,
        Description = BondConstants.YcFwdRateDesc)]
    public static object YcFwdRate(
        [ExcelArgument(Name = "r1", Description = BondConstants.ArgR1)] object r1,
        [ExcelArgument(Name = "t1", Description = BondConstants.ArgT1)] object t1,
        [ExcelArgument(Name = "r2", Description = BondConstants.ArgR2)] object r2,
        [ExcelArgument(Name = "t2", Description = BondConstants.ArgT2)] object t2)
        => Fn.Run(Category.Bonds, () => YieldCurve.ForwardRate(
               In.Rate("r1", r1), In.Years("t1", t1),
               In.Rate("r2", r2), In.Years("t2", t2)));

    [ExcelFunction(Name = BondConstants.YcInterpolateName, Category = Cat.Bonds, IsThreadSafe = true,
        Description = BondConstants.YcInterpolateDesc)]
    public static object YcInterpolate(
        [ExcelArgument(Name = "maturities", Description = BondConstants.ArgMaturities)]      object maturities,
        [ExcelArgument(Name = "zeroRates",  Description = BondConstants.ArgZeroRates)]       object zeroRates,
        [ExcelArgument(Name = "T",          Description = BondConstants.ArgTargetMaturityT)] object t)
        => Fn.Run(Category.Bonds, () => YieldCurve.InterpolateZeroRate(
               In.Vector("maturities", maturities),
               In.Vector("zeroRates", zeroRates),
               In.Years("T", t)));

    [ExcelFunction(Name = BondConstants.YcParYieldName, Category = Cat.Bonds, IsThreadSafe = true,
        Description = BondConstants.YcParYieldDesc)]
    public static object YcParYield(
        [ExcelArgument(Name = "maturities",     Description = BondConstants.ArgMaturities)]        object maturities,
        [ExcelArgument(Name = "zeroRates",      Description = BondConstants.ArgZeroRates)]         object zeroRates,
        [ExcelArgument(Name = "targetMaturity", Description = BondConstants.ArgTargetMaturityPar)] object targetMaturity,
        [ExcelArgument(Name = "frequency",      Description = BondConstants.ArgFreqCouponDefault)] object frequency)
        => Fn.Run(Category.Bonds, () => YieldCurve.ParYield(
               In.Vector("maturities", maturities),
               In.Vector("zeroRates", zeroRates),
               In.Years("targetMaturity", targetMaturity),
               In.PosInt("frequency", frequency, UserSettings.Current.DefaultFrequency)));

    // ── MortgageMath ──────────────────────────────────────────────────────────

    [ExcelFunction(Name = BondConstants.MortPaymentName, Category = Cat.Bonds, IsThreadSafe = true,
        Description = BondConstants.MortPaymentDesc)]
    public static object MortPayment(
        [ExcelArgument(Name = "principal",       Description = BondConstants.ArgPrincipal)]       object principal,
        [ExcelArgument(Name = "annualRate",      Description = BondConstants.ArgAnnualRate)]      object annualRate,
        [ExcelArgument(Name = "years",           Description = BondConstants.ArgLoanTermYears)]   object years,
        [ExcelArgument(Name = "paymentsPerYear", Description = BondConstants.ArgPaymentsPerYear)] object paymentsPerYear)
        => Fn.Run(Category.Bonds, () => MortgageMath.Payment(
               In.Price("principal", principal), In.Rate("annualRate", annualRate),
               In.Years("years", years),
               In.PosInt("paymentsPerYear", paymentsPerYear, BondConstants.DefaultPaymentsPerYear)));

    [ExcelFunction(Name = BondConstants.MortBalanceName, Category = Cat.Bonds, IsThreadSafe = true,
        Description = BondConstants.MortBalanceDesc)]
    public static object MortBalance(
        [ExcelArgument(Name = "principal",       Description = BondConstants.ArgPrincipalOrig)]      object principal,
        [ExcelArgument(Name = "annualRate",      Description = BondConstants.ArgAnnualRate)]         object annualRate,
        [ExcelArgument(Name = "years",           Description = BondConstants.ArgLoanTermTotalYears)] object years,
        [ExcelArgument(Name = "paymentsMade",    Description = BondConstants.ArgPaymentsMade)]       object paymentsMade,
        [ExcelArgument(Name = "paymentsPerYear", Description = BondConstants.ArgPaymentsPerYear)]    object paymentsPerYear)
        => Fn.Run(Category.Bonds, () => MortgageMath.OutstandingBalance(
               In.Price("principal", principal), In.Rate("annualRate", annualRate),
               In.Years("years", years), In.Count("paymentsMade", paymentsMade),
               In.PosInt("paymentsPerYear", paymentsPerYear, BondConstants.DefaultPaymentsPerYear)));

    [ExcelFunction(Name = BondConstants.MortTotalInterestName, Category = Cat.Bonds, IsThreadSafe = true,
        Description = BondConstants.MortTotalInterestDesc)]
    public static object MortTotalInterest(
        [ExcelArgument(Name = "principal",       Description = BondConstants.ArgPrincipal)]       object principal,
        [ExcelArgument(Name = "annualRate",      Description = BondConstants.ArgAnnualRate)]      object annualRate,
        [ExcelArgument(Name = "years",           Description = BondConstants.ArgLoanTermYears)]   object years,
        [ExcelArgument(Name = "paymentsPerYear", Description = BondConstants.ArgPaymentsPerYear)] object paymentsPerYear)
        => Fn.Run(Category.Bonds, () => MortgageMath.TotalInterest(
               In.Price("principal", principal), In.Rate("annualRate", annualRate),
               In.Years("years", years),
               In.PosInt("paymentsPerYear", paymentsPerYear, BondConstants.DefaultPaymentsPerYear)));

    [ExcelFunction(Name = BondConstants.MortEarName, Category = Cat.Bonds, IsThreadSafe = true,
        Description = BondConstants.MortEarDesc)]
    public static object MortEar(
        [ExcelArgument(Name = "nominalRate", Description = BondConstants.ArgNominalRate)]     object nominalRate,
        [ExcelArgument(Name = "frequency",   Description = BondConstants.ArgCompoundingFreq)] object frequency)
        => Fn.Run(Category.Bonds, () => MortgageMath.EffectiveAnnualRate(
               In.Rate("nominalRate", nominalRate),
               In.PosInt("frequency", frequency)));
}
