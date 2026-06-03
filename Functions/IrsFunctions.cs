using ExcelDna.Integration;
using Aleksej.Finance.Derivatives;
using Aleksej.Finance.Excel.Constants;
using Aleksej.Finance.Excel.Helpers;
using Aleksej.Finance.Excel.Settings;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>Vanilla fixed-for-floating interest rate swap valuation (Hull Ch. 7).</summary>
public static class IrsFunctions
{
    [ExcelFunction(Name = IrsConstants.ValueName, Category = Cat.Derivatives, IsThreadSafe = true,
        Description = IrsConstants.ValueDesc, HelpTopic = IrsConstants.Help)]
    public static object IrsValue(
        [ExcelArgument(Name = "notional",          Description = IrsConstants.Notional)]        object notional,
        [ExcelArgument(Name = "fixedRate",         Description = IrsConstants.FixedRate)]       object fixedRate,
        [ExcelArgument(Name = "paymentTimes",      Description = IrsConstants.PaymentTimes)]    object paymentTimes,
        [ExcelArgument(Name = "zeroRates",         Description = IrsConstants.ZeroRates)]       object zeroRates,
        [ExcelArgument(Name = "nextFloatCoupon",   Description = IrsConstants.NextFloatCoupon)] object nextFloatCoupon,
        [ExcelArgument(Name = "timeToNextReset",   Description = IrsConstants.TimeToNextReset)] object timeToNextReset,
        [ExcelArgument(Name = "zeroAtNextReset",   Description = IrsConstants.ZeroAtNextReset)] object zeroAtNextReset,
        [ExcelArgument(Name = "isPayFixed",        Description = IrsConstants.IsPayFixed)]      object isPayFixed)
        => Fn.Run(Category.Derivatives, () => InterestRateSwap.SwapValue(
               In.Price("notional", notional), In.Rate("fixedRate", fixedRate),
               In.Vector("paymentTimes", paymentTimes), In.Vector("zeroRates", zeroRates),
               In.Rate("nextFloatCoupon", nextFloatCoupon), In.Years("timeToNextReset", timeToNextReset),
               In.Rate("zeroAtNextReset", zeroAtNextReset),
               In.Flag("isPayFixed", isPayFixed, IrsConstants.IsPayFixedDefault)));

    [ExcelFunction(Name = IrsConstants.ParRateName, Category = Cat.Derivatives, IsThreadSafe = true,
        Description = IrsConstants.ParRateDesc, HelpTopic = IrsConstants.Help)]
    public static object IrsParRate(
        [ExcelArgument(Name = "paymentTimes", Description = IrsConstants.PaymentTimesAsc)] object paymentTimes,
        [ExcelArgument(Name = "zeroRates",    Description = IrsConstants.ZeroRates)]       object zeroRates,
        [ExcelArgument(Name = "frequency",    Description = IrsConstants.Frequency)]       object frequency)
        => Fn.Run(Category.Derivatives, () => InterestRateSwap.ParSwapRate(
               In.Vector("paymentTimes", paymentTimes), In.Vector("zeroRates", zeroRates),
               In.PosInt("frequency", frequency, UserSettings.Current.DefaultFrequency)));

    [ExcelFunction(Name = IrsConstants.FixedLegName, Category = Cat.Derivatives, IsThreadSafe = true,
        Description = IrsConstants.FixedLegDesc, HelpTopic = IrsConstants.Help)]
    public static object IrsFixedLeg(
        [ExcelArgument(Name = "notional",         Description = IrsConstants.Notional)]         object notional,
        [ExcelArgument(Name = "fixedRate",        Description = IrsConstants.FixedRateShort)]   object fixedRate,
        [ExcelArgument(Name = "paymentTimes",     Description = IrsConstants.PaymentTimesPlain)]object paymentTimes,
        [ExcelArgument(Name = "zeroRates",        Description = IrsConstants.ZeroRates)]        object zeroRates,
        [ExcelArgument(Name = "frequency",        Description = IrsConstants.Frequency)]        object frequency,
        [ExcelArgument(Name = "includePrincipal", Description = IrsConstants.IncludePrincipal)] object includePrincipal)
        => Fn.Run(Category.Derivatives, () => InterestRateSwap.FixedLegPV(
               In.Price("notional", notional), In.Rate("fixedRate", fixedRate),
               In.Vector("paymentTimes", paymentTimes), In.Vector("zeroRates", zeroRates),
               In.PosInt("frequency", frequency, UserSettings.Current.DefaultFrequency),
               In.Flag("includePrincipal", includePrincipal)));

    [ExcelFunction(Name = IrsConstants.FloatLegName, Category = Cat.Derivatives, IsThreadSafe = true,
        Description = IrsConstants.FloatLegDesc, HelpTopic = IrsConstants.Help)]
    public static object IrsFloatLeg(
        [ExcelArgument(Name = "notional",         Description = IrsConstants.Notional)]               object notional,
        [ExcelArgument(Name = "nextFloatCoupon",  Description = IrsConstants.NextFloatCouponPeriod)]  object nextFloatCoupon,
        [ExcelArgument(Name = "timeToNextReset",  Description = IrsConstants.TimeToNextResetYears)]   object timeToNextReset,
        [ExcelArgument(Name = "zeroAtNextReset",  Description = IrsConstants.ZeroAtNextReset)]        object zeroAtNextReset)
        => Fn.Run(Category.Derivatives, () => InterestRateSwap.FloatingLegPV(
               In.Price("notional", notional), In.Rate("nextFloatCoupon", nextFloatCoupon),
               In.Years("timeToNextReset", timeToNextReset), In.Rate("zeroAtNextReset", zeroAtNextReset)));

    [ExcelFunction(Name = IrsConstants.Dv01Name, Category = Cat.Derivatives, IsThreadSafe = true,
        Description = IrsConstants.Dv01Desc, HelpTopic = IrsConstants.Help)]
    public static object IrsDv01(
        [ExcelArgument(Name = "notional",        Description = IrsConstants.Notional)]              object notional,
        [ExcelArgument(Name = "fixedRate",       Description = IrsConstants.FixedRateShort)]        object fixedRate,
        [ExcelArgument(Name = "paymentTimes",    Description = IrsConstants.PaymentTimesPlain)]     object paymentTimes,
        [ExcelArgument(Name = "zeroRates",       Description = IrsConstants.ZeroRates)]             object zeroRates,
        [ExcelArgument(Name = "nextFloatCoupon", Description = IrsConstants.NextFloatCouponShort)]  object nextFloatCoupon,
        [ExcelArgument(Name = "timeToNextReset", Description = IrsConstants.TimeToNextResetShort)]  object timeToNextReset,
        [ExcelArgument(Name = "zeroAtNextReset", Description = IrsConstants.ZeroAtNextResetShort)]  object zeroAtNextReset,
        [ExcelArgument(Name = "isPayFixed",      Description = IrsConstants.IsPayFixedShort)]       object isPayFixed)
        => Fn.Run(Category.Derivatives, () => InterestRateSwap.DV01(
               In.Price("notional", notional), In.Rate("fixedRate", fixedRate),
               In.Vector("paymentTimes", paymentTimes), In.Vector("zeroRates", zeroRates),
               In.Rate("nextFloatCoupon", nextFloatCoupon), In.Years("timeToNextReset", timeToNextReset),
               In.Rate("zeroAtNextReset", zeroAtNextReset),
               In.Flag("isPayFixed", isPayFixed, IrsConstants.IsPayFixedDefault)));
}
