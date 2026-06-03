using ExcelDna.Integration;
using Aleksej.Finance.Derivatives;
using Aleksej.Finance.Excel.Constants;
using Aleksej.Finance.Excel.Helpers;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>Black's (1976) model for interest rate caps, floors, and swaptions (Hull Ch. 29).</summary>
public static class BlackModelFunctions
{
    [ExcelFunction(Name = BlackModelConstants.CapletName, Category = Cat.Derivatives, IsThreadSafe = true,
        Description = BlackModelConstants.CapletDesc, HelpTopic = BlackModelConstants.Help)]
    public static object BmCaplet(
        [ExcelArgument(Name = "notional",        Description = BlackModelConstants.Notional)]        object notional,
        [ExcelArgument(Name = "forwardRate",     Description = BlackModelConstants.ForwardRate)]     object forwardRate,
        [ExcelArgument(Name = "strike",          Description = BlackModelConstants.StrikeCapFloor)]  object strike,
        [ExcelArgument(Name = "T",               Description = BlackModelConstants.TCaplet)]         object t,
        [ExcelArgument(Name = "r",               Description = BlackModelConstants.RZero)]           object r,
        [ExcelArgument(Name = "sigma",           Description = BlackModelConstants.SigmaForward)]    object sigma,
        [ExcelArgument(Name = "accrualFraction", Description = BlackModelConstants.AccrualFraction)] object accrualFraction,
        [ExcelArgument(Name = "isFloor",         Description = BlackModelConstants.IsFloor)]         object isFloor)
        => Fn.Run(Category.Derivatives, () => BlackModel.CapletPrice(
               In.Price("notional", notional), In.Rate("forwardRate", forwardRate), In.Rate("strike", strike),
               In.Years("T", t), In.Rate("r", r), In.Vol("sigma", sigma),
               RangeHelper.IsMissing(accrualFraction)
                   ? BlackModelConstants.AccrualFractionDefault
                   : In.Years("accrualFraction", accrualFraction),
               In.Flag("isFloor", isFloor)));

    [ExcelFunction(Name = BlackModelConstants.CapName, Category = Cat.Derivatives, IsThreadSafe = true,
        Description = BlackModelConstants.CapDesc, HelpTopic = BlackModelConstants.Help)]
    public static object BmCap(
        [ExcelArgument(Name = "notional",     Description = BlackModelConstants.Notional)]            object notional,
        [ExcelArgument(Name = "strike",       Description = BlackModelConstants.StrikeCap)]           object strike,
        [ExcelArgument(Name = "sigma",        Description = BlackModelConstants.SigmaFlat)]           object sigma,
        [ExcelArgument(Name = "paymentTimes", Description = BlackModelConstants.PaymentTimesReset)]   object paymentTimes,
        [ExcelArgument(Name = "zeroRates",    Description = BlackModelConstants.ZeroRates)]           object zeroRates,
        [ExcelArgument(Name = "forwardRates", Description = BlackModelConstants.ForwardRates)]        object forwardRates,
        [ExcelArgument(Name = "accrualFracs", Description = BlackModelConstants.AccrualFracs)]        object accrualFracs)
        => Fn.Run(Category.Derivatives, () => BlackModel.CapPrice(
               In.Price("notional", notional), In.Rate("strike", strike), In.Vol("sigma", sigma),
               In.Vector("paymentTimes", paymentTimes), In.Vector("zeroRates", zeroRates),
               In.Vector("forwardRates", forwardRates), In.Vector("accrualFracs", accrualFracs)));

    [ExcelFunction(Name = BlackModelConstants.FloorName, Category = Cat.Derivatives, IsThreadSafe = true,
        Description = BlackModelConstants.FloorDesc, HelpTopic = BlackModelConstants.Help)]
    public static object BmFloor(
        [ExcelArgument(Name = "notional",     Description = BlackModelConstants.Notional)]              object notional,
        [ExcelArgument(Name = "strike",       Description = BlackModelConstants.StrikeFloor)]           object strike,
        [ExcelArgument(Name = "sigma",        Description = BlackModelConstants.SigmaFlat)]             object sigma,
        [ExcelArgument(Name = "paymentTimes", Description = BlackModelConstants.PaymentTimesResetShort)]object paymentTimes,
        [ExcelArgument(Name = "zeroRates",    Description = BlackModelConstants.ZeroRatesShort)]        object zeroRates,
        [ExcelArgument(Name = "forwardRates", Description = BlackModelConstants.ForwardRatesShort)]     object forwardRates,
        [ExcelArgument(Name = "accrualFracs", Description = BlackModelConstants.AccrualFracsShort)]     object accrualFracs)
        => Fn.Run(Category.Derivatives, () => BlackModel.FloorPrice(
               In.Price("notional", notional), In.Rate("strike", strike), In.Vol("sigma", sigma),
               In.Vector("paymentTimes", paymentTimes), In.Vector("zeroRates", zeroRates),
               In.Vector("forwardRates", forwardRates), In.Vector("accrualFracs", accrualFracs)));

    [ExcelFunction(Name = BlackModelConstants.FwdSwapRateName, Category = Cat.Derivatives, IsThreadSafe = true,
        Description = BlackModelConstants.FwdSwapRateDesc, HelpTopic = BlackModelConstants.Help)]
    public static object BmFwdSwapRate(
        [ExcelArgument(Name = "paymentTimes", Description = BlackModelConstants.PaymentTimesSwapExpiry)] object paymentTimes,
        [ExcelArgument(Name = "zeroRates",    Description = BlackModelConstants.ZeroRates)]              object zeroRates,
        [ExcelArgument(Name = "accrualFracs", Description = BlackModelConstants.AccrualFracsShort)]      object accrualFracs)
        => Fn.Run(Category.Derivatives, () => BlackModel.ForwardSwapRate(
               In.Vector("paymentTimes", paymentTimes), In.Vector("zeroRates", zeroRates),
               In.Vector("accrualFracs", accrualFracs)));

    [ExcelFunction(Name = BlackModelConstants.SwaptionName, Category = Cat.Derivatives, IsThreadSafe = true,
        Description = BlackModelConstants.SwaptionDesc, HelpTopic = BlackModelConstants.Help)]
    public static object BmSwaption(
        [ExcelArgument(Name = "notional",     Description = BlackModelConstants.Notional)]               object notional,
        [ExcelArgument(Name = "strike",       Description = BlackModelConstants.StrikeSwap)]             object strike,
        [ExcelArgument(Name = "T",            Description = BlackModelConstants.TSwaption)]              object t,
        [ExcelArgument(Name = "sigma",        Description = BlackModelConstants.SigmaSwap)]              object sigma,
        [ExcelArgument(Name = "paymentTimes", Description = BlackModelConstants.PaymentTimesSwapCoupon)] object paymentTimes,
        [ExcelArgument(Name = "zeroRates",    Description = BlackModelConstants.ZeroRates)]              object zeroRates,
        [ExcelArgument(Name = "accrualFracs", Description = BlackModelConstants.AccrualFracsShort)]      object accrualFracs,
        [ExcelArgument(Name = "isPayer",      Description = BlackModelConstants.IsPayer)]               object isPayer)
        => Fn.Run(Category.Derivatives, () => BlackModel.SwaptionPrice(
               In.Price("notional", notional), In.Rate("strike", strike), In.Years("T", t),
               In.Vol("sigma", sigma),
               In.Vector("paymentTimes", paymentTimes), In.Vector("zeroRates", zeroRates),
               In.Vector("accrualFracs", accrualFracs),
               In.Flag("isPayer", isPayer, BlackModelConstants.IsPayerDefault)));
}
