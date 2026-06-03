using ExcelDna.Integration;
using Aleksej.Finance.Derivatives;
using Aleksej.Finance.Excel.Constants;
using Aleksej.Finance.Excel.Helpers;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>Forward Rate Agreement (FRA) pricing and settlement (Hull Ch. 4).</summary>
public static class FraFunctions
{
    [ExcelFunction(Name = FraConstants.RateName, Category = Cat.Derivatives, IsThreadSafe = true,
        Description = FraConstants.RateDesc, HelpTopic = FraConstants.Help)]
    public static object FraRate(
        [ExcelArgument(Name = "r1", Description = FraConstants.R1Cont)]  object r1,
        [ExcelArgument(Name = "t1", Description = FraConstants.T1Start)] object t1,
        [ExcelArgument(Name = "r2", Description = FraConstants.R2Cont)]  object r2,
        [ExcelArgument(Name = "t2", Description = FraConstants.T2End)]   object t2)
        => Fn.Run(Category.Derivatives, () => ForwardRateAgreement.ForwardRate(
               In.Rate("r1", r1), In.Years("t1", t1), In.Rate("r2", r2), In.Years("t2", t2)));

    [ExcelFunction(Name = FraConstants.RateSimpleName, Category = Cat.Derivatives, IsThreadSafe = true,
        Description = FraConstants.RateSimpleDesc, HelpTopic = FraConstants.Help)]
    public static object FraRateSimple(
        [ExcelArgument(Name = "r1", Description = FraConstants.R1)]           object r1,
        [ExcelArgument(Name = "t1", Description = FraConstants.T1StartShort)] object t1,
        [ExcelArgument(Name = "r2", Description = FraConstants.R2)]           object r2,
        [ExcelArgument(Name = "t2", Description = FraConstants.T2EndShort)]   object t2)
        => Fn.Run(Category.Derivatives, () => ForwardRateAgreement.ForwardRateSimple(
               In.Rate("r1", r1), In.Years("t1", t1), In.Rate("r2", r2), In.Years("t2", t2)));

    [ExcelFunction(Name = FraConstants.ValueName, Category = Cat.Derivatives, IsThreadSafe = true,
        Description = FraConstants.ValueDesc, HelpTopic = FraConstants.Help)]
    public static object FraValue(
        [ExcelArgument(Name = "notional", Description = FraConstants.Notional)]  object notional,
        [ExcelArgument(Name = "fraRate",  Description = FraConstants.FraRateK)]  object fraRate,
        [ExcelArgument(Name = "r1",       Description = FraConstants.R1Current)] object r1,
        [ExcelArgument(Name = "t1",       Description = FraConstants.T1Accrual)] object t1,
        [ExcelArgument(Name = "r2",       Description = FraConstants.R2Current)] object r2,
        [ExcelArgument(Name = "t2",       Description = FraConstants.T2Accrual)] object t2,
        [ExcelArgument(Name = "isLong",   Description = FraConstants.IsLong)]    object isLong)
        => Fn.Run(Category.Derivatives, () => ForwardRateAgreement.FraValue(
               In.Price("notional", notional), In.Rate("fraRate", fraRate),
               In.Rate("r1", r1), In.Years("t1", t1),
               In.Rate("r2", r2), In.Years("t2", t2),
               In.Flag("isLong", isLong, FraConstants.IsLongDefault)));

    [ExcelFunction(Name = FraConstants.SettlementName, Category = Cat.Derivatives, IsThreadSafe = true,
        Description = FraConstants.SettlementDesc, HelpTopic = FraConstants.Help)]
    public static object FraSettlement(
        [ExcelArgument(Name = "notional",   Description = FraConstants.Notional)]       object notional,
        [ExcelArgument(Name = "fraRate",    Description = FraConstants.FraRateK)]       object fraRate,
        [ExcelArgument(Name = "marketRate", Description = FraConstants.MarketRate)]     object marketRate,
        [ExcelArgument(Name = "t1",         Description = FraConstants.T1AccrualShort)] object t1,
        [ExcelArgument(Name = "t2",         Description = FraConstants.T2AccrualShort)] object t2,
        [ExcelArgument(Name = "isLong",     Description = FraConstants.IsLongShort)]    object isLong)
        => Fn.Run(Category.Derivatives, () => ForwardRateAgreement.FraSettlement(
               In.Price("notional", notional), In.Rate("fraRate", fraRate),
               In.Rate("marketRate", marketRate),
               In.Years("t1", t1), In.Years("t2", t2),
               In.Flag("isLong", isLong, FraConstants.IsLongDefault)));

    [ExcelFunction(Name = FraConstants.Dv01Name, Category = Cat.Derivatives, IsThreadSafe = true,
        Description = FraConstants.Dv01Desc, HelpTopic = FraConstants.Help)]
    public static object FraDv01(
        [ExcelArgument(Name = "notional", Description = FraConstants.Notional)]       object notional,
        [ExcelArgument(Name = "fraRate",  Description = FraConstants.FraRateAgreed)]  object fraRate,
        [ExcelArgument(Name = "r1",       Description = FraConstants.R1)]             object r1,
        [ExcelArgument(Name = "t1",       Description = FraConstants.T1AccrualShort)] object t1,
        [ExcelArgument(Name = "r2",       Description = FraConstants.R2)]             object r2,
        [ExcelArgument(Name = "t2",       Description = FraConstants.T2AccrualShort)] object t2,
        [ExcelArgument(Name = "isLong",   Description = FraConstants.IsLongShort)]    object isLong)
        => Fn.Run(Category.Derivatives, () => ForwardRateAgreement.DV01(
               In.Price("notional", notional), In.Rate("fraRate", fraRate),
               In.Rate("r1", r1), In.Years("t1", t1),
               In.Rate("r2", r2), In.Years("t2", t2),
               In.Flag("isLong", isLong, FraConstants.IsLongDefault)));
}
