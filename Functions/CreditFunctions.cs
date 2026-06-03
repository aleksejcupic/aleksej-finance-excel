using ExcelDna.Integration;
using Aleksej.Finance.Credit;
using Aleksej.Finance.Excel.Constants;
using Aleksej.Finance.Excel.Helpers;
using Aleksej.Finance.Excel.Settings;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>Credit risk: Merton structural model and CDS pricing via hazard rates (Hull Ch. 24-25).</summary>
public static class CreditFunctions
{
    [ExcelFunction(Name = CreditConstants.MertonEquityName, Category = Cat.Credit, IsThreadSafe = true,
        Description = CreditConstants.MertonEquityDesc, HelpTopic = CreditConstants.Help)]
    public static object CrMertonEquity(
        [ExcelArgument(Name = "V",      Description = CreditConstants.V)]      object v,
        [ExcelArgument(Name = "D",      Description = CreditConstants.D)]      object d,
        [ExcelArgument(Name = "T",      Description = CreditConstants.TMaturity)] object t,
        [ExcelArgument(Name = "r",      Description = CreditConstants.R)]      object r,
        [ExcelArgument(Name = "sigmaV", Description = CreditConstants.SigmaV)] object sigmaV)
        => Fn.Run(Category.Credit, () => CreditDerivatives.MertonEquityValue(
               In.Price("V", v), In.Price("D", d), In.Years("T", t),
               In.Rate("r", r), In.Vol("sigmaV", sigmaV)));

    [ExcelFunction(Name = CreditConstants.MertonDebtName, Category = Cat.Credit, IsThreadSafe = true,
        Description = CreditConstants.MertonDebtDesc, HelpTopic = CreditConstants.Help)]
    public static object CrMertonDebt(
        [ExcelArgument(Name = "V",      Description = CreditConstants.VShort)]      object v,
        [ExcelArgument(Name = "D",      Description = CreditConstants.DShort)]      object d,
        [ExcelArgument(Name = "T",      Description = CreditConstants.TMaturityShort)] object t,
        [ExcelArgument(Name = "r",      Description = CreditConstants.RShort)]      object r,
        [ExcelArgument(Name = "sigmaV", Description = CreditConstants.SigmaVShort)] object sigmaV)
        => Fn.Run(Category.Credit, () => CreditDerivatives.MertonDebtValue(
               In.Price("V", v), In.Price("D", d), In.Years("T", t),
               In.Rate("r", r), In.Vol("sigmaV", sigmaV)));

    [ExcelFunction(Name = CreditConstants.DefaultProbName, Category = Cat.Credit, IsThreadSafe = true,
        Description = CreditConstants.DefaultProbDesc, HelpTopic = CreditConstants.Help)]
    public static object CrDefaultProb(
        [ExcelArgument(Name = "V",      Description = CreditConstants.VShort)]      object v,
        [ExcelArgument(Name = "D",      Description = CreditConstants.DShort)]      object d,
        [ExcelArgument(Name = "T",      Description = CreditConstants.TMaturityShort)] object t,
        [ExcelArgument(Name = "r",      Description = CreditConstants.RShort)]      object r,
        [ExcelArgument(Name = "sigmaV", Description = CreditConstants.SigmaVShort)] object sigmaV)
        => Fn.Run(Category.Credit, () => CreditDerivatives.MertonDefaultProbability(
               In.Price("V", v), In.Price("D", d), In.Years("T", t),
               In.Rate("r", r), In.Vol("sigmaV", sigmaV)));

    [ExcelFunction(Name = CreditConstants.DistToDefaultName, Category = Cat.Credit, IsThreadSafe = true,
        Description = CreditConstants.DistToDefaultDesc, HelpTopic = CreditConstants.Help)]
    public static object CrDistToDefault(
        [ExcelArgument(Name = "V",      Description = CreditConstants.VShort)]      object v,
        [ExcelArgument(Name = "D",      Description = CreditConstants.DShort)]      object d,
        [ExcelArgument(Name = "T",      Description = CreditConstants.TMaturityShort)] object t,
        [ExcelArgument(Name = "r",      Description = CreditConstants.RShort)]      object r,
        [ExcelArgument(Name = "sigmaV", Description = CreditConstants.SigmaVShort)] object sigmaV)
        => Fn.Run(Category.Credit, () => CreditDerivatives.DistanceToDefault(
               In.Price("V", v), In.Price("D", d), In.Years("T", t),
               In.Rate("r", r), In.Vol("sigmaV", sigmaV)));

    [ExcelFunction(Name = CreditConstants.CreditSpreadName, Category = Cat.Credit, IsThreadSafe = true,
        Description = CreditConstants.CreditSpreadDesc, HelpTopic = CreditConstants.Help)]
    public static object CrCreditSpread(
        [ExcelArgument(Name = "V",      Description = CreditConstants.VShort)]      object v,
        [ExcelArgument(Name = "D",      Description = CreditConstants.DShort)]      object d,
        [ExcelArgument(Name = "T",      Description = CreditConstants.TMaturityShort)] object t,
        [ExcelArgument(Name = "r",      Description = CreditConstants.RShort)]      object r,
        [ExcelArgument(Name = "sigmaV", Description = CreditConstants.SigmaVShort)] object sigmaV)
        => Fn.Run(Category.Credit, () => CreditDerivatives.MertonCreditSpread(
               In.Price("V", v), In.Price("D", d), In.Years("T", t),
               In.Rate("r", r), In.Vol("sigmaV", sigmaV)));

    [ExcelFunction(Name = CreditConstants.SurvivalProbName, Category = Cat.Credit, IsThreadSafe = true,
        Description = CreditConstants.SurvivalProbDesc, HelpTopic = CreditConstants.Help)]
    public static object CrSurvivalProb(
        [ExcelArgument(Name = "hazardRate", Description = CreditConstants.HazardRate)] object hazardRate,
        [ExcelArgument(Name = "T",          Description = CreditConstants.THorizon)]    object t)
        => Fn.Run(Category.Credit, () => CreditDerivatives.SurvivalProbability(
               In.Num("hazardRate", hazardRate), In.Years("T", t)));

    [ExcelFunction(Name = CreditConstants.HazardFromSpreadName, Category = Cat.Credit, IsThreadSafe = true,
        Description = CreditConstants.HazardFromSpreadDesc, HelpTopic = CreditConstants.Help)]
    public static object CrHazardFromSpread(
        [ExcelArgument(Name = "creditSpread",  Description = CreditConstants.CreditSpread)]     object creditSpread,
        [ExcelArgument(Name = "recoveryRate",  Description = CreditConstants.RecoverySettings)] object recoveryRate)
        => Fn.Run(Category.Credit, () => CreditDerivatives.HazardRateFromSpread(
               In.Rate("creditSpread", creditSpread),
               In.Prob("recoveryRate", recoveryRate, UserSettings.Current.DefaultRecoveryRate)));

    [ExcelFunction(Name = CreditConstants.CdsSpreadName, Category = Cat.Credit, IsThreadSafe = true,
        Description = CreditConstants.CdsSpreadDesc, HelpTopic = CreditConstants.Help)]
    public static object CrCdsSpread(
        [ExcelArgument(Name = "hazardRate",   Description = CreditConstants.HazardRate)] object hazardRate,
        [ExcelArgument(Name = "r",            Description = CreditConstants.RFlat)]      object r,
        [ExcelArgument(Name = "maturity",     Description = CreditConstants.CdsMaturity)] object maturity,
        [ExcelArgument(Name = "recoveryRate", Description = CreditConstants.Recovery)]   object recoveryRate,
        [ExcelArgument(Name = "frequency",    Description = CreditConstants.Frequency)]  object frequency)
        => Fn.Run(Category.Credit, () => CreditDerivatives.CdsFairSpread(
               In.Num("hazardRate", hazardRate), In.Rate("r", r), In.Years("maturity", maturity),
               In.Prob("recoveryRate", recoveryRate, UserSettings.Current.DefaultRecoveryRate),
               In.PosInt("frequency", frequency, CreditConstants.DefaultFrequency)));

    [ExcelFunction(Name = CreditConstants.CdsMtmName, Category = Cat.Credit, IsThreadSafe = true,
        Description = CreditConstants.CdsMtmDesc, HelpTopic = CreditConstants.Help)]
    public static object CrCdsMtm(
        [ExcelArgument(Name = "contractedSpread",  Description = CreditConstants.ContractedSpread)]  object contractedSpread,
        [ExcelArgument(Name = "hazardRate",        Description = CreditConstants.CurrentHazardRate)] object hazardRate,
        [ExcelArgument(Name = "r",                 Description = CreditConstants.CurrentR)]           object r,
        [ExcelArgument(Name = "notional",          Description = CreditConstants.Notional)]          object notional,
        [ExcelArgument(Name = "remainingMaturity", Description = CreditConstants.RemainingMaturity)] object remainingMaturity,
        [ExcelArgument(Name = "recoveryRate",      Description = CreditConstants.Recovery)]          object recoveryRate,
        [ExcelArgument(Name = "frequency",         Description = CreditConstants.FrequencyShort)]    object frequency)
        => Fn.Run(Category.Credit, () => CreditDerivatives.CdsMtm(
               In.Rate("contractedSpread", contractedSpread), In.Num("hazardRate", hazardRate), In.Rate("r", r),
               In.Price("notional", notional), In.Years("remainingMaturity", remainingMaturity),
               In.Prob("recoveryRate", recoveryRate, UserSettings.Current.DefaultRecoveryRate),
               In.PosInt("frequency", frequency, CreditConstants.DefaultFrequency)));
}
