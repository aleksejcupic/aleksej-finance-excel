using ExcelDna.Integration;
using AleksejCupic.FinancialMath.Credit;
using Aleksej.Finance.Excel.Helpers;
using Aleksej.Finance.Excel.Settings;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>Credit risk: Merton structural model and CDS pricing via hazard rates (Hull Ch. 24-25).</summary>
public static class CreditFunctions
{
    private static bool Enabled => UserSettings.Load().EnableCredit;
    private static string Off   => RangeHelper.DisabledMessage("Credit");
    private static double RecoveryDefault => UserSettings.Load().DefaultRecoveryRate;

    [ExcelFunction(Name = "CR_MERTON_EQUITY", Category = "Finance | Credit", IsThreadSafe = true,
        Description = "Merton model equity value: equity is a call on firm assets V with strike = debt D. E = V*N(d1) - D*exp(-rT)*N(d2).",
        HelpTopic = "https://aleksejcupic.github.io/financial-math/credit/credit-derivatives")]
    public static object CrMertonEquity(
        [ExcelArgument(Name = "V",      Description = "Current total asset value of the firm")]          object v,
        [ExcelArgument(Name = "D",      Description = "Face value of debt (default boundary at maturity)")] object d,
        [ExcelArgument(Name = "T",      Description = "Debt maturity in years")]                          object t,
        [ExcelArgument(Name = "r",      Description = "Continuously compounded risk-free rate")]          object r,
        [ExcelArgument(Name = "sigmaV", Description = "Annualised volatility of firm asset value")]       object sigmaV)
        => Enabled ? CreditDerivatives.MertonEquityValue(
                         RangeHelper.Scalar(v), RangeHelper.Scalar(d), RangeHelper.Scalar(t),
                         RangeHelper.Scalar(r), RangeHelper.Scalar(sigmaV))
                   : (object)Off;

    [ExcelFunction(Name = "CR_MERTON_DEBT", Category = "Finance | Credit", IsThreadSafe = true,
        Description = "Merton model PV of debt. PV(Debt) = V - Equity = D*exp(-rT)*N(d2) + V*N(-d1).")]
    public static object CrMertonDebt(
        [ExcelArgument(Name = "V",      Description = "Firm asset value")]     object v,
        [ExcelArgument(Name = "D",      Description = "Face value of debt")]   object d,
        [ExcelArgument(Name = "T",      Description = "Debt maturity")]        object t,
        [ExcelArgument(Name = "r",      Description = "Risk-free rate")]       object r,
        [ExcelArgument(Name = "sigmaV", Description = "Asset value vol")]      object sigmaV)
        => Enabled ? CreditDerivatives.MertonDebtValue(
                         RangeHelper.Scalar(v), RangeHelper.Scalar(d), RangeHelper.Scalar(t),
                         RangeHelper.Scalar(r), RangeHelper.Scalar(sigmaV))
                   : (object)Off;

    [ExcelFunction(Name = "CR_DEFAULT_PROB", Category = "Finance | Credit", IsThreadSafe = true,
        Description = "Merton risk-neutral probability of default at maturity T. PD = N(-d2). Higher leverage/vol → higher PD.")]
    public static object CrDefaultProb(
        [ExcelArgument(Name = "V",      Description = "Firm asset value")]     object v,
        [ExcelArgument(Name = "D",      Description = "Face value of debt")]   object d,
        [ExcelArgument(Name = "T",      Description = "Debt maturity")]        object t,
        [ExcelArgument(Name = "r",      Description = "Risk-free rate")]       object r,
        [ExcelArgument(Name = "sigmaV", Description = "Asset value vol")]      object sigmaV)
        => Enabled ? CreditDerivatives.MertonDefaultProbability(
                         RangeHelper.Scalar(v), RangeHelper.Scalar(d), RangeHelper.Scalar(t),
                         RangeHelper.Scalar(r), RangeHelper.Scalar(sigmaV))
                   : (object)Off;

    [ExcelFunction(Name = "CR_DIST_TO_DEFAULT", Category = "Finance | Credit", IsThreadSafe = true,
        Description = "Distance to Default (DD) under Merton. DD = d2 = (ln(V/D)+(r-σ²/2)T)/(σ√T). Higher DD = safer firm.")]
    public static object CrDistToDefault(
        [ExcelArgument(Name = "V",      Description = "Firm asset value")]     object v,
        [ExcelArgument(Name = "D",      Description = "Face value of debt")]   object d,
        [ExcelArgument(Name = "T",      Description = "Debt maturity")]        object t,
        [ExcelArgument(Name = "r",      Description = "Risk-free rate")]       object r,
        [ExcelArgument(Name = "sigmaV", Description = "Asset value vol")]      object sigmaV)
        => Enabled ? CreditDerivatives.DistanceToDefault(
                         RangeHelper.Scalar(v), RangeHelper.Scalar(d), RangeHelper.Scalar(t),
                         RangeHelper.Scalar(r), RangeHelper.Scalar(sigmaV))
                   : (object)Off;

    [ExcelFunction(Name = "CR_CREDIT_SPREAD", Category = "Finance | Credit", IsThreadSafe = true,
        Description = "Credit spread implied by the Merton model: -ln(PV_Debt / D*exp(-rT)) / T.")]
    public static object CrCreditSpread(
        [ExcelArgument(Name = "V",      Description = "Firm asset value")]     object v,
        [ExcelArgument(Name = "D",      Description = "Face value of debt")]   object d,
        [ExcelArgument(Name = "T",      Description = "Debt maturity")]        object t,
        [ExcelArgument(Name = "r",      Description = "Risk-free rate")]       object r,
        [ExcelArgument(Name = "sigmaV", Description = "Asset value vol")]      object sigmaV)
        => Enabled ? CreditDerivatives.MertonCreditSpread(
                         RangeHelper.Scalar(v), RangeHelper.Scalar(d), RangeHelper.Scalar(t),
                         RangeHelper.Scalar(r), RangeHelper.Scalar(sigmaV))
                   : (object)Off;

    [ExcelFunction(Name = "CR_SURVIVAL_PROB", Category = "Finance | Credit", IsThreadSafe = true,
        Description = "Survival probability from a constant hazard rate. Q(t) = exp(-lambda*t).")]
    public static object CrSurvivalProb(
        [ExcelArgument(Name = "hazardRate", Description = "Constant hazard rate lambda")] object hazardRate,
        [ExcelArgument(Name = "T",          Description = "Time horizon in years")]        object t)
        => Enabled ? CreditDerivatives.SurvivalProbability(RangeHelper.Scalar(hazardRate), RangeHelper.Scalar(t))
                   : (object)Off;

    [ExcelFunction(Name = "CR_HAZARD_FROM_SPREAD", Category = "Finance | Credit", IsThreadSafe = true,
        Description = "Implied hazard rate from a credit spread: lambda ≈ spread / (1 - recovery). Recovery defaults to Settings value.")]
    public static object CrHazardFromSpread(
        [ExcelArgument(Name = "creditSpread",  Description = "Market credit spread (continuously compounded)")] object creditSpread,
        [ExcelArgument(Name = "recoveryRate",  Description = "Recovery rate on default (default from Settings)")] object recoveryRate)
        => Enabled ? CreditDerivatives.HazardRateFromSpread(
                         RangeHelper.Scalar(creditSpread),
                         RangeHelper.IsMissing(recoveryRate) ? RecoveryDefault : RangeHelper.Scalar(recoveryRate))
                   : (object)Off;

    [ExcelFunction(Name = "CR_CDS_SPREAD", Category = "Finance | Credit", IsThreadSafe = true,
        Description = "Fair CDS running spread via hazard-rate model. s = (1-R)*ProtectionLegPV / PremiumLegPV.")]
    public static object CrCdsSpread(
        [ExcelArgument(Name = "hazardRate",   Description = "Constant hazard rate lambda")]                    object hazardRate,
        [ExcelArgument(Name = "r",            Description = "Continuously compounded risk-free rate (flat)")]   object r,
        [ExcelArgument(Name = "maturity",     Description = "CDS maturity in years")]                          object maturity,
        [ExcelArgument(Name = "recoveryRate", Description = "Recovery rate (default from Settings)")]           object recoveryRate,
        [ExcelArgument(Name = "frequency",    Description = "Premium payment frequency per year (default 4)")]  object frequency)
        => Enabled ? CreditDerivatives.CdsFairSpread(
                         RangeHelper.Scalar(hazardRate), RangeHelper.Scalar(r), RangeHelper.Scalar(maturity),
                         RangeHelper.IsMissing(recoveryRate) ? RecoveryDefault : RangeHelper.Scalar(recoveryRate),
                         RangeHelper.IsMissing(frequency) ? 4 : RangeHelper.ScalarInt(frequency))
                   : (object)Off;

    [ExcelFunction(Name = "CR_CDS_MTM", Category = "Finance | Credit", IsThreadSafe = true,
        Description = "Mark-to-market value of an existing CDS (protection buyer perspective). Positive if credit has deteriorated since inception.")]
    public static object CrCdsMtm(
        [ExcelArgument(Name = "contractedSpread",  Description = "Spread agreed at inception")]                      object contractedSpread,
        [ExcelArgument(Name = "hazardRate",        Description = "Current implied hazard rate")]                     object hazardRate,
        [ExcelArgument(Name = "r",                 Description = "Current risk-free rate")]                          object r,
        [ExcelArgument(Name = "notional",          Description = "Notional of the CDS")]                             object notional,
        [ExcelArgument(Name = "remainingMaturity", Description = "Remaining maturity in years")]                     object remainingMaturity,
        [ExcelArgument(Name = "recoveryRate",      Description = "Recovery rate (default from Settings)")]           object recoveryRate,
        [ExcelArgument(Name = "frequency",         Description = "Premium payment frequency (default 4)")]           object frequency)
        => Enabled ? CreditDerivatives.CdsMtm(
                         RangeHelper.Scalar(contractedSpread), RangeHelper.Scalar(hazardRate), RangeHelper.Scalar(r),
                         RangeHelper.Scalar(notional), RangeHelper.Scalar(remainingMaturity),
                         RangeHelper.IsMissing(recoveryRate) ? RecoveryDefault : RangeHelper.Scalar(recoveryRate),
                         RangeHelper.IsMissing(frequency) ? 4 : RangeHelper.ScalarInt(frequency))
                   : (object)Off;
}
