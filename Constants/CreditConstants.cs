namespace Aleksej.Finance.Excel.Constants;

/// <summary>Names, descriptions, help topics, defaults, and argument descriptions for the credit (CR_*) functions.</summary>
internal static class CreditConstants
{
    public const string Help = Cat.HelpBase;

    /// <summary>Default premium payment frequency per year when omitted.</summary>
    public const int DefaultFrequency = 4;

    // ── Function names & descriptions ─────────────────────────────────────────
    public const string MertonEquityName = "CR_MERTON_EQUITY";
    public const string MertonEquityDesc = "Merton model equity value: equity is a call on firm assets V with strike = debt D. E = V*N(d1) - D*exp(-rT)*N(d2).";

    public const string MertonDebtName = "CR_MERTON_DEBT";
    public const string MertonDebtDesc = "Merton model PV of debt. PV(Debt) = V - Equity = D*exp(-rT)*N(d2) + V*N(-d1).";

    public const string DefaultProbName = "CR_DEFAULT_PROB";
    public const string DefaultProbDesc = "Merton risk-neutral probability of default at maturity T. PD = N(-d2). Higher leverage/vol → higher PD.";

    public const string DistToDefaultName = "CR_DIST_TO_DEFAULT";
    public const string DistToDefaultDesc = "Distance to Default (DD) under Merton. DD = d2 = (ln(V/D)+(r-σ²/2)T)/(σ√T). Higher DD = safer firm.";

    public const string CreditSpreadName = "CR_CREDIT_SPREAD";
    public const string CreditSpreadDesc = "Credit spread implied by the Merton model: -ln(PV_Debt / D*exp(-rT)) / T.";

    public const string SurvivalProbName = "CR_SURVIVAL_PROB";
    public const string SurvivalProbDesc = "Survival probability from a constant hazard rate. Q(t) = exp(-lambda*t).";

    public const string HazardFromSpreadName = "CR_HAZARD_FROM_SPREAD";
    public const string HazardFromSpreadDesc = "Implied hazard rate from a credit spread: lambda ≈ spread / (1 - recovery). Recovery defaults to Settings value.";

    public const string CdsSpreadName = "CR_CDS_SPREAD";
    public const string CdsSpreadDesc = "Fair CDS running spread via hazard-rate model. s = (1-R)*ProtectionLegPV / PremiumLegPV.";

    public const string CdsMtmName = "CR_CDS_MTM";
    public const string CdsMtmDesc = "Mark-to-market value of an existing CDS (protection buyer perspective). Positive if credit has deteriorated since inception.";

    // ── Argument descriptions ─────────────────────────────────────────────────
    public const string V                 = "Current total asset value of the firm";
    public const string VShort            = "Firm asset value";
    public const string D                 = "Face value of debt (default boundary at maturity)";
    public const string DShort            = "Face value of debt";
    public const string TMaturity         = "Debt maturity in years";
    public const string TMaturityShort    = "Debt maturity";
    public const string R                 = "Continuously compounded risk-free rate";
    public const string RShort            = "Risk-free rate";
    public const string SigmaV            = "Annualised volatility of firm asset value";
    public const string SigmaVShort       = "Asset value vol";

    public const string HazardRate        = "Constant hazard rate lambda";
    public const string THorizon          = "Time horizon in years";

    public const string CreditSpread      = "Market credit spread (continuously compounded)";
    public const string RecoverySettings  = "Recovery rate on default (default from Settings)";

    public const string RFlat             = "Continuously compounded risk-free rate (flat)";
    public const string CdsMaturity       = "CDS maturity in years";
    public const string Recovery          = "Recovery rate (default from Settings)";
    public const string Frequency         = "Premium payment frequency per year (default 4)";

    public const string ContractedSpread  = "Spread agreed at inception";
    public const string CurrentHazardRate = "Current implied hazard rate";
    public const string CurrentR          = "Current risk-free rate";
    public const string Notional          = "Notional of the CDS";
    public const string RemainingMaturity = "Remaining maturity in years";
    public const string FrequencyShort    = "Premium payment frequency (default 4)";
}
