namespace Aleksej.Finance.Excel.Constants;

/// <summary>
/// Excel function category strings (the grouping shown in the Insert Function dialog)
/// and the base URL for online help topics. Centralized so every function references
/// one source instead of repeating literals.
/// </summary>
internal static class Cat
{
    public const string Options     = "Finance | Options";
    public const string Bonds       = "Finance | Bonds";
    public const string Derivatives = "Finance | Derivatives";
    public const string Credit      = "Finance | Credit";
    public const string Portfolio   = "Finance | Portfolio";
    public const string Risk        = "Finance | Risk";
    public const string Equity      = "Finance | Equity";
    public const string Fees        = "Finance | Fees";
    public const string Attribution = "Finance | Attribution";

    /// <summary>Online help target (the repo README documents every function).</summary>
    public const string HelpBase = "https://github.com/aleksejcupic/aleksej-finance-excel";
}
