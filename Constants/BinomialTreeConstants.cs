namespace Aleksej.Finance.Excel.Constants;

/// <summary>Names, descriptions, help topics, and defaults for the binomial tree (BT_*) functions.</summary>
internal static class BinomialTreeConstants
{
    public const string Help = Cat.HelpBase + "/options/binomial-tree";

    public const int DefaultSteps = 200;

    public const string PriceName = "BT_PRICE";
    public const string PriceDesc = "CRR binomial tree option price. Supports both European and American exercise. Converges to Black-Scholes for European options.";

    public const string DeltaName = "BT_DELTA";
    public const string DeltaDesc = "Binomial tree Delta — extracted from level-1 nodes. More accurate for American options than BS_DELTA.";

    public const string GammaName = "BT_GAMMA";
    public const string GammaDesc = "Binomial tree Gamma — second-order finite difference from level-1 nodes.";

    public const string StepsDefaultDesc = "Number of time steps (default 200, higher = more accurate)";
    public const string StepsDesc        = "Number of time steps";
    public const string IsPutDefaultDesc = "TRUE for put, FALSE for call (default FALSE)";
    public const string IsPutDesc        = "TRUE for put";
    public const string IsAmericanDefaultDesc = "TRUE for American (early exercise), FALSE for European (default FALSE)";
    public const string IsAmericanDesc        = "TRUE for American";
}
