namespace Aleksej.Finance.Excel.Constants;

/// <summary>Names, descriptions, help, and defaults for the Monte Carlo (MC_*) functions.</summary>
internal static class MonteCarloConstants
{
    public const string Help = Cat.HelpBase + "/options/monte-carlo";

    public const int DefaultPaths = 10_000;
    public const int DefaultSteps = 50;
    public const int DefaultSeed  = 42;

    public const string EuropeanName = "MC_EUROPEAN";
    public const string EuropeanDesc = "European option price via Monte Carlo GBM simulation. Converges to Black-Scholes. Use BS_CALL/PUT for production; this is for validation and exotic comparisons.";

    public const string AmericanName = "MC_AMERICAN";
    public const string AmericanDesc = "American option price via Longstaff-Schwartz LSM (2001). Uses Laguerre basis regression for early exercise decisions. NOTE: slow for large path counts — use BT_PRICE for quick estimates.";

    // Argument descriptions
    public const string PathsFull = "Number of simulated paths (default 10000, more = more accurate but slower)";
    public const string Paths     = "Number of simulated paths (default 10000)";
    public const string StepsFull = "Time steps per path (default 50)";
    public const string Steps     = "Backward-induction time steps (default 50)";
    public const string SeedShort = "Random seed (default 42)";
}
