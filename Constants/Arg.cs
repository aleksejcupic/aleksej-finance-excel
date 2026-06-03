namespace Aleksej.Finance.Excel.Constants;

/// <summary>
/// Descriptions for arguments that recur across many functions (S, K, T, r, sigma, …).
/// Function-specific argument descriptions live in that function file's own constants
/// class. Centralized here so the common ones are written once.
/// </summary>
internal static class Arg
{
    public const string S          = "Current asset price";
    public const string K          = "Strike price";
    public const string T          = "Time to expiry in years";
    public const string R          = "Continuous risk-free rate";
    public const string Sigma      = "Annualised volatility";
    public const string IsPut      = "TRUE for put, FALSE for call";
    public const string MarketPrice = "Observed market option price";
    public const string Rf         = "Foreign / dividend continuous yield";
    public const string Face       = "Face (par) value";
    public const string CouponRate = "Annual coupon rate (e.g. 0.05 = 5%)";
    public const string Ytm        = "Yield to maturity (continuous)";
    public const string Years      = "Maturity in years";
    public const string Seed       = "Random seed for reproducibility (default 42)";
}
