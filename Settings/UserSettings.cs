using System;
using System.IO;
using System.Xml.Serialization;

namespace Aleksej.Finance.Excel.Settings;

/// <summary>
/// Persistent user settings for the Finance Add-In.
/// Serialized to %AppData%\Aleksej.Finance.Excel\settings.xml.
/// Changes take effect immediately — no Excel restart required.
/// </summary>
[XmlRoot("FinanceAddinSettings")]
public class UserSettings
{
    private static readonly string SettingsPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "Aleksej.Finance.Excel", "settings.xml");

    private static readonly XmlSerializer Serializer = new(typeof(UserSettings));

    // ── Session cache ─────────────────────────────────────────────────────────
    // Functions recalc constantly; without a cache every cell read would hit disk
    // and deserialize XML. Current caches one instance per session; the ribbon calls
    // Invalidate() after each Save() so changes still take effect immediately.
    private static UserSettings? _cached;

    /// <summary>The cached current settings. Loaded once per session (see Invalidate).</summary>
    public static UserSettings Current => _cached ??= Load();

    /// <summary>Drops the cached settings so the next access reloads from disk.</summary>
    public static void Invalidate() => _cached = null;

    /// <summary>Test seam: pins the cached settings without touching disk.</summary>
    internal static void UseForTesting(UserSettings settings) => _cached = settings;

    // ── Default calculation parameters ───────────────────────────────────────

    /// <summary>Annual risk-free rate used when rf is omitted (e.g. 0.05 = 5%).</summary>
    public double DefaultRiskFreeRate { get; set; } = 0.05;

    /// <summary>Trading days per year for annualisation (252 = standard US equity).</summary>
    public int DefaultTradingDays { get; set; } = 252;

    /// <summary>VaR confidence level when not explicitly specified (0.95 = 95%).</summary>
    public double DefaultConfidence { get; set; } = 0.95;

    /// <summary>Default coupon/payment frequency per year (2 = semi-annual).</summary>
    public int DefaultFrequency { get; set; } = 2;

    /// <summary>EWMA decay factor λ (0.94 = JP Morgan RiskMetrics standard).</summary>
    public double DefaultLambda { get; set; } = 0.94;

    /// <summary>Default recovery rate for CDS functions (0.40 = 40%, industry standard).</summary>
    public double DefaultRecoveryRate { get; set; } = 0.40;

    /// <summary>
    /// When true, validation/calculation failures return standard Excel errors
    /// (#NUM! / #VALUE!) so IFERROR can catch them. When false (default), functions
    /// return a descriptive text message explaining what went wrong.
    /// </summary>
    public bool ErrorsAsExcelError { get; set; } = false;

    // ── Function category toggles ─────────────────────────────────────────────

    /// <summary>Enable BS_, GK_, BT_, MC_, EX_, OF_ option pricing functions.</summary>
    public bool EnableOptions { get; set; } = true;

    /// <summary>Enable BOND_, YC_, MORT_ fixed-income functions.</summary>
    public bool EnableBonds { get; set; } = true;

    /// <summary>Enable FWD_, FRA_, IRS_, BM_, SR_ derivatives functions.</summary>
    public bool EnableDerivatives { get; set; } = true;

    /// <summary>Enable CR_ credit risk functions.</summary>
    public bool EnableCredit { get; set; } = true;

    /// <summary>Enable PORT_, RISK_, VOL_ portfolio and risk functions.</summary>
    public bool EnablePortfolioRisk { get; set; } = true;

    /// <summary>Enable FEE_, ATTR_, EQ_ fees and attribution functions.</summary>
    public bool EnableFeesAttribution { get; set; } = true;

    /// <summary>Enable EDGAR_ and TREASURY_ live data functions (Phase 2).</summary>
    public bool EnableLiveData { get; set; } = false;

    // ── Live data settings (Phase 2) ──────────────────────────────────────────

    /// <summary>SEC EDGAR User-Agent header (required by policy: "Name email@example.com").</summary>
    public string EdgarUserAgent { get; set; } = string.Empty;

    /// <summary>Live data cache time-to-live in minutes.</summary>
    public int DataCacheTtlMinutes { get; set; } = 15;

    // ── Load / Save ───────────────────────────────────────────────────────────

    /// <summary>
    /// Loads settings from disk. Returns defaults if the file does not exist or cannot be read.
    /// </summary>
    public static UserSettings Load()
    {
        try
        {
            if (!File.Exists(SettingsPath)) return new UserSettings();
            using var stream = File.OpenRead(SettingsPath);
            return (UserSettings?)Serializer.Deserialize(stream) ?? new UserSettings();
        }
        catch
        {
            return new UserSettings();
        }
    }

    /// <summary>Persists the current settings to disk.</summary>
    public void Save()
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath)!);
            using var stream = File.Create(SettingsPath);
            Serializer.Serialize(stream, this);
        }
        catch { /* silently ignore write failures */ }
    }

    /// <summary>Resets all settings to their factory defaults and saves.</summary>
    public static void ResetToDefaults()
    {
        new UserSettings().Save();
    }
}
