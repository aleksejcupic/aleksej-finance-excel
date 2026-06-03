using System;
using ExcelDna.Integration;
using Aleksej.Finance.Excel.Settings;

namespace Aleksej.Finance.Excel.Helpers;

/// <summary>Function category, mapped 1:1 to the UserSettings Enable* toggles.</summary>
internal enum Category
{
    Options,
    Bonds,
    Derivatives,
    Credit,
    PortfolioRisk,
    FeesAttribution,
    LiveData,
}

/// <summary>
/// Central execution wrapper for every UDF. Handles the category-enabled check,
/// turns thrown <see cref="FinanceInputException"/> / unexpected exceptions into
/// friendly cell output (descriptive text, or Excel errors when the user opts in),
/// and provides an async path for heavy functions.
/// </summary>
internal static class Fn
{
    internal static bool IsEnabled(Category c) => c switch
    {
        Category.Options         => UserSettings.Current.EnableOptions,
        Category.Bonds           => UserSettings.Current.EnableBonds,
        Category.Derivatives     => UserSettings.Current.EnableDerivatives,
        Category.Credit          => UserSettings.Current.EnableCredit,
        Category.PortfolioRisk   => UserSettings.Current.EnablePortfolioRisk,
        Category.FeesAttribution => UserSettings.Current.EnableFeesAttribution,
        Category.LiveData        => UserSettings.Current.EnableLiveData,
        _                        => true,
    };

    private static string Label(Category c) => c switch
    {
        Category.PortfolioRisk   => "Portfolio & Risk",
        Category.FeesAttribution => "Fees & Attribution",
        Category.LiveData        => "Live Data",
        _                        => c.ToString(),
    };

    /// <summary>Runs a synchronous function body with enabled-check + error handling.</summary>
    internal static object Run(Category cat, Func<object> body)
    {
        if (!IsEnabled(cat)) return RangeHelper.DisabledMessage(Label(cat));
        try
        {
            return body();
        }
        catch (FinanceInputException ex)
        {
            return Format(ex.Message, isValidation: true);
        }
        catch (Exception ex)
        {
            return Format($"#ERROR: {ex.Message}", isValidation: false);
        }
    }

    /// <summary>
    /// Runs an expensive function body on a background thread via ExcelAsyncUtil.
    /// The cell shows #N/A while computing, then Excel re-evaluates with the result.
    /// <paramref name="args"/> is the recalculation cache key. Functions using this
    /// must NOT be marked IsThreadSafe.
    /// </summary>
    internal static object RunAsync(string name, object[] args, Category cat, Func<object> compute)
    {
        if (!IsEnabled(cat)) return RangeHelper.DisabledMessage(Label(cat));
        return ExcelAsyncUtil.Run(name, args, () =>
        {
            try
            {
                return compute();
            }
            catch (FinanceInputException ex)
            {
                return Format(ex.Message, isValidation: true);
            }
            catch (Exception ex)
            {
                return Format($"#ERROR: {ex.Message}", isValidation: false);
            }
        });
    }

    /// <summary>
    /// Returns the descriptive message as text, or a standard Excel error when the
    /// user has enabled <see cref="UserSettings.ErrorsAsExcelError"/> (so IFERROR works).
    /// </summary>
    private static object Format(string message, bool isValidation)
    {
        if (!UserSettings.Current.ErrorsAsExcelError) return message;
        return isValidation ? ExcelError.ExcelErrorNum : ExcelError.ExcelErrorValue;
    }
}
