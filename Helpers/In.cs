using System;
using System.Globalization;
using ExcelDna.Integration;

namespace Aleksej.Finance.Excel.Helpers;

/// <summary>
/// Typed, validating, normalizing accessors for Excel function arguments.
/// Every wrapper passes raw <c>object</c> cell values through these instead of
/// <see cref="RangeHelper"/> directly, so bad input throws a
/// <see cref="FinanceInputException"/> with a clear message (surfaced in the cell
/// by <see cref="Fn"/>) rather than silently producing NaN.
/// </summary>
internal static class In
{
    // ── Core scalar parse + normalize ─────────────────────────────────────────

    /// <summary>
    /// Parses an Excel cell to a finite double. Normalizes percent text ("5%" → 0.05),
    /// trims strings, and parses under invariant then current culture. A blank/omitted
    /// cell yields <paramref name="def"/> if supplied, otherwise throws "required".
    /// Excel errors (e.g. #N/A flowing in) throw rather than poisoning the result.
    /// </summary>
    private static double ParseScalar(string name, object val, double? def)
    {
        switch (val)
        {
            case double d:
                return d;
            case int i:
                return i;
            case bool b:
                throw new FinanceInputException($"{name} must be a number, not TRUE/FALSE.");
            case ExcelMissing:
            case ExcelEmpty:
                if (def.HasValue) return def.Value;
                throw new FinanceInputException($"{name} is required.");
            case ExcelError e:
                throw new FinanceInputException($"{name} received an Excel error ({e}).");
            case string s:
                return ParseString(name, s, def);
            default:
                throw new FinanceInputException($"{name} could not be interpreted as a number.");
        }
    }

    private static double ParseString(string name, string raw, double? def)
    {
        string s = raw.Trim();
        if (s.Length == 0)
        {
            if (def.HasValue) return def.Value;
            throw new FinanceInputException($"{name} is required.");
        }

        // Percent TEXT only (e.g. "5%" → 0.05). A numeric 0.05 from a %-formatted
        // cell already arrives as a double and never reaches here, so no double-division.
        bool isPercent = s.EndsWith("%", StringComparison.Ordinal);
        if (isPercent) s = s[..^1].Trim();

        if (!double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double v) &&
            !double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out v))
            throw new FinanceInputException($"{name}: cannot interpret \"{raw}\" as a number.");

        return isPercent ? v / 100.0 : v;
    }

    private static double Finite(string name, double v)
    {
        if (double.IsNaN(v) || double.IsInfinity(v))
            throw new FinanceInputException($"{name} must be a finite number.");
        return v;
    }

    // ── Scalar validators ─────────────────────────────────────────────────────

    /// <summary>Strictly positive, finite. For prices, strikes, firm/equity/debt values.</summary>
    internal static double Price(string name, object val)
    {
        double v = Finite(name, ParseScalar(name, val, null));
        if (v <= 0) throw new FinanceInputException($"{name} must be greater than 0.");
        return v;
    }

    /// <summary>Finite; may be negative (rates can be). Optional default.</summary>
    internal static double Rate(string name, object val, double? def = null) =>
        Finite(name, ParseScalar(name, val, def));

    /// <summary>Strictly positive, finite. For volatilities.</summary>
    internal static double Vol(string name, object val)
    {
        double v = Finite(name, ParseScalar(name, val, null));
        if (v <= 0) throw new FinanceInputException($"{name} (volatility) must be greater than 0.");
        return v;
    }

    /// <summary>Non-negative, finite. For times/maturities in years.</summary>
    internal static double Years(string name, object val)
    {
        double v = Finite(name, ParseScalar(name, val, null));
        if (v < 0) throw new FinanceInputException($"{name} must be 0 or greater.");
        return v;
    }

    /// <summary>Finite and within [0, 1]. For probabilities, confidence, recovery, lambda.</summary>
    internal static double Prob(string name, object val, double? def = null)
    {
        double v = Finite(name, ParseScalar(name, val, def));
        if (v < 0 || v > 1) throw new FinanceInputException($"{name} must be between 0 and 1.");
        return v;
    }

    /// <summary>Integer ≥ 1. For frequency, paths, steps, nDays, trading days, seed.</summary>
    internal static int PosInt(string name, object val, int? def = null)
    {
        double v = Finite(name, ParseScalar(name, val, def));
        int i = (int)Math.Round(v);
        if (i < 1) throw new FinanceInputException($"{name} must be a whole number ≥ 1.");
        return i;
    }

    /// <summary>Integer ≥ 0. For counts that can legitimately be zero (e.g. payments made).</summary>
    internal static int Count(string name, object val, int? def = null)
    {
        double v = Finite(name, ParseScalar(name, val, def));
        int i = (int)Math.Round(v);
        if (i < 0) throw new FinanceInputException($"{name} must be a whole number ≥ 0.");
        return i;
    }

    /// <summary>Any finite number (may be negative or zero). For model params, generic inputs.</summary>
    internal static double Num(string name, object val, double? def = null) =>
        Finite(name, ParseScalar(name, val, def));

    /// <summary>Boolean flag. Accepts TRUE/FALSE, 1/0, "true"/"false"; missing → default.</summary>
    internal static bool Flag(string name, object val, bool def = false) =>
        RangeHelper.ScalarBool(val, def);

    // ── Array validators (modular — usable by any array-consuming function) ────

    /// <summary>Non-empty flat double[] from any range shape (row, column, or 2-D block).</summary>
    internal static double[] Vector(string name, object range)
    {
        double[] a = RangeHelper.ToDoubleArray(range);
        if (a.Length == 0)
            throw new FinanceInputException($"{name} must contain at least one numeric value.");
        return a;
    }

    /// <summary>Non-empty double[rows, cols] from a 2-D range.</summary>
    internal static double[,] Matrix(string name, object range)
    {
        double[,] m = RangeHelper.ToDoubleMatrix(range);
        if (m.GetLength(0) == 0 || m.GetLength(1) == 0)
            throw new FinanceInputException($"{name} must contain at least one numeric value.");
        return m;
    }
}
