using System;
using System.Collections.Generic;
using ExcelDna.Integration;
using Aleksej.Finance.Excel.Settings;

namespace Aleksej.Finance.Excel.Helpers;

/// <summary>
/// Converts Excel cell ranges and scalar values to the types expected by the
/// AleksejCupic.FinancialMath library. All function wrappers go through this class.
/// </summary>
internal static class RangeHelper
{
    // ── Scalar extraction ─────────────────────────────────────────────────────

    /// <summary>
    /// Extracts a double from an Excel cell value.
    /// Returns defaultVal for missing/empty cells.
    /// </summary>
    internal static double Scalar(object val, double defaultVal = 0.0)
    {
        return val switch
        {
            double d              => d,
            int i                 => i,
            ExcelMissing          => defaultVal,
            ExcelEmpty            => defaultVal,
            ExcelError            => double.NaN,
            string s when double.TryParse(s, out double r) => r,
            _                     => defaultVal
        };
    }

    /// <summary>Returns true if an argument was omitted in the Excel formula.</summary>
    internal static bool IsMissing(object val) =>
        val is ExcelMissing or ExcelEmpty;

    /// <summary>
    /// Extracts an int from an Excel cell value.
    /// Returns defaultVal for missing/empty cells.
    /// </summary>
    internal static int ScalarInt(object val, int defaultVal = 0) =>
        (int)Math.Round(Scalar(val, defaultVal));

    /// <summary>
    /// Extracts a bool from an Excel cell value.
    /// Accepts TRUE/FALSE, 1/0, or "true"/"false".
    /// </summary>
    internal static bool ScalarBool(object val, bool defaultVal = false)
    {
        return val switch
        {
            bool b                 => b,
            double d               => d != 0,
            ExcelMissing           => defaultVal,
            ExcelEmpty             => defaultVal,
            string s               => string.Equals(s, "true", StringComparison.OrdinalIgnoreCase),
            _                      => defaultVal
        };
    }

    // ── Array extraction ──────────────────────────────────────────────────────

    /// <summary>
    /// Converts an Excel range (or single cell) to a flat double[] array.
    /// Skips empty/error cells. Handles column ranges, row ranges, and 2-D blocks.
    /// </summary>
    internal static double[] ToDoubleArray(object range)
    {
        var list = new List<double>();
        AddValues(range, list);
        return list.ToArray();
    }

    /// <summary>
    /// Converts an Excel range to a double[rows, cols] matrix.
    /// Rows correspond to the first dimension, columns to the second.
    /// </summary>
    internal static double[,] ToDoubleMatrix(object range)
    {
        if (range is double d) return new double[1, 1] { { d } };

        if (range is object[,] arr)
        {
            int rows = arr.GetLength(0), cols = arr.GetLength(1);
            var m = new double[rows, cols];
            for (int r = 0; r < rows; r++)
                for (int c = 0; c < cols; c++)
                    m[r, c] = Scalar(arr[r, c]);
            return m;
        }

        if (range is double[,] dm) return dm;

        // Single scalar
        return new double[1, 1] { { Scalar(range) } };
    }

    // ── Error / disabled helpers ──────────────────────────────────────────────

    /// <summary>
    /// Returns a category-disabled message string if the category is turned off.
    /// Use as: if (Disabled(cfg.EnableOptions, "Options")) return DisabledMsg("Options");
    /// </summary>
    internal static string DisabledMessage(string category) =>
        $"{category} functions are disabled — enable them in the Finance Tools ribbon.";

    // ── Private helpers ───────────────────────────────────────────────────────

    private static void AddValues(object range, List<double> list)
    {
        switch (range)
        {
            case double d:
                list.Add(d);
                break;
            case object[,] arr:
                foreach (object obj in arr)
                    if (obj is double dv) list.Add(dv);
                break;
            case double[,] dm:
                foreach (double dv in dm) list.Add(dv);
                break;
            case ExcelMissing:
            case ExcelEmpty:
                break;
            default:
                if (double.TryParse(range?.ToString(), out double parsed))
                    list.Add(parsed);
                break;
        }
    }
}
