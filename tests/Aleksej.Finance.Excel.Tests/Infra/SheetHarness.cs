using System;
using System.Diagnostics;
using ExcelDna.Testing;

namespace Aleksej.Finance.Excel.Tests.Infra
{
    /// <summary>
    /// Reusable mechanic for driving the loaded add-in: write inputs, set a formula,
    /// recalc, read the result. Excel COM objects are driven via <c>dynamic</c> (late-bound
    /// IDispatch) since the embedded interop only surfaces Application/Workbook.
    /// </summary>
    internal static class SheetHarness
    {
        private const int ResultRow = 1;
        private const int ResultCol = 50;   // AX1 — well clear of any input data

        public static dynamic NewSheet()
        {
            dynamic wb = Util.Workbook;
            return wb.Worksheets.Add();
        }

        /// <summary>Set a literal formula in the result cell, recalc, return its value.</summary>
        public static object Eval(dynamic ws, string formula)
        {
            dynamic cell = ws.Cells[ResultRow, ResultCol];
            cell.Formula = formula.StartsWith("=") ? formula : "=" + formula;
            ((dynamic)Util.Application).CalculateFull();
            return cell.Value2;
        }

        /// <summary>Recalc until the result cell turns numeric (for async UDFs) or times out.</summary>
        public static object Poll(dynamic ws, string formula, int timeoutMs = 30000)
        {
            dynamic cell = ws.Cells[ResultRow, ResultCol];
            cell.Formula = formula.StartsWith("=") ? formula : "=" + formula;
            dynamic app = Util.Application;
            var sw = Stopwatch.StartNew();
            object v = cell.Value2;
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                if (v is double) return v;
                app.CalculateFull();
                System.Threading.Thread.Sleep(200);
                v = cell.Value2;
            }
            return v;
        }

        /// <summary>Write a vector down a column; returns its A1 range reference.</summary>
        public static string WriteColumn(dynamic ws, int row, int col, double[] v)
        {
            for (int i = 0; i < v.Length; i++)
                ws.Cells[row + i, col].Value2 = v[i];
            return ColName(col) + row + ":" + ColName(col) + (row + v.Length - 1);
        }

        /// <summary>Write a matrix as a 2-D block; returns its A1 range reference.</summary>
        public static string WriteBlock(dynamic ws, int row, int col, double[,] m)
        {
            int r = m.GetLength(0), c = m.GetLength(1);
            for (int i = 0; i < r; i++)
                for (int j = 0; j < c; j++)
                    ws.Cells[row + i, col + j].Value2 = m[i, j];
            return ColName(col) + row + ":" + ColName(col + c - 1) + (row + r - 1);
        }

        /// <summary>Read a spilled column of length n starting at the result cell.</summary>
        public static double[] ReadSpillColumn(dynamic ws, int n)
        {
            var outv = new double[n];
            for (int i = 0; i < n; i++)
            {
                object v = ws.Cells[ResultRow + i, ResultCol].Value2;
                outv[i] = v is double d ? d : double.NaN;
            }
            return outv;
        }

        /// <summary>Judge a result against a case; returns false + message on failure.</summary>
        public static bool Passes(object res, UdfCase c, out string msg)
        {
            msg = "";
            bool isNum = res is double;
            double d = isNum ? (double)res : double.NaN;
            switch (c.Check)
            {
                case Check.Num:
                    if (!isNum) { msg = $"{c.Func}: non-numeric '{res}'"; return false; }
                    if (Math.Abs(d - c.Expected) > c.Tol) { msg = $"{c.Func}: got {d}, expected {c.Expected} +/- {c.Tol}"; return false; }
                    return true;
                case Check.Pos:
                    if (!(isNum && d > 0)) { msg = $"{c.Func}: expected positive number, got '{res}'"; return false; }
                    return true;
                case Check.Neg:
                    if (!(isNum && d < 0)) { msg = $"{c.Func}: expected negative number, got '{res}'"; return false; }
                    return true;
                case Check.In01:
                    if (!(isNum && d >= 0 && d <= 1)) { msg = $"{c.Func}: expected number in [0,1], got '{res}'"; return false; }
                    return true;
                case Check.IsNum:
                    if (!isNum) { msg = $"{c.Func}: expected a number, got '{res}'"; return false; }
                    return true;
                case Check.Text:
                    if (!string.Equals(res as string, c.ExpectedText, StringComparison.Ordinal))
                    { msg = $"{c.Func}: got '{res}', expected text '{c.ExpectedText}'"; return false; }
                    return true;
            }
            return true;
        }

        public static string ColName(int col)
        {
            string s = "";
            while (col > 0) { int m = (col - 1) % 26; s = (char)('A' + m) + s; col = (col - m - 1) / 26; }
            return s;
        }
    }
}
