using System;
using System.Collections.Generic;
using Xunit;
using Aleksej.Finance.Excel.Tests.Data;
using Aleksej.Finance.Excel.Tests.Infra;

namespace Aleksej.Finance.Excel.Tests.Tests
{
    [ExcelTestSettings(AddIn = "Aleksej.Finance.ExcelAddin.xll")]
    public class ScalarUdfTests
    {
        // Evaluates every scalar UDF through the loaded add-in in real Excel and aggregates
        // all mismatches into a single readable failure.
        [ExcelFact]
        public void AllScalarUdfsEvaluateCorrectly()
        {
            dynamic ws = SheetHarness.NewSheet();
            var failures = new List<string>();
            foreach (var c in ScalarCases.All)
            {
                object res;
                try { res = SheetHarness.Eval(ws, c.Formula); }
                catch (Exception ex) { failures.Add($"{c.Func}: threw {ex.Message}"); continue; }
                if (!SheetHarness.Passes(res, c, out string msg)) failures.Add(msg);
            }
            Assert.True(failures.Count == 0,
                $"{failures.Count}/{ScalarCases.All.Length} scalar UDF(s) failed:\n" + string.Join("\n", failures));
        }
    }
}
