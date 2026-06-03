using System;
using System.Collections.Generic;
using System.Linq;
using Xunit;
using Aleksej.Finance.Excel.Tests.Infra;

namespace Aleksej.Finance.Excel.Tests.Tests
{
    // Array-returning portfolio UDFs return a weight vector. We assert via SUM(result)==1
    // (which works regardless of dynamic-array spilling), so the check is robust on any Excel.
    [ExcelTestSettings(AddIn = "Aleksej.Finance.ExcelAddin.xll")]
    public class ArrayUdfTests
    {
        [ExcelFact]
        public void ArrayUdfWeightsSumToOne()
        {
            dynamic ws = SheetHarness.NewSheet();
            SheetHarness.WriteColumn(ws, 1, 11, new[] { 0.5, 0.3, 0.2 });       // K  weights
            SheetHarness.WriteColumn(ws, 1, 12, new[] { 0.08, 0.12, 0.18 });    // L  mu
            SheetHarness.WriteBlock(ws, 1, 13, new[,] { { 0.04, 0.01, 0.00 }, { 0.01, 0.06, 0.01 }, { 0.00, 0.01, 0.09 } }); // M:O cov

            var cases = new (string func, string formula)[]
            {
                ("PORT_MIN_VAR",      "SUM(PORT_MIN_VAR(M1:O3))"),
                ("PORT_RISK_PARITY",  "SUM(PORT_RISK_PARITY(M1:O3))"),
                ("PORT_MAX_SHARPE",   "SUM(PORT_MAX_SHARPE(L1:L3,M1:O3,0.02))"),
                ("PORT_RISK_CONTRIB", "SUM(PORT_RISK_CONTRIB(K1:K3,M1:O3))"),
            };

            var failures = new List<string>();
            foreach (var c in cases)
            {
                object res = SheetHarness.Eval(ws, c.formula);
                if (!(res is double d)) { failures.Add($"{c.func}: non-numeric result '{res}'"); continue; }
                if (Math.Abs(d - 1.0) > 0.001) failures.Add($"{c.func}: SUM(weights)={d} (expected 1)");
            }
            Assert.True(failures.Count == 0, "Array UDF(s) failed:\n" + string.Join("\n", failures));
        }
    }
}
