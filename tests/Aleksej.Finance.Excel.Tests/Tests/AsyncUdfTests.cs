using System;
using Xunit;
using Aleksej.Finance.Excel.Tests.Infra;

namespace Aleksej.Finance.Excel.Tests.Tests
{
    // Monte Carlo UDFs run via ExcelAsyncUtil — the cell shows #N/A until the background
    // pool fills it. Poll (recalc + wait) until numeric, then assert.
    [ExcelTestSettings(AddIn = "Aleksej.Finance.ExcelAddin.xll")]
    public class AsyncUdfTests
    {
        [ExcelFact]
        public void McEuropeanConvergesToBlackScholes()
        {
            dynamic ws = SheetHarness.NewSheet();
            object v = SheetHarness.Poll(ws, "MC_EUROPEAN(100,100,1,0.05,0.2,50000,50,FALSE,42)", 30000);
            Assert.True(v is double, $"MC_EUROPEAN never produced a number: '{v}'");
            Assert.True(Math.Abs((double)v - 10.4506) < 0.5, $"MC_EUROPEAN={v}, expected ~10.45");
        }

        [ExcelFact]
        public void McAmericanResolvesToPositive()
        {
            dynamic ws = SheetHarness.NewSheet();
            object v = SheetHarness.Poll(ws, "MC_AMERICAN(100,100,1,0.05,0.2,20000,50,TRUE,42)", 30000);
            Assert.True(v is double d && d > 0, $"MC_AMERICAN did not resolve to a positive number: '{v}'");
        }
    }
}
