using System;
using Xunit;
using Aleksej.Finance.Excel.Tests.Infra;

namespace Aleksej.Finance.Excel.Tests.Tests
{
    // Validation + normalization behaviour, end-to-end through the plugin.
    // NOTE: assumes the default error mode (UserSettings.ErrorsAsExcelError = false), which
    // returns descriptive text. If the "Errors as Excel errors" ribbon toggle is on, the
    // cell returns #NUM! instead and the first two asserts would need adjusting.
    [ExcelTestSettings(AddIn = "Aleksej.Finance.ExcelAddin.xll")]
    public class ErrorHandlingTests
    {
        [ExcelFact]
        public void ValidationReturnsDescriptiveMessages()
        {
            dynamic ws = SheetHarness.NewSheet();
            Assert.Equal("sigma (volatility) must be greater than 0.",
                SheetHarness.Eval(ws, "BS_CALL(100,100,1,0.05,0)"));
            Assert.Equal("S must be greater than 0.",
                SheetHarness.Eval(ws, "BS_CALL(-5,100,1,0.05,0.2)"));
        }

        [ExcelFact]
        public void PercentTextIsNormalizedToFraction()
        {
            dynamic ws = SheetHarness.NewSheet();
            // sigma supplied as the TEXT "20%" must normalize to 0.20 -> same as numeric 0.20.
            object v = SheetHarness.Eval(ws, "BS_CALL(100,100,1,0.05,\"20%\")");
            Assert.True(v is double d && Math.Abs(d - 10.4506) < 0.001, $"got '{v}'");
        }
    }
}
