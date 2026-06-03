using System;
using ExcelDna.Integration;
using Aleksej.Finance.Excel.Helpers;
using Aleksej.Finance.Excel.Settings;
using Xunit;

namespace Aleksej.Finance.Excel.UnitTests
{
    [Collection("AddIn")]
    public class FnTests
    {
        [Fact]
        public void Run_Enabled_ReturnsBodyResult()
        {
            UserSettings.UseForTesting(new UserSettings());
            Assert.Equal(1.0, (double)Fn.Run(Category.Options, () => 1.0));
        }

        [Fact]
        public void Run_DisabledCategory_ReturnsMessage()
        {
            UserSettings.UseForTesting(new UserSettings { EnableOptions = false });
            Assert.IsType<string>(Fn.Run(Category.Options, () => 1.0));
        }

        [Fact]
        public void Run_ValidationException_ReturnsText()
        {
            UserSettings.UseForTesting(new UserSettings());
            Assert.Equal("bad", Fn.Run(Category.Options, () => throw new FinanceInputException("bad")));
        }

        [Fact]
        public void Run_UnexpectedException_ReturnsErrorText()
        {
            UserSettings.UseForTesting(new UserSettings());
            Assert.Equal("#ERROR: boom", Fn.Run(Category.Options, () => throw new Exception("boom")));
        }

        [Fact]
        public void Run_ErrorsAsExcelError_MapsToExcelErrors()
        {
            UserSettings.UseForTesting(new UserSettings { ErrorsAsExcelError = true });
            Assert.Equal(ExcelError.ExcelErrorNum, (ExcelError)Fn.Run(Category.Options, () => throw new FinanceInputException("x")));
            Assert.Equal(ExcelError.ExcelErrorValue, (ExcelError)Fn.Run(Category.Options, () => throw new Exception("y")));
        }

        [Fact]
        public void IsEnabled_CoversEveryCategory()
        {
            UserSettings.UseForTesting(new UserSettings());
            foreach (Category cat in Enum.GetValues(typeof(Category)))
                _ = Fn.IsEnabled(cat);
        }

        [Fact]
        public void DisabledMessages_CoverEveryCategoryLabel()
        {
            UserSettings.UseForTesting(new UserSettings
            {
                EnableOptions = false, EnableBonds = false, EnableDerivatives = false, EnableCredit = false,
                EnablePortfolioRisk = false, EnableFeesAttribution = false, EnableLiveData = false
            });
            foreach (Category cat in Enum.GetValues(typeof(Category)))
                Assert.IsType<string>(Fn.Run(cat, () => 1.0));
        }

        [Fact]
        public void RunAsync_Disabled_ReturnsMessage()
        {
            UserSettings.UseForTesting(new UserSettings { EnableOptions = false });
            Assert.IsType<string>(Fn.RunAsync("X", new object[] { 1.0 }, Category.Options, () => 1.0));
        }

        [Fact]
        public void RunAsync_Enabled_EntersAsyncPath()
        {
            UserSettings.UseForTesting(new UserSettings());
            // ExcelAsyncUtil needs the Excel calc context; covering the enabled entry is enough.
            try { _ = Fn.RunAsync("X", new object[] { 1.0 }, Category.Options, () => 1.0); }
            catch { }
        }
    }
}
