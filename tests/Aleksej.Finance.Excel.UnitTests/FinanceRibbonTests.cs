using ExcelDna.Integration.CustomUI;
using Aleksej.Finance.Excel.Ribbon;
using Aleksej.Finance.Excel.UnitTests.Infra;
using Xunit;

namespace Aleksej.Finance.Excel.UnitTests
{
    [Collection("AddIn")]
    public class FinanceRibbonTests
    {
        [Fact]
        public void RibbonMembersExecute()
        {
            using var backup = new SettingsBackup();   // ribbon handlers write settings; restore after
            var r = new FinanceRibbon();
            IRibbonControl c = null!;              // handlers ignore the control argument

            Assert.Contains("customUI", r.GetCustomUI("ribbonId"));
            r.OnRibbonLoad(null!);

            // getters
            Assert.NotNull(r.GetRiskFreeRate(c));
            Assert.NotNull(r.GetLambda(c));
            _ = r.GetTradingDaysIndex(c);
            _ = r.GetConfidenceIndex(c);
            _ = r.GetFrequencyIndex(c);
            _ = r.GetEnableOptions(c);
            _ = r.GetEnableBonds(c);
            _ = r.GetEnableDerivatives(c);
            _ = r.GetEnableCredit(c);
            _ = r.GetEnablePortfolioRisk(c);
            _ = r.GetEnableFeesAttribution(c);
            _ = r.GetEnableLiveData(c);
            _ = r.GetErrorsAsExcelError(c);
            Assert.StartsWith("v", r.GetVersionLabel(c));

            // change handlers (valid + invalid branches)
            r.OnRiskFreeRateChange(c, "0.04");
            r.OnRiskFreeRateChange(c, "not-a-number");
            r.OnLambdaChange(c, "0.95");
            r.OnLambdaChange(c, "bad");
            r.OnTradingDaysChange(c, "id", 1);
            r.OnTradingDaysChange(c, "id", 2);
            r.OnConfidenceChange(c, "id", 1);
            r.OnConfidenceChange(c, "id", 2);
            r.OnFrequencyChange(c, "id", 0);
            r.OnFrequencyChange(c, "id", 2);
            r.OnFrequencyChange(c, "id", 3);

            // toggles
            r.OnEnableOptions(c, false);
            r.OnEnableBonds(c, false);
            r.OnEnableDerivatives(c, false);
            r.OnEnableCredit(c, false);
            r.OnEnablePortfolioRisk(c, false);
            r.OnEnableFeesAttribution(c, false);
            r.OnEnableLiveData(c, true);
            r.OnErrorsAsExcelError(c, true);

            r.OnResetDefaults(c);
            // NOTE: OnOpenDocs is intentionally NOT called - it launches the docs URL in a
            // real browser (Process.Start), which we don't want firing during tests.
            try { r.OnAbout(c); } catch { }       // XlCall needs Excel; covers version lookup first
        }
    }
}
