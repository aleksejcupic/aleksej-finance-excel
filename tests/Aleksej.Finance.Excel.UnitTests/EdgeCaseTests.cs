using ExcelDna.Integration;
using Aleksej.Finance.Excel.Helpers;
using Aleksej.Finance.Excel.Settings;
using Aleksej.Finance.Excel.UnitTests.Infra;
using Xunit;

namespace Aleksej.Finance.Excel.UnitTests
{
    [Collection("AddIn")]
    public class EdgeCaseTests
    {
        [Fact]
        public void UserSettings_LoadSaveResetRoundtrip()
        {
            using var _ = new SettingsBackup();

            UserSettings.Invalidate();
            Assert.True(UserSettings.Load().EnableOptions);                 // Load (defaults or file)

            new UserSettings { DefaultRiskFreeRate = 0.077 }.Save();        // Save
            UserSettings.Invalidate();
            Assert.Equal(0.077, UserSettings.Load().DefaultRiskFreeRate, 6);

            UserSettings.ResetToDefaults();                                 // ResetToDefaults
            UserSettings.Invalidate();
            Assert.Equal(0.05, UserSettings.Load().DefaultRiskFreeRate, 6);

            UserSettings.UseForTesting(new UserSettings());
            Assert.Same(UserSettings.Current, UserSettings.Current);        // cached instance
        }

        [Fact]
        public void RangeHelper_Branches()
        {
            Assert.Equal(5.0, RangeHelper.Scalar(5));                       // int
            Assert.True(double.IsNaN(RangeHelper.Scalar(ExcelError.ExcelErrorNA)));
            Assert.Equal(0.0, RangeHelper.Scalar(ExcelEmpty.Value));
            Assert.Equal(3.0, RangeHelper.Scalar("3"));
            Assert.Equal(0.0, RangeHelper.Scalar(new object()));           // unhandled -> default
            Assert.Equal(3, RangeHelper.ScalarInt(2.6));
            Assert.False(RangeHelper.ScalarBool(ExcelMissing.Value));
            Assert.True(RangeHelper.ScalarBool(1.0));
            Assert.False(RangeHelper.ScalarBool("nope"));
            Assert.True(RangeHelper.IsMissing(ExcelMissing.Value));

            Assert.Equal(5.0, RangeHelper.ToDoubleMatrix(5.0)[0, 0]);       // scalar -> [1,1]
            Assert.Equal(2.0, RangeHelper.ToDoubleMatrix(new double[,] { { 1, 2 } })[0, 1]); // double[,]
            Assert.Equal(new[] { 9.0 }, RangeHelper.ToDoubleArray(9.0));    // scalar -> [9]
            Assert.Contains("disabled", RangeHelper.DisabledMessage("Options"));
        }

        [Fact]
        public void In_Branches()
        {
            Assert.Throws<FinanceInputException>(() => In.Price("S", true));                     // bool
            Assert.Throws<FinanceInputException>(() => In.Rate("r", "xyz"));                     // unparseable
            Assert.Throws<FinanceInputException>(() => In.Num("x", double.PositiveInfinity));    // infinity
            Assert.Throws<FinanceInputException>(() => In.Rate("r", ""));                        // empty, no default
            Assert.Throws<FinanceInputException>(() => In.Years("T", double.NaN));               // NaN
            Assert.Throws<FinanceInputException>(() => In.Price("S", ExcelError.ExcelErrorNA));  // Excel error

            Assert.Equal(0.03, In.Rate("r", "", 0.03));                                          // empty -> default
            Assert.Equal(3.0, In.Num("x", ExcelMissing.Value, 3.0));                             // missing -> default
            Assert.Equal(7, In.Count("p", 7.0, 1));
            Assert.Equal(2.0, In.Matrix("m", new object[,] { { 1.0, 2.0 } })[0, 1]);             // matrix
        }

        [Fact]
        public void MoreBranches()
        {
            using var _ = new SettingsBackup();

            // RangeHelper branches not hit elsewhere
            Assert.Equal(new[] { 1.0, 2.0 }, RangeHelper.ToDoubleArray(new double[,] { { 1, 2 } })); // double[,]
            Assert.Equal(new[] { 5.0 }, RangeHelper.ToDoubleArray("5"));                             // string parse
            Assert.Equal(5.0, RangeHelper.ToDoubleMatrix("5")[0, 0]);                                // scalar fallback
            Assert.True(RangeHelper.ScalarBool(true));                                               // bool case
            Assert.True(In.Flag("b", "true"));                                                       // flag from text

            // Corrupt settings file -> Load() falls back to defaults (covers the catch).
            string path = System.IO.Path.Combine(
                System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData),
                "Aleksej.Finance.Excel", "settings.xml");
            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path)!);
            System.IO.File.WriteAllText(path, "not valid xml <<<");
            UserSettings.Invalidate();
            Assert.True(UserSettings.Load().EnableOptions);
        }
    }
}
