using Aleksej.Finance.Excel.Helpers;
using Xunit;

namespace Aleksej.Finance.Excel.UnitTests
{
    /// <summary>Pure tests of the Excel-range conversion helpers (no Excel required).</summary>
    public class RangeHelperTests
    {
        [Fact]
        public void ToDoubleArray_FlattensBlockRowMajor()
            => Assert.Equal(new[] { 1.0, 2.0, 3.0, 4.0 },
                            RangeHelper.ToDoubleArray(new object[,] { { 1.0, 2.0 }, { 3.0, 4.0 } }));

        [Fact]
        public void ToDoubleArray_SkipsEmptyAndError()
            => Assert.Equal(new[] { 1.0, 2.0 },
                            RangeHelper.ToDoubleArray(new object[,] { { 1.0, ExcelDna.Integration.ExcelEmpty.Value }, { 2.0, ExcelDna.Integration.ExcelError.ExcelErrorNA } }));

        [Fact]
        public void ToDoubleMatrix_PreservesShape()
        {
            double[,] m = RangeHelper.ToDoubleMatrix(new object[,] { { 1.0, 2.0 }, { 3.0, 4.0 } });
            Assert.Equal(2, m.GetLength(0));
            Assert.Equal(2, m.GetLength(1));
            Assert.Equal(3.0, m[1, 0]);
        }

        [Fact] public void Scalar_ParsesNumericString() => Assert.Equal(1.5, RangeHelper.Scalar("1.5"));
        [Fact] public void Scalar_UsesDefaultOnMissing() => Assert.Equal(7.0, RangeHelper.Scalar(ExcelDna.Integration.ExcelMissing.Value, 7.0));
        [Fact] public void ScalarBool_ParsesText() => Assert.True(RangeHelper.ScalarBool("true"));
        [Fact] public void ScalarInt_Rounds() => Assert.Equal(3, RangeHelper.ScalarInt(2.6));
    }
}
