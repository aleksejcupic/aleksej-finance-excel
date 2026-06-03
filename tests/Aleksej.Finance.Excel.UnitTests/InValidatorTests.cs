using ExcelDna.Integration;
using Aleksej.Finance.Excel.Helpers;
using Xunit;

namespace Aleksej.Finance.Excel.UnitTests
{
    /// <summary>
    /// Pure (no-Excel) tests of the input validation + normalization layer. These run
    /// headless in CI and produce coverage of the add-in's own logic.
    /// </summary>
    public class InValidatorTests
    {
        // ── Price (> 0) ──
        [Fact] public void Price_AcceptsPositive() => Assert.Equal(100.0, In.Price("S", 100.0));
        [Fact] public void Price_ParsesNumericString() => Assert.Equal(100.0, In.Price("S", "100"));
        [Fact] public void Price_RejectsZero() => Assert.Throws<FinanceInputException>(() => In.Price("S", 0.0));
        [Fact] public void Price_RejectsNegative() => Assert.Throws<FinanceInputException>(() => In.Price("S", -5.0));
        [Fact] public void Price_RejectsMissing() => Assert.Throws<FinanceInputException>(() => In.Price("S", ExcelMissing.Value));

        // ── Vol (> 0) ──
        [Fact] public void Vol_AcceptsPositive() => Assert.Equal(0.2, In.Vol("sigma", 0.2));
        [Fact] public void Vol_RejectsZero() => Assert.Throws<FinanceInputException>(() => In.Vol("sigma", 0.0));

        // ── Years (>= 0) ──
        [Fact] public void Years_AcceptsZero() => Assert.Equal(0.0, In.Years("T", 0.0));
        [Fact] public void Years_RejectsNegative() => Assert.Throws<FinanceInputException>(() => In.Years("T", -1.0));

        // ── Prob ([0,1]) ──
        [Fact] public void Prob_AcceptsInRange() => Assert.Equal(0.95, In.Prob("c", 0.95));
        [Fact] public void Prob_RejectsAboveOne() => Assert.Throws<FinanceInputException>(() => In.Prob("c", 1.5));
        [Fact] public void Prob_RejectsNegative() => Assert.Throws<FinanceInputException>(() => In.Prob("c", -0.1));
        [Fact] public void Prob_UsesDefaultOnMissing() => Assert.Equal(0.95, In.Prob("c", ExcelMissing.Value, 0.95));

        // ── PosInt (>= 1) / Count (>= 0) ──
        [Fact] public void PosInt_AcceptsAndRounds() => Assert.Equal(5, In.PosInt("n", 5.0));
        [Fact] public void PosInt_RejectsZero() => Assert.Throws<FinanceInputException>(() => In.PosInt("n", 0.0));
        [Fact] public void PosInt_UsesDefaultOnMissing() => Assert.Equal(10, In.PosInt("n", ExcelMissing.Value, 10));
        [Fact] public void Count_AcceptsZero() => Assert.Equal(0, In.Count("p", 0.0));
        [Fact] public void Count_RejectsNegative() => Assert.Throws<FinanceInputException>(() => In.Count("p", -1.0));

        // ── Rate (any finite) + percent normalization ──
        [Fact] public void Rate_AcceptsNegative() => Assert.Equal(-0.01, In.Rate("r", -0.01));
        [Fact] public void Rate_NumericIsUnchanged() => Assert.Equal(0.05, In.Rate("r", 0.05), 10);   // not halved/divided
        [Fact] public void Rate_PercentTextNormalized() => Assert.Equal(0.05, In.Rate("r", "5%"), 10); // text "5%" -> 0.05
        [Fact] public void Rate_UsesDefaultOnMissing() => Assert.Equal(0.03, In.Rate("r", ExcelMissing.Value, 0.03));
        [Fact] public void Rate_ThrowsWhenMissingNoDefault() => Assert.Throws<FinanceInputException>(() => In.Rate("r", ExcelMissing.Value));

        // ── Num (finite, signed) ──
        [Fact] public void Num_AcceptsNegative() => Assert.Equal(-3.5, In.Num("x", -3.5));
        [Fact] public void Num_RejectsNaN() => Assert.Throws<FinanceInputException>(() => In.Num("x", double.NaN));

        // ── Flag ──
        [Fact] public void Flag_TrueFalse() { Assert.True(In.Flag("b", true)); Assert.False(In.Flag("b", false)); }
        [Fact] public void Flag_FromNumber() { Assert.True(In.Flag("b", 1.0)); Assert.False(In.Flag("b", 0.0)); }
        [Fact] public void Flag_DefaultOnMissing() => Assert.True(In.Flag("b", ExcelMissing.Value, true));

        // ── Excel errors flowing in must be rejected (not silently NaN) ──
        [Fact] public void ExcelError_Throws() => Assert.Throws<FinanceInputException>(() => In.Rate("r", ExcelError.ExcelErrorNA));

        // ── Vector / Matrix ──
        [Fact]
        public void Vector_FromColumnRange()
            => Assert.Equal(new[] { 1.0, 2.0, 3.0 }, In.Vector("v", new object[,] { { 1.0 }, { 2.0 }, { 3.0 } }));

        [Fact]
        public void Vector_RejectsEmpty()
            => Assert.Throws<FinanceInputException>(() => In.Vector("v", new object[,] { { ExcelEmpty.Value } }));

        [Fact]
        public void Matrix_PreservesShape()
        {
            double[,] m = In.Matrix("cov", new object[,] { { 1.0, 2.0 }, { 3.0, 4.0 } });
            Assert.Equal(2, m.GetLength(0));
            Assert.Equal(2, m.GetLength(1));
            Assert.Equal(4.0, m[1, 1]);
        }
    }
}
