namespace Aleksej.Finance.Excel.Tests.Infra
{
    /// <summary>How a UDF result is judged.</summary>
    public enum Check { Num, Pos, Neg, In01, IsNum, Text }

    /// <summary>
    /// One end-to-end case: a literal Excel formula (without '='), the expected value,
    /// a tolerance, and how to judge it. Mirrors the self-checking workbook's cases.
    /// </summary>
    public sealed class UdfCase
    {
        public readonly string Func;
        public readonly string Formula;
        public readonly double Expected;
        public readonly double Tol;
        public readonly string? ExpectedText;
        public readonly Check Check;

        public UdfCase(string func, string formula, double expected, double tol, Check check)
        {
            Func = func; Formula = formula; Expected = expected; Tol = tol; Check = check;
        }

        // Text (error-message) cases.
        public UdfCase(string func, string formula, string expectedText)
        {
            Func = func; Formula = formula; ExpectedText = expectedText; Check = Check.Text;
        }

        public override string ToString() => Func;
    }
}
