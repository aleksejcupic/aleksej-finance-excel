using System;

namespace Aleksej.Finance.Excel.Helpers;

/// <summary>
/// Thrown by the <see cref="In"/> input parsers when a user-supplied argument is
/// missing, unparseable, or outside its valid range. The <see cref="Exception.Message"/>
/// is end-user-facing (e.g. "sigma must be &gt; 0") and is surfaced in the cell by
/// <see cref="Fn.Run(Category, Func{object})"/>.
/// </summary>
public sealed class FinanceInputException : Exception
{
    public FinanceInputException(string message) : base(message) { }
}
