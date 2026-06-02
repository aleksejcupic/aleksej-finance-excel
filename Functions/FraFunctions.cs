using ExcelDna.Integration;
using AleksejCupic.FinancialMath.Derivatives;
using Aleksej.Finance.Excel.Helpers;
using Aleksej.Finance.Excel.Settings;

namespace Aleksej.Finance.Excel.Functions;

/// <summary>Forward Rate Agreement (FRA) pricing and settlement (Hull Ch. 4).</summary>
public static class FraFunctions
{
    private static bool Enabled => UserSettings.Load().EnableDerivatives;
    private static string Off   => RangeHelper.DisabledMessage("Derivatives");

    [ExcelFunction(Name = "FRA_RATE", Category = "Finance | Derivatives", IsThreadSafe = true,
        Description = "Continuously compounded forward rate for [t1, t2] from zero rates. f = (r2*t2 - r1*t1)/(t2-t1).")]
    public static object FraRate(
        [ExcelArgument(Name = "r1", Description = "Zero rate to t1 (continuous)")] object r1,
        [ExcelArgument(Name = "t1", Description = "Start of forward period (years)")] object t1,
        [ExcelArgument(Name = "r2", Description = "Zero rate to t2 (continuous)")] object r2,
        [ExcelArgument(Name = "t2", Description = "End of forward period (years)")] object t2)
        => Enabled ? ForwardRateAgreement.ForwardRate(
                         RangeHelper.Scalar(r1), RangeHelper.Scalar(t1), RangeHelper.Scalar(r2), RangeHelper.Scalar(t2))
                   : (object)Off;

    [ExcelFunction(Name = "FRA_RATE_SIMPLE", Category = "Finance | Derivatives", IsThreadSafe = true,
        Description = "Simply compounded (LIBOR/SOFR style) forward rate for [t1, t2]. R_F = (exp(f*(t2-t1)) - 1)/(t2-t1).")]
    public static object FraRateSimple(
        [ExcelArgument(Name = "r1", Description = "Zero rate to t1")] object r1,
        [ExcelArgument(Name = "t1", Description = "Start of forward period")] object t1,
        [ExcelArgument(Name = "r2", Description = "Zero rate to t2")] object r2,
        [ExcelArgument(Name = "t2", Description = "End of forward period")] object t2)
        => Enabled ? ForwardRateAgreement.ForwardRateSimple(
                         RangeHelper.Scalar(r1), RangeHelper.Scalar(t1), RangeHelper.Scalar(r2), RangeHelper.Scalar(t2))
                   : (object)Off;

    [ExcelFunction(Name = "FRA_VALUE", Category = "Finance | Derivatives", IsThreadSafe = true,
        Description = "Present value of an FRA. Long receives R_K (FRA rate) and pays market forward rate. V = L*(R_K - R_F)*delta*exp(-r2*t2).")]
    public static object FraValue(
        [ExcelArgument(Name = "notional", Description = "Notional principal")]                  object notional,
        [ExcelArgument(Name = "fraRate",  Description = "Agreed FRA rate R_K (simply compounded)")] object fraRate,
        [ExcelArgument(Name = "r1",       Description = "Current zero rate to t1")]             object r1,
        [ExcelArgument(Name = "t1",       Description = "Start of accrual period (years)")]     object t1,
        [ExcelArgument(Name = "r2",       Description = "Current zero rate to t2")]             object r2,
        [ExcelArgument(Name = "t2",       Description = "End of accrual period (years)")]       object t2,
        [ExcelArgument(Name = "isLong",   Description = "TRUE = long (receive fixed R_K), FALSE = short")] object isLong)
        => Enabled ? ForwardRateAgreement.FraValue(
                         RangeHelper.Scalar(notional), RangeHelper.Scalar(fraRate),
                         RangeHelper.Scalar(r1), RangeHelper.Scalar(t1),
                         RangeHelper.Scalar(r2), RangeHelper.Scalar(t2),
                         RangeHelper.IsMissing(isLong) ? true : RangeHelper.ScalarBool(isLong))
                   : (object)Off;

    [ExcelFunction(Name = "FRA_SETTLEMENT", Category = "Finance | Derivatives", IsThreadSafe = true,
        Description = "FRA settlement cash flow at t1. Settlement = L*(R_K - R_M)*delta / (1 + R_M*delta). Positive = long profits.")]
    public static object FraSettlement(
        [ExcelArgument(Name = "notional",   Description = "Notional principal")]                    object notional,
        [ExcelArgument(Name = "fraRate",    Description = "Agreed FRA rate R_K (simply compounded)")] object fraRate,
        [ExcelArgument(Name = "marketRate", Description = "Realised market rate R_M at settlement")] object marketRate,
        [ExcelArgument(Name = "t1",         Description = "Start of accrual period")]               object t1,
        [ExcelArgument(Name = "t2",         Description = "End of accrual period")]                 object t2,
        [ExcelArgument(Name = "isLong",     Description = "TRUE = long FRA")]                       object isLong)
        => Enabled ? ForwardRateAgreement.FraSettlement(
                         RangeHelper.Scalar(notional), RangeHelper.Scalar(fraRate),
                         RangeHelper.Scalar(marketRate),
                         RangeHelper.Scalar(t1), RangeHelper.Scalar(t2),
                         RangeHelper.IsMissing(isLong) ? true : RangeHelper.ScalarBool(isLong))
                   : (object)Off;

    [ExcelFunction(Name = "FRA_DV01", Category = "Finance | Derivatives", IsThreadSafe = true,
        Description = "FRA DV01 — change in value for a 1bp parallel shift in zero rates.")]
    public static object FraDv01(
        [ExcelArgument(Name = "notional", Description = "Notional principal")]           object notional,
        [ExcelArgument(Name = "fraRate",  Description = "Agreed FRA rate")]              object fraRate,
        [ExcelArgument(Name = "r1",       Description = "Zero rate to t1")]              object r1,
        [ExcelArgument(Name = "t1",       Description = "Start of accrual period")]      object t1,
        [ExcelArgument(Name = "r2",       Description = "Zero rate to t2")]              object r2,
        [ExcelArgument(Name = "t2",       Description = "End of accrual period")]        object t2,
        [ExcelArgument(Name = "isLong",   Description = "TRUE = long FRA")]              object isLong)
        => Enabled ? ForwardRateAgreement.DV01(
                         RangeHelper.Scalar(notional), RangeHelper.Scalar(fraRate),
                         RangeHelper.Scalar(r1), RangeHelper.Scalar(t1),
                         RangeHelper.Scalar(r2), RangeHelper.Scalar(t2),
                         RangeHelper.IsMissing(isLong) ? true : RangeHelper.ScalarBool(isLong))
                   : (object)Off;
}
