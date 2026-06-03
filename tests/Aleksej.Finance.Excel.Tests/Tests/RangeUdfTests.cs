using System;
using System.Collections.Generic;
using Xunit;
using Aleksej.Finance.Excel.Tests.Infra;

namespace Aleksej.Finance.Excel.Tests.Tests
{
    // UDFs that consume cell ranges (In.Vector / In.Matrix): risk metrics, portfolio,
    // yield-curve, IRS, Black-model caps/floors, attribution. Data blocks are laid out at
    // fixed cells; each case's formula references them.
    [ExcelTestSettings(AddIn = "Aleksej.Finance.ExcelAddin.xll")]
    public class RangeUdfTests
    {
        [ExcelFact]
        public void AllRangeUdfsEvaluateCorrectly()
        {
            dynamic ws = SheetHarness.NewSheet();
            LayoutData(ws);

            var cases = new[]
            {
                new UdfCase("SHARPE_RATIO",  "SHARPE_RATIO(H1:H10,0.02)", 0,0, Check.IsNum),
                new UdfCase("VAR_HISTORICAL","VAR_HISTORICAL(H1:H10,0.95)", 0,0, Check.IsNum),
                new UdfCase("VAR_CVAR",      "VAR_CVAR(H1:H10,0.95)", 0,0, Check.IsNum),
                new UdfCase("VAR_PARAMETRIC","VAR_PARAMETRIC(H1:H10,0.95)", 0,0, Check.IsNum),
                new UdfCase("ANN_RETURN",    "ANN_RETURN(S1:S3,10)", 0.20, 0.0001, Check.Num),
                new UdfCase("ANN_VOL",       "ANN_VOL(H1:H10,252)", 0,0, Check.Pos),
                new UdfCase("MAX_DRAWDOWN",  "MAX_DRAWDOWN(Q1:Q3)", 0.20, 0.0001, Check.Num),
                new UdfCase("RISK_SORTINO",  "RISK_SORTINO(H1:H10,0.02,252)", 0,0, Check.IsNum),
                new UdfCase("RISK_CALMAR",   "RISK_CALMAR(H1:H10,252)", 0,0, Check.IsNum),
                new UdfCase("RISK_BETA",     "RISK_BETA(H1:H10,H1:H10)", 1.0, 0.0001, Check.Num),
                new UdfCase("RISK_ALPHA",    "RISK_ALPHA(H1:H10,I1:I10,0.02,252)", 0,0, Check.IsNum),
                new UdfCase("RISK_TREYNOR",  "RISK_TREYNOR(H1:H10,I1:I10,0.02,252)", 0,0, Check.IsNum),
                new UdfCase("RISK_TE",       "RISK_TE(H1:H10,I1:I10,252)", 0,0, Check.Pos),
                new UdfCase("RISK_IR",       "RISK_IR(H1:H10,I1:I10,252)", 0,0, Check.IsNum),
                new UdfCase("VOL_EWMA_LATEST","VOL_EWMA_LATEST(H1:H10,0.94)", 0,0, Check.Pos),
                new UdfCase("PORT_RETURN",   "PORT_RETURN(K1:K3,L1:L3)", 0.112, 0.0001, Check.Num),
                new UdfCase("PORT_VOL",      "PORT_VOL(K1:K3,M1:O3)", 0,0, Check.Pos),
                new UdfCase("PORT_SHARPE",   "PORT_SHARPE(K1:K3,L1:L3,M1:O3,0.02)", 0,0, Check.IsNum),
                new UdfCase("YC_INTERPOLATE","YC_INTERPOLATE(X1:X4,Y1:Y4,1.0)", 0.032, 0.0001, Check.Num),
                new UdfCase("YC_PAR_YIELD",  "YC_PAR_YIELD(X1:X4,Y1:Y4,2,2)", 0,0, Check.Pos),
                new UdfCase("IRS_VALUE",     "IRS_VALUE(100,0.05,X1:X4,Y1:Y4,0.026,0.25,0.029,TRUE)", 0,0, Check.IsNum),
                new UdfCase("IRS_PAR_RATE",  "IRS_PAR_RATE(X1:X4,Y1:Y4,2)", 0,0, Check.Pos),
                new UdfCase("IRS_FIXED_LEG", "IRS_FIXED_LEG(100,0.05,X1:X4,Y1:Y4,2,FALSE)", 0,0, Check.Pos),
                new UdfCase("IRS_DV01",      "IRS_DV01(100,0.05,X1:X4,Y1:Y4,0.026,0.25,0.029,TRUE)", 0,0, Check.IsNum),
                new UdfCase("BM_CAP",        "BM_CAP(100,0.04,0.2,AK1:AK3,AL1:AL3,AM1:AM3,AN1:AN3)", 0,0, Check.Pos),
                new UdfCase("BM_FLOOR",      "BM_FLOOR(100,0.04,0.2,AK1:AK3,AL1:AL3,AM1:AM3,AN1:AN3)", 0,0, Check.Pos),
                new UdfCase("BM_FWD_SWAP_RATE","BM_FWD_SWAP_RATE(AK1:AK3,AL1:AL3,AN1:AN3)", 0,0, Check.Pos),
                new UdfCase("BM_SWAPTION",   "BM_SWAPTION(100,0.035,1,0.1,AK1:AK3,AL1:AL3,AN1:AN3,TRUE)", 0,0, Check.Pos),
                new UdfCase("FWD_PV_INCOME", "FWD_PV_INCOME(AP1:AP3,AQ1:AQ3,0.05)", 0,0, Check.Pos),
                new UdfCase("ATTR_TWR",      "ATTR_TWR(R1:R3)", 0.09725, 0.0001, Check.Num),
                new UdfCase("ATTR_MDIETZ",   "ATTR_MDIETZ(1000,1200,AS1,AT1,365)", 0.095238, 0.001, Check.Num),
                new UdfCase("ATTR_IRR",      "ATTR_IRR(U1:U2,V1:V2,0.1)", 0.09531, 0.0001, Check.Num),
                new UdfCase("ATTR_NPV",      "ATTR_NPV(U1:U2,V1:V2,0.05)", 46.35, 0.5, Check.Num),
                new UdfCase("ATTR_BHB_ALLOC","ATTR_BHB_ALLOC(AF1:AF2,AG1:AG2,AH1:AH2,AI1:AI2)", 0,0, Check.IsNum),
                new UdfCase("ATTR_BHB_SELECT","ATTR_BHB_SELECT(AF1:AF2,AG1:AG2,AH1:AH2,AI1:AI2)", 0,0, Check.IsNum),
                new UdfCase("ATTR_BHB_INTERACT","ATTR_BHB_INTERACT(AF1:AF2,AG1:AG2,AH1:AH2,AI1:AI2)", 0,0, Check.IsNum),
                new UdfCase("FEE_HWM",       "FEE_HWM(AA1:AA5)", 130, 0.0001, Check.Num),
                new UdfCase("EQ_PORT_VALUE", "EQ_PORT_VALUE(AC1:AC3,AD1:AD3)", 3000, 0.001, Check.Num),
            };

            var failures = new List<string>();
            foreach (var c in cases)
            {
                object res;
                try { res = SheetHarness.Eval(ws, c.Formula); }
                catch (Exception ex) { failures.Add($"{c.Func}: threw {ex.Message}"); continue; }
                if (!SheetHarness.Passes(res, c, out string msg)) failures.Add(msg);
            }
            Assert.True(failures.Count == 0,
                $"{failures.Count}/{cases.Length} range UDF(s) failed:\n" + string.Join("\n", failures));
        }

        internal static void LayoutData(dynamic ws)
        {
            SheetHarness.WriteColumn(ws, 1, 8,  new[]{0.012,-0.008,0.015,0.004,-0.011,0.020,-0.006,0.009,-0.003,0.014}); // H
            SheetHarness.WriteColumn(ws, 1, 9,  new[]{0.010,-0.005,0.012,0.003,-0.009,0.016,-0.004,0.007,-0.002,0.011}); // I
            SheetHarness.WriteColumn(ws, 1, 11, new[]{0.5,0.3,0.2});      // K  weights
            SheetHarness.WriteColumn(ws, 1, 12, new[]{0.08,0.12,0.18});   // L  mu
            SheetHarness.WriteBlock (ws, 1, 13, new[,]{{0.04,0.01,0.00},{0.01,0.06,0.01},{0.00,0.01,0.09}}); // M:O cov
            SheetHarness.WriteColumn(ws, 1, 17, new[]{0.10,-0.20,0.05});  // Q  drawdown series
            SheetHarness.WriteColumn(ws, 1, 18, new[]{0.10,0.05,-0.05});  // R  TWR series
            SheetHarness.WriteColumn(ws, 1, 19, new[]{0.01,0.02,0.03});   // S  ann-return series
            SheetHarness.WriteColumn(ws, 1, 21, new[]{-1000.0,1100.0});   // U  cash flows
            SheetHarness.WriteColumn(ws, 1, 22, new[]{0.0,1.0});          // V  times
            SheetHarness.WriteColumn(ws, 1, 24, new[]{0.5,1.0,1.5,2.0});  // X  maturities
            SheetHarness.WriteColumn(ws, 1, 25, new[]{0.03,0.032,0.034,0.035}); // Y  zeros
            SheetHarness.WriteColumn(ws, 1, 27, new[]{100.0,110.0,105.0,130.0,120.0}); // AA  NAV series
            SheetHarness.WriteColumn(ws, 1, 29, new[]{10.0,20.0,5.0});    // AC  positions
            SheetHarness.WriteColumn(ws, 1, 30, new[]{100.0,50.0,200.0}); // AD  prices
            SheetHarness.WriteColumn(ws, 1, 32, new[]{0.5,0.5});          // AF  port weights
            SheetHarness.WriteColumn(ws, 1, 33, new[]{0.4,0.6});          // AG  bench weights
            SheetHarness.WriteColumn(ws, 1, 34, new[]{0.10,0.04});        // AH  port returns
            SheetHarness.WriteColumn(ws, 1, 35, new[]{0.08,0.03});        // AI  bench returns
            SheetHarness.WriteColumn(ws, 1, 37, new[]{0.5,1.0,1.5});      // AK  payment times
            SheetHarness.WriteColumn(ws, 1, 38, new[]{0.03,0.032,0.034}); // AL  zeros
            SheetHarness.WriteColumn(ws, 1, 39, new[]{0.035,0.036,0.037});// AM  forwards
            SheetHarness.WriteColumn(ws, 1, 40, new[]{0.5,0.5,0.5});      // AN  accruals
            SheetHarness.WriteColumn(ws, 1, 42, new[]{2.0,2.0,2.0});      // AP  income cash flows
            SheetHarness.WriteColumn(ws, 1, 43, new[]{0.25,0.5,0.75});    // AQ  income times
            SheetHarness.WriteColumn(ws, 1, 45, new[]{100.0});            // AS  Dietz cash flow
            SheetHarness.WriteColumn(ws, 1, 46, new[]{182.5});            // AT  Dietz day
        }
    }
}
