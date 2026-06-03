using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ExcelDna.Integration;
using Aleksej.Finance.Excel.Settings;
using Xunit;

namespace Aleksej.Finance.Excel.UnitTests
{
    /// <summary>
    /// Invokes every UDF wrapper directly (no Excel) with valid inputs, exercising the
    /// function bodies, the Fn execution wrapper, the In validators, and RangeHelper.
    /// This is what lifts coverage of the add-in assembly from a few % to most of it.
    /// </summary>
    [Collection("AddIn")]
    public class FunctionCoverageTests
    {
        // ── range/matrix inputs (object[,] like a real Excel range) ──
        static object[,] Col(params double[] v)
        {
            var a = new object[v.Length, 1];
            for (int i = 0; i < v.Length; i++) a[i, 0] = v[i];
            return a;
        }
        static object[,] Mat(double[,] m)
        {
            var a = new object[m.GetLength(0), m.GetLength(1)];
            for (int i = 0; i < m.GetLength(0); i++)
                for (int j = 0; j < m.GetLength(1); j++) a[i, j] = m[i, j];
            return a;
        }

        static readonly object RET   = Col(0.012, -0.008, 0.015, 0.004, -0.011, 0.020, -0.006, 0.009, -0.003, 0.014);
        static readonly object BENCH = Col(0.010, -0.005, 0.012, 0.003, -0.009, 0.016, -0.004, 0.007, -0.002, 0.011);
        static readonly object W3    = Col(0.5, 0.3, 0.2);
        static readonly object MU3   = Col(0.08, 0.12, 0.18);
        static readonly object COV3  = Mat(new[,] { { 0.04, 0.01, 0.00 }, { 0.01, 0.06, 0.01 }, { 0.00, 0.01, 0.09 } });
        static readonly object MAT4  = Col(0.5, 1.0, 1.5, 2.0);
        static readonly object ZERO4 = Col(0.03, 0.032, 0.034, 0.035);
        static readonly object NAV5  = Col(100, 110, 105, 130, 120);
        static readonly object POS3  = Col(10, 20, 5);
        static readonly object PRC3  = Col(100, 50, 200);
        static readonly object BPW   = Col(0.5, 0.5);
        static readonly object BBW   = Col(0.4, 0.6);
        static readonly object BPR   = Col(0.10, 0.04);
        static readonly object BBR   = Col(0.08, 0.03);
        static readonly object BMT   = Col(0.5, 1.0, 1.5);
        static readonly object BMZ   = Col(0.03, 0.032, 0.034);
        static readonly object BMF   = Col(0.035, 0.036, 0.037);
        static readonly object BMA   = Col(0.5, 0.5, 0.5);
        static readonly object PVCF  = Col(2, 2, 2);
        static readonly object PVT   = Col(0.25, 0.5, 0.75);
        static readonly object CF2   = Col(-1000, 1100);
        static readonly object TIM2  = Col(0, 1);
        static readonly object MDCF  = Col(100);
        static readonly object MDD   = Col(182.5);

        static object[] A(params object[] args) => args;

        static Dictionary<string, object[]> ValidArgs() => new()
        {
            // Black-Scholes
            ["BS_CALL"] = A(100.0, 100.0, 1.0, 0.05, 0.2),
            ["BS_PUT"] = A(100.0, 100.0, 1.0, 0.05, 0.2),
            ["BS_DELTA"] = A(100.0, 100.0, 1.0, 0.05, 0.2, false),
            ["BS_GAMMA"] = A(100.0, 100.0, 1.0, 0.05, 0.2),
            ["BS_VEGA"] = A(100.0, 100.0, 1.0, 0.05, 0.2),
            ["BS_THETA"] = A(100.0, 100.0, 1.0, 0.05, 0.2, false),
            ["BS_RHO"] = A(100.0, 100.0, 1.0, 0.05, 0.2, false),
            ["BS_IV"] = A(10.4506, 100.0, 100.0, 1.0, 0.05, false),
            ["BS_VANNA"] = A(100.0, 100.0, 1.0, 0.05, 0.2),
            ["BS_CHARM"] = A(100.0, 100.0, 1.0, 0.05, 0.2),
            ["BS_VOLGA"] = A(100.0, 100.0, 1.0, 0.05, 0.2),
            ["BS_SPEED"] = A(100.0, 100.0, 1.0, 0.05, 0.2),
            ["BS_ZOMMA"] = A(100.0, 100.0, 1.0, 0.05, 0.2),
            // Binomial
            ["BT_PRICE"] = A(100.0, 100.0, 1.0, 0.05, 0.2, 200.0, false, false),
            ["BT_DELTA"] = A(100.0, 100.0, 1.0, 0.05, 0.2, 200.0, false, false),
            ["BT_GAMMA"] = A(100.0, 100.0, 1.0, 0.05, 0.2, 200.0, false, false),
            // Exotic
            ["EX_BINARY_CASH"] = A(100.0, 100.0, 1.0, 0.05, 0.2, 10.0, false),
            ["EX_BINARY_ASSET"] = A(100.0, 100.0, 1.0, 0.05, 0.2, false),
            ["EX_BARRIER_CALL"] = A(100.0, 100.0, 90.0, 1.0, 0.05, 0.2, false, false),
            ["EX_BARRIER_PUT"] = A(100.0, 100.0, 90.0, 1.0, 0.05, 0.2, false, false),
            ["EX_ASIAN_GEO"] = A(100.0, 100.0, 1.0, 0.05, 0.2, false),
            ["EX_ASIAN_ARITH"] = A(100.0, 100.0, 1.0, 0.05, 0.2, 12.0, 2000.0, false, 42.0),
            ["EX_LOOKBACK_CALL"] = A(100.0, 100.0, 1.0, 0.05, 0.2),
            ["EX_LOOKBACK_PUT"] = A(100.0, 100.0, 1.0, 0.05, 0.2),
            // Monte Carlo (async)
            ["MC_EUROPEAN"] = A(100.0, 100.0, 1.0, 0.05, 0.2, 1000.0, 50.0, false, 42.0),
            ["MC_AMERICAN"] = A(100.0, 100.0, 1.0, 0.05, 0.2, 1000.0, 50.0, true, 42.0),
            // Garman-Kohlhagen
            ["GK_CALL"] = A(1.2, 1.25, 1.0, 0.05, 0.03, 0.15),
            ["GK_PUT"] = A(1.2, 1.25, 1.0, 0.05, 0.03, 0.15),
            ["GK_DELTA"] = A(1.2, 1.25, 1.0, 0.05, 0.03, 0.15, false),
            ["GK_GAMMA"] = A(1.2, 1.25, 1.0, 0.05, 0.03, 0.15),
            ["GK_VEGA"] = A(1.2, 1.25, 1.0, 0.05, 0.03, 0.15),
            ["GK_THETA"] = A(1.2, 1.25, 1.0, 0.05, 0.03, 0.15, false),
            ["GK_RHO"] = A(1.2, 1.25, 1.0, 0.05, 0.03, 0.15, false),
            ["GK_RHO_FOREIGN"] = A(1.2, 1.25, 1.0, 0.05, 0.03, 0.15, false),
            ["GK_IV"] = A(0.05, 1.2, 1.25, 1.0, 0.05, 0.03, false),
            // Options on futures
            ["OF_CALL"] = A(100.0, 100.0, 1.0, 0.05, 0.2),
            ["OF_PUT"] = A(100.0, 100.0, 1.0, 0.05, 0.2),
            ["OF_CALL_FROM_PUT"] = A(7.5, 100.0, 100.0, 1.0, 0.05),
            ["OF_DELTA"] = A(100.0, 100.0, 1.0, 0.05, 0.2, false),
            ["OF_GAMMA"] = A(100.0, 100.0, 1.0, 0.05, 0.2),
            ["OF_VEGA"] = A(100.0, 100.0, 1.0, 0.05, 0.2),
            ["OF_IV"] = A(6.0, 100.0, 100.0, 1.0, 0.05, false),
            // Bonds
            ["BOND_PRICE"] = A(1000.0, 0.05, 0.05, 10.0, 2.0),
            ["BOND_YTM"] = A(1000.0, 1000.0, 0.05, 10.0, 2.0),
            ["BOND_DURATION"] = A(1000.0, 0.05, 0.05, 10.0, 2.0),
            ["BOND_MOD_DURATION"] = A(1000.0, 0.05, 0.05, 10.0, 2.0),
            ["BOND_CONVEXITY"] = A(1000.0, 0.05, 0.05, 10.0, 2.0),
            ["BOND_DV01"] = A(1000.0, 0.05, 0.05, 10.0, 2.0),
            ["BOND_PRICE_CHANGE"] = A(1000.0, 0.05, 0.05, 10.0, 0.001, 2.0),
            ["YC_DF"] = A(0.05, 3.0),
            ["YC_TO_CONT"] = A(0.06, 2.0),
            ["YC_FROM_CONT"] = A(0.06, 2.0),
            ["YC_FWD_RATE"] = A(0.05, 1.0, 0.05, 2.0),
            ["YC_INTERPOLATE"] = A(MAT4, ZERO4, 1.0),
            ["YC_PAR_YIELD"] = A(MAT4, ZERO4, 2.0, 2.0),
            ["MORT_PAYMENT"] = A(100000.0, 0.06, 30.0, 12.0),
            ["MORT_BALANCE"] = A(100000.0, 0.06, 30.0, 0.0, 12.0),
            ["MORT_TOTAL_INTEREST"] = A(100000.0, 0.06, 30.0, 12.0),
            ["MORT_EAR"] = A(0.06, 12.0),
            // Forwards/Futures
            ["FWD_PRICE"] = A(100.0, 0.05, 1.0),
            ["FWD_PRICE_YIELD"] = A(100.0, 0.05, 0.02, 1.0),
            ["FWD_PRICE_INCOME"] = A(100.0, 5.0, 0.05, 1.0),
            ["FWD_FX"] = A(1.2, 0.04, 0.01, 0.5),
            ["FWD_COMMODITY"] = A(50.0, 0.03, 0.02, 0.01, 1.0),
            ["FWD_VALUE"] = A(110.0, 100.0, 0.05, 1.0),
            ["FWD_VALUE_SHORT"] = A(110.0, 100.0, 0.05, 1.0),
            ["FWD_PV_INCOME"] = A(PVCF, PVT, 0.05),
            // FRA
            ["FRA_RATE"] = A(0.03, 1.0, 0.035, 2.0),
            ["FRA_RATE_SIMPLE"] = A(0.03, 1.0, 0.035, 2.0),
            ["FRA_VALUE"] = A(1000000.0, 0.05, 0.03, 1.0, 0.035, 2.0, true),
            ["FRA_SETTLEMENT"] = A(1000000.0, 0.05, 0.06, 1.0, 2.0, true),
            ["FRA_DV01"] = A(1000000.0, 0.05, 0.03, 1.0, 0.035, 2.0, true),
            // IRS
            ["IRS_VALUE"] = A(100.0, 0.05, MAT4, ZERO4, 0.026, 0.25, 0.029, true),
            ["IRS_PAR_RATE"] = A(MAT4, ZERO4, 2.0),
            ["IRS_FIXED_LEG"] = A(100.0, 0.05, MAT4, ZERO4, 2.0, false),
            ["IRS_FLOAT_LEG"] = A(100.0, 0.026, 0.25, 0.03),
            ["IRS_DV01"] = A(100.0, 0.05, MAT4, ZERO4, 0.026, 0.25, 0.029, true),
            // Black model
            ["BM_CAPLET"] = A(100.0, 0.04, 0.04, 1.0, 0.03, 0.2, 0.5, false),
            ["BM_CAP"] = A(100.0, 0.04, 0.2, BMT, BMZ, BMF, BMA),
            ["BM_FLOOR"] = A(100.0, 0.04, 0.2, BMT, BMZ, BMF, BMA),
            ["BM_FWD_SWAP_RATE"] = A(BMT, BMZ, BMA),
            ["BM_SWAPTION"] = A(100.0, 0.035, 1.0, 0.1, BMT, BMZ, BMA, true),
            // Short rate
            ["SR_VASICEK_PRICE"] = A(0.03, 5.0, 0.3, 0.05, 0.02),
            ["SR_VASICEK_YIELD"] = A(0.03, 5.0, 0.3, 0.05, 0.02),
            ["SR_VASICEK_LRYIELD"] = A(0.3, 0.05, 0.02),
            ["SR_VASICEK_OPTION"] = A(0.03, 1.0, 5.0, 0.8, 0.3, 0.05, 0.02, false),
            ["SR_CIR_PRICE"] = A(0.03, 5.0, 0.3, 0.05, 0.02),
            ["SR_CIR_YIELD"] = A(0.03, 5.0, 0.3, 0.05, 0.02),
            ["SR_CIR_LRYIELD"] = A(0.3, 0.05, 0.02),
            // Credit
            ["CR_MERTON_EQUITY"] = A(100.0, 80.0, 1.0, 0.05, 0.25),
            ["CR_MERTON_DEBT"] = A(100.0, 80.0, 1.0, 0.05, 0.25),
            ["CR_DEFAULT_PROB"] = A(100.0, 80.0, 1.0, 0.05, 0.25),
            ["CR_DIST_TO_DEFAULT"] = A(100.0, 80.0, 1.0, 0.05, 0.25),
            ["CR_CREDIT_SPREAD"] = A(100.0, 80.0, 1.0, 0.05, 0.25),
            ["CR_SURVIVAL_PROB"] = A(0.02, 5.0),
            ["CR_HAZARD_FROM_SPREAD"] = A(0.012, 0.4),
            ["CR_CDS_SPREAD"] = A(0.02, 0.03, 5.0, 0.4, 4.0),
            ["CR_CDS_MTM"] = A(0.005, 0.03, 0.03, 10000000.0, 5.0, 0.4, 4.0),
            // Portfolio
            ["PORT_RETURN"] = A(W3, MU3),
            ["PORT_VOL"] = A(W3, COV3),
            ["PORT_SHARPE"] = A(W3, MU3, COV3, 0.02),
            ["PORT_MIN_VAR"] = A(COV3),
            ["PORT_MAX_SHARPE"] = A(MU3, COV3, 0.02),
            ["PORT_RISK_PARITY"] = A(COV3),
            ["PORT_RISK_CONTRIB"] = A(W3, COV3),
            // Risk
            ["SHARPE_RATIO"] = A(RET, 0.02, 252.0),
            ["VAR_HISTORICAL"] = A(RET, 0.95),
            ["VAR_CVAR"] = A(RET, 0.95),
            ["VAR_PARAMETRIC"] = A(RET, 0.95),
            ["ANN_RETURN"] = A(RET, 252.0),
            ["ANN_VOL"] = A(RET, 252.0),
            ["MAX_DRAWDOWN"] = A(RET),
            ["RISK_SORTINO"] = A(RET, 0.02, 252.0),
            ["RISK_CALMAR"] = A(RET, 252.0),
            ["RISK_BETA"] = A(RET, BENCH),
            ["RISK_ALPHA"] = A(RET, BENCH, 0.02, 252.0),
            ["RISK_TREYNOR"] = A(RET, BENCH, 0.02, 252.0),
            ["RISK_TE"] = A(RET, BENCH, 252.0),
            ["RISK_IR"] = A(RET, BENCH, 252.0),
            ["VOL_EWMA_LATEST"] = A(RET, 0.94),
            ["VOL_GARCH_LONGRUN"] = A(0.000002, 0.05, 0.9),
            ["VOL_GARCH_FORECAST"] = A(0.0002, 0.000002, 0.05, 0.9, 10.0),
            // Fees
            ["FEE_MGMT"] = A(100000000.0, 0.02, 90.0, 365.0),
            ["FEE_PERF"] = A(120.0, 100.0, 100.0, 0.2, 0.0),
            ["FEE_HWM"] = A(NAV5),
            ["FEE_EXPENSE_DRAG"] = A(0.1, 0.01, 5.0),
            ["FEE_NET_RETURN"] = A(0.1, 0.01),
            ["FEE_CARRIED_INT"] = A(1280.0, 1000.0, 0.08, 1.0, 0.2),
            ["FEE_TRANSACTION_COST"] = A(1000000.0, 0.001, 4.0),
            // Attribution
            ["ATTR_TWR"] = A(Col(0.10, 0.05, -0.05)),
            ["ATTR_MDIETZ"] = A(1000.0, 1200.0, MDCF, MDD, 365.0),
            ["ATTR_IRR"] = A(CF2, TIM2, 0.1),
            ["ATTR_NPV"] = A(CF2, TIM2, 0.05),
            ["ATTR_ALLOC"] = A(0.5, 0.4, 0.05, 0.04),
            ["ATTR_SELECT"] = A(0.4, 0.1, 0.08),
            ["ATTR_INTERACT"] = A(0.5, 0.4, 0.1, 0.08),
            ["ATTR_BHB_ALLOC"] = A(BPW, BBW, BPR, BBR),
            ["ATTR_BHB_SELECT"] = A(BPW, BBW, BPR, BBR),
            ["ATTR_BHB_INTERACT"] = A(BPW, BBW, BPR, BBR),
            ["ATTR_ACTIVE_RETURN"] = A(0.07, 0.05),
            // Equity
            ["EQ_MKTCAP"] = A(10000.0, 100.0),
            ["EQ_EV"] = A(1000000.0, 300000.0, 100000.0),
            ["EQ_PORT_VALUE"] = A(POS3, PRC3),
            ["EQ_PE"] = A(100.0, 5.0),
            ["EQ_PB"] = A(50.0, 20.0),
            ["EQ_PS"] = A(2000.0, 500.0),
            ["EQ_EVTOEBITDA"] = A(1000.0, 100.0),
            ["EQ_DIV_YIELD"] = A(2.0, 50.0),
            ["EQ_UNREAL_PNL"] = A(10.0, 100.0, 120.0),
            ["EQ_REAL_PNL"] = A(10.0, 100.0, 90.0),
            ["EQ_KELLY"] = A(0.1, 0.2),
            ["EQ_HALF_KELLY"] = A(0.1, 0.2),
        };

        static Dictionary<string, MethodInfo> BuildNameToMethod()
        {
            var map = new Dictionary<string, MethodInfo>();
            Assembly asm = typeof(Aleksej.Finance.Excel.Functions.OptionsFunctions).Assembly;
            foreach (Type t in asm.GetTypes())
            {
                if (t.Namespace != "Aleksej.Finance.Excel.Functions") continue;
                foreach (MethodInfo m in t.GetMethods(BindingFlags.Public | BindingFlags.Static))
                {
                    var attr = m.GetCustomAttribute<ExcelFunctionAttribute>();
                    if (attr?.Name != null) map[attr.Name] = m;
                }
            }
            return map;
        }

        [Fact]
        public void EveryUdfComputesWithValidInputs()
        {
            UserSettings.UseForTesting(new UserSettings());   // all categories enabled, text errors
            var map = BuildNameToMethod();
            var args = ValidArgs();
            var failures = new List<string>();

            foreach (var kv in args)
            {
                if (!map.TryGetValue(kv.Key, out MethodInfo? mi)) { failures.Add($"{kv.Key}: method not found"); continue; }
                object? result;
                try { result = mi.Invoke(null, kv.Value); }
                catch (Exception ex)
                {
                    // Async (Monte Carlo) functions go through ExcelAsyncUtil, which needs the
                    // Excel calc context; tolerate them throwing outside Excel.
                    if (kv.Key.StartsWith("MC_")) continue;
                    failures.Add($"{kv.Key}: threw {ex.InnerException?.Message ?? ex.Message}");
                    continue;
                }
                if (result is string s) failures.Add($"{kv.Key}: returned message '{s}'");
                else if (result is null) failures.Add($"{kv.Key}: returned null");
            }

            // Every discovered UDF must have a test case (catches new functions added later).
            var untested = map.Keys.Except(args.Keys).ToList();
            Assert.True(failures.Count == 0 && untested.Count == 0,
                $"{failures.Count} failed:\n{string.Join("\n", failures)}" +
                (untested.Count > 0 ? $"\nUDFs with no test args: {string.Join(", ", untested)}" : ""));
        }
    }
}
