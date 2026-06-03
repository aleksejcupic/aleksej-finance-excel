using Aleksej.Finance.Excel.Tests.Infra;

namespace Aleksej.Finance.Excel.Tests.Data
{
    /// <summary>
    /// Every scalar-input UDF as a literal Excel formula. Golden values are reused from the
    /// math library's own xUnit suite; where an exact value isn't pinned, a structural check
    /// (positive / negative / in [0,1] / is-a-number) is used.
    /// </summary>
    public static class ScalarCases
    {
        public static readonly UdfCase[] All =
        {
            // ── Black-Scholes ──
            new("BS_CALL",  "BS_CALL(100,100,1,0.05,0.2)",  10.4506, 0.001, Check.Num),
            new("BS_PUT",   "BS_PUT(100,100,1,0.05,0.2)",    5.5735, 0.001, Check.Num),
            new("BS_DELTA", "BS_DELTA(100,100,1,0.05,0.2,FALSE)", 0.6368, 0.001, Check.Num),
            new("BS_GAMMA", "BS_GAMMA(100,100,1,0.05,0.2)", 0.01876, 0.001, Check.Num),
            new("BS_VEGA",  "BS_VEGA(100,100,1,0.05,0.2)",   0.3752, 0.001, Check.Num),
            new("BS_THETA", "BS_THETA(100,100,1,0.05,0.2,FALSE)", 0, 0, Check.Neg),
            new("BS_RHO",   "BS_RHO(100,100,1,0.05,0.2,FALSE)",   0, 0, Check.Pos),
            new("BS_IV",    "BS_IV(10.4506,100,100,1,0.05,FALSE)", 0.20, 0.001, Check.Num),
            new("BS_VANNA", "BS_VANNA(100,100,1,0.05,0.2)", 0, 0, Check.IsNum),
            new("BS_CHARM", "BS_CHARM(100,100,1,0.05,0.2)", 0, 0, Check.IsNum),
            new("BS_VOLGA", "BS_VOLGA(100,100,1,0.05,0.2)", 0, 0, Check.IsNum),
            new("BS_SPEED", "BS_SPEED(100,100,1,0.05,0.2)", 0, 0, Check.IsNum),
            new("BS_ZOMMA", "BS_ZOMMA(100,100,1,0.05,0.2)", 0, 0, Check.IsNum),

            // ── Binomial Tree ──
            new("BT_PRICE", "BT_PRICE(100,100,1,0.05,0.2,300,FALSE,FALSE)", 10.4506, 0.1, Check.Num),
            new("BT_DELTA", "BT_DELTA(100,100,1,0.05,0.2,300,FALSE,FALSE)", 0, 0, Check.In01),
            new("BT_GAMMA", "BT_GAMMA(100,100,1,0.05,0.2,300,FALSE)", 0, 0, Check.Pos),

            // ── Exotic ──
            new("EX_BINARY_CASH",  "EX_BINARY_CASH(100,100,1,0.05,0.2,10,FALSE)", 0, 0, Check.Pos),
            new("EX_BINARY_ASSET", "EX_BINARY_ASSET(100,100,1,0.05,0.2,FALSE)", 0, 0, Check.Pos),
            new("EX_BARRIER_CALL", "EX_BARRIER_CALL(100,100,90,1,0.05,0.2,FALSE,FALSE)", 0, 0, Check.Pos),
            new("EX_BARRIER_PUT",  "EX_BARRIER_PUT(100,100,90,1,0.05,0.2,FALSE,FALSE)", 0, 0, Check.Pos),
            new("EX_ASIAN_GEO",    "EX_ASIAN_GEO(100,100,1,0.05,0.2,FALSE)", 0, 0, Check.Pos),
            new("EX_ASIAN_ARITH",  "EX_ASIAN_ARITH(100,100,1,0.05,0.2,12,20000,FALSE,42)", 0, 0, Check.Pos),
            new("EX_LOOKBACK_CALL","EX_LOOKBACK_CALL(100,100,1,0.05,0.2)", 0, 0, Check.Pos),
            new("EX_LOOKBACK_PUT", "EX_LOOKBACK_PUT(100,100,1,0.05,0.2)", 0, 0, Check.Pos),

            // ── Garman-Kohlhagen ──
            new("GK_CALL", "GK_CALL(1.2,1.25,1,0.05,0.03,0.15)", 0, 0, Check.Pos),
            new("GK_PUT",  "GK_PUT(1.2,1.25,1,0.05,0.03,0.15)", 0, 0, Check.Pos),
            new("GK_DELTA","GK_DELTA(1.2,1.25,1,0.05,0.03,0.15,FALSE)", 0, 0, Check.In01),
            new("GK_GAMMA","GK_GAMMA(1.2,1.25,1,0.05,0.03,0.15)", 0, 0, Check.Pos),
            new("GK_VEGA", "GK_VEGA(1.2,1.25,1,0.05,0.03,0.15)", 0, 0, Check.Pos),
            new("GK_THETA","GK_THETA(1.2,1.25,1,0.05,0.03,0.15,FALSE)", 0, 0, Check.IsNum),
            new("GK_RHO",  "GK_RHO(1.2,1.25,1,0.05,0.03,0.15,FALSE)", 0, 0, Check.IsNum),
            new("GK_RHO_FOREIGN","GK_RHO_FOREIGN(1.2,1.25,1,0.05,0.03,0.15,FALSE)", 0, 0, Check.IsNum),
            new("GK_IV",   "GK_IV(GK_CALL(1.2,1.25,1,0.05,0.03,0.15),1.2,1.25,1,0.05,0.03,FALSE)", 0.15, 0.001, Check.Num),

            // ── Options on Futures ──
            new("OF_CALL", "OF_CALL(100,100,1,0.05,0.2)", 0, 0, Check.Pos),
            new("OF_PUT",  "OF_PUT(100,100,1,0.05,0.2)", 0, 0, Check.Pos),
            new("OF_CALL_FROM_PUT", "OF_CALL_FROM_PUT(7.5,100,100,1,0.05)", 7.5, 0.000001, Check.Num), // ATM parity: call=put
            new("OF_DELTA","OF_DELTA(100,100,1,0.05,0.2,FALSE)", 0, 0, Check.In01),
            new("OF_GAMMA","OF_GAMMA(100,100,1,0.05,0.2)", 0, 0, Check.Pos),
            new("OF_VEGA", "OF_VEGA(100,100,1,0.05,0.2)", 0, 0, Check.Pos),
            new("OF_IV",   "OF_IV(OF_CALL(100,100,1,0.05,0.2),100,100,1,0.05,FALSE)", 0.20, 0.001, Check.Num),

            // ── Bonds ──
            new("BOND_PRICE", "BOND_PRICE(1000,0.05,0.05,10)", 1000, 0.001, Check.Num),
            new("BOND_YTM",   "BOND_YTM(1000,1000,0.05,10)", 0.05, 0.0001, Check.Num),
            new("BOND_DURATION", "BOND_DURATION(1000,0.05,0.05,10,2)", 0, 0, Check.Pos),
            new("BOND_MOD_DURATION", "BOND_MOD_DURATION(1000,0.05,0.05,10,2)", 0, 0, Check.Pos),
            new("BOND_CONVEXITY", "BOND_CONVEXITY(1000,0.05,0.05,10,2)", 0, 0, Check.Pos),
            new("BOND_DV01",  "BOND_DV01(1000,0.05,0.05,10,2)", 0, 0, Check.Pos),
            new("BOND_PRICE_CHANGE", "BOND_PRICE_CHANGE(1000,0.05,0.05,10,0.001,2)", 0, 0, Check.Neg),
            new("YC_DF",      "YC_DF(0.05,3)", 0.860708, 0.0001, Check.Num),
            new("YC_TO_CONT", "YC_TO_CONT(0.06,2)", 0, 0, Check.IsNum),
            new("YC_FROM_CONT","YC_FROM_CONT(0.06,2)", 0, 0, Check.IsNum),
            new("YC_FWD_RATE","YC_FWD_RATE(0.05,1,0.05,2)", 0.05, 0.0001, Check.Num),
            new("MORT_PAYMENT", "MORT_PAYMENT(100000,0.06,30,12)", 599.55, 0.1, Check.Num),
            new("MORT_BALANCE", "MORT_BALANCE(100000,0.06,30,0,12)", 100000, 1, Check.Num),
            new("MORT_TOTAL_INTEREST", "MORT_TOTAL_INTEREST(100000,0.06,30,12)", 0, 0, Check.Pos),
            new("MORT_EAR",   "MORT_EAR(0.06,12)", 0.061678, 0.0001, Check.Num),

            // ── Forwards & Futures ──
            new("FWD_PRICE", "FWD_PRICE(100,0.05,1)", 105.1271, 0.001, Check.Num),
            new("FWD_PRICE_YIELD", "FWD_PRICE_YIELD(100,0.05,0.02,1)", 103.0455, 0.001, Check.Num),
            new("FWD_PRICE_INCOME", "FWD_PRICE_INCOME(100,5,0.05,1)", 0, 0, Check.Pos),
            new("FWD_FX", "FWD_FX(1.2,0.04,0.01,0.5)", 1.21813, 0.001, Check.Num),
            new("FWD_COMMODITY", "FWD_COMMODITY(50,0.03,0.02,0.01,1)", 0, 0, Check.Pos),
            new("FWD_VALUE", "FWD_VALUE(110,100,0.05,1)", 9.5123, 0.001, Check.Num),
            new("FWD_VALUE_SHORT", "FWD_VALUE_SHORT(110,100,0.05,1)", -9.5123, 0.001, Check.Num),

            // ── FRA ──
            new("FRA_RATE", "FRA_RATE(0.03,1,0.035,2)", 0.04, 0.0001, Check.Num),
            new("FRA_RATE_SIMPLE", "FRA_RATE_SIMPLE(0.03,1,0.035,2)", 0, 0, Check.Pos),
            new("FRA_VALUE", "FRA_VALUE(1000000,0.05,0.03,1,0.035,2,TRUE)", 0, 0, Check.IsNum),
            new("FRA_SETTLEMENT", "FRA_SETTLEMENT(1000000,0.05,0.06,1,2,TRUE)", 0, 0, Check.IsNum),
            new("FRA_DV01", "FRA_DV01(1000000,0.05,0.03,1,0.035,2,TRUE)", 0, 0, Check.IsNum),

            // ── IRS / Black model (scalar-only ones) ──
            new("IRS_FLOAT_LEG", "IRS_FLOAT_LEG(100,0.026,0.25,0.03)", 0, 0, Check.Pos),
            new("BM_CAPLET", "BM_CAPLET(100,0.04,0.04,1,0.03,0.2,0.5,FALSE)", 0, 0, Check.Pos),

            // ── Short-Rate Models ──
            new("SR_VASICEK_PRICE", "SR_VASICEK_PRICE(0.03,5,0.3,0.05,0.02)", 0, 0, Check.In01),
            new("SR_VASICEK_YIELD", "SR_VASICEK_YIELD(0.03,5,0.3,0.05,0.02)", 0, 0, Check.IsNum),
            new("SR_VASICEK_LRYIELD","SR_VASICEK_LRYIELD(0.3,0.05,0.02)", 0, 0, Check.IsNum),
            new("SR_VASICEK_OPTION", "SR_VASICEK_OPTION(0.03,1,5,0.8,0.3,0.05,0.02,FALSE)", 0, 0, Check.Pos),
            new("SR_CIR_PRICE", "SR_CIR_PRICE(0.03,5,0.3,0.05,0.02)", 0, 0, Check.In01),
            new("SR_CIR_YIELD", "SR_CIR_YIELD(0.03,5,0.3,0.05,0.02)", 0, 0, Check.IsNum),
            new("SR_CIR_LRYIELD","SR_CIR_LRYIELD(0.3,0.05,0.02)", 0, 0, Check.IsNum),

            // ── Credit ──
            new("CR_MERTON_EQUITY", "CR_MERTON_EQUITY(100,80,1,0.05,0.2)", 0, 0, Check.Pos),
            new("CR_MERTON_DEBT",   "CR_MERTON_DEBT(100,80,1,0.05,0.2)", 0, 0, Check.Pos),
            new("CR_DEFAULT_PROB",  "CR_DEFAULT_PROB(100,80,1,0.05,0.2)", 0, 0, Check.In01),
            new("CR_DIST_TO_DEFAULT","CR_DIST_TO_DEFAULT(100,80,1,0.05,0.2)", 0, 0, Check.Pos),
            new("CR_CREDIT_SPREAD", "CR_CREDIT_SPREAD(100,80,1,0.05,0.2)", 0, 0, Check.Pos),
            new("CR_SURVIVAL_PROB", "CR_SURVIVAL_PROB(0.02,5)", 0.904837, 0.0001, Check.Num),
            new("CR_HAZARD_FROM_SPREAD", "CR_HAZARD_FROM_SPREAD(0.012,0.4)", 0.02, 0.0001, Check.Num),
            new("CR_CDS_SPREAD", "CR_CDS_SPREAD(0.02,0.03,5,0.4,4)", 0, 0, Check.Pos),
            new("CR_CDS_MTM", "CR_CDS_MTM(0.005,0.03,0.03,10000000,5,0.4,4)", 0, 0, Check.IsNum),

            // ── Volatility (GARCH, scalar) ──
            new("VOL_GARCH_LONGRUN", "VOL_GARCH_LONGRUN(0.000002,0.05,0.9)", 0.00004, 0.000001, Check.Num),
            new("VOL_GARCH_FORECAST","VOL_GARCH_FORECAST(0.0002,0.000002,0.05,0.9,10)", 0, 0, Check.Pos),

            // ── Fees ──
            new("FEE_MGMT", "FEE_MGMT(100000000,0.02,90,365)", 493150.68, 1, Check.Num),
            new("FEE_PERF", "FEE_PERF(120,100,100,0.2,0)", 4.0, 0.0001, Check.Num),
            new("FEE_EXPENSE_DRAG", "FEE_EXPENSE_DRAG(0.1,0.01,5)", 0, 0, Check.Pos),
            new("FEE_NET_RETURN", "FEE_NET_RETURN(0.1,0.01)", 0.089109, 0.0001, Check.Num),
            new("FEE_CARRIED_INT", "FEE_CARRIED_INT(1280,1000,0.08,1,0.2)", 40, 0.0001, Check.Num),
            new("FEE_TRANSACTION_COST", "FEE_TRANSACTION_COST(1000000,0.001,4)", 1200, 0.0001, Check.Num),

            // ── Attribution (scalar) ──
            new("ATTR_ALLOC", "ATTR_ALLOC(0.5,0.4,0.05,0.04)", 0, 0, Check.IsNum),
            new("ATTR_SELECT", "ATTR_SELECT(0.4,0.1,0.08)", 0.008, 0.0001, Check.Num),
            new("ATTR_INTERACT", "ATTR_INTERACT(0.5,0.4,0.1,0.08)", 0.002, 0.0001, Check.Num),
            new("ATTR_ACTIVE_RETURN", "ATTR_ACTIVE_RETURN(0.07,0.05)", 0.02, 0.0001, Check.Num),

            // ── Equity ──
            new("EQ_MKTCAP", "EQ_MKTCAP(10000,100)", 1000000, 0.001, Check.Num),
            new("EQ_EV", "EQ_EV(1000000,300000,100000)", 1200000, 0.001, Check.Num),
            new("EQ_PE", "EQ_PE(100,5)", 20, 0.0001, Check.Num),
            new("EQ_PB", "EQ_PB(50,20)", 2.5, 0.0001, Check.Num),
            new("EQ_PS", "EQ_PS(2000,500)", 4.0, 0.0001, Check.Num),
            new("EQ_EVTOEBITDA", "EQ_EVTOEBITDA(1000,100)", 10, 0.0001, Check.Num),
            new("EQ_DIV_YIELD", "EQ_DIV_YIELD(2,50)", 0.04, 0.0001, Check.Num),
            new("EQ_UNREAL_PNL", "EQ_UNREAL_PNL(10,100,120)", 200, 0.0001, Check.Num),
            new("EQ_REAL_PNL", "EQ_REAL_PNL(10,100,90)", -100, 0.0001, Check.Num),
            new("EQ_KELLY", "EQ_KELLY(0.1,0.2)", 2.5, 0.0001, Check.Num),
            new("EQ_HALF_KELLY", "EQ_HALF_KELLY(0.1,0.2)", 1.25, 0.0001, Check.Num),
        };
    }
}
