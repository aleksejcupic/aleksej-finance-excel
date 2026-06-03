# Generates Aleksej.Finance-Tests.xlsx — a self-checking workbook covering all 152 UDFs.
# Output: Aleksej.Finance-Tests.xlsx in this folder. Open with the add-in loaded, Ctrl+Alt+F9.
# Single source of truth for the test cases (also mirrored by the ExcelDna.Testing suite).
$ErrorActionPreference = 'Stop'

function ColL([int]$n){ $s=''; while($n -gt 0){ $m=($n-1)%26; $s=[char](65+$m)+$s; $n=[int](($n-$m-1)/26) }; return $s }
function Esc($t){ return ($t -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;') }

$cells = @{}
function Add-RowObj([int]$row){ if(-not $cells.ContainsKey($row)){ $cells[$row]=New-Object System.Collections.ArrayList } }
function TCell([int]$row,[int]$col,[string]$txt,[int]$style=0){ Add-RowObj $row; $ref=(ColL $col)+$row; $sa= if($style -ne 0){' s="'+$style+'"'}else{''}; [void]$cells[$row].Add(@{C=$col;X='<c r="'+$ref+'" t="inlineStr"'+$sa+'><is><t xml:space="preserve">'+(Esc $txt)+'</t></is></c>'}) }
function NCell([int]$row,[int]$col,$val,[int]$style=0){ Add-RowObj $row; $ref=(ColL $col)+$row; $sa= if($style -ne 0){' s="'+$style+'"'}else{''}; [void]$cells[$row].Add(@{C=$col;X='<c r="'+$ref+'"'+$sa+'><v>'+$val+'</v></c>'}) }
function FCell([int]$row,[int]$col,[string]$f,[int]$style=0){ Add-RowObj $row; $ref=(ColL $col)+$row; $sa= if($style -ne 0){' s="'+$style+'"'}else{''}; [void]$cells[$row].Add(@{C=$col;X='<c r="'+$ref+'"'+$sa+'><f>'+(Esc $f)+'</f></c>'}) }

# ── Data blocks (right of the table) referenced by range/matrix functions ──
function WriteCol([int]$row,[int]$col,[double[]]$v){ for($i=0;$i -lt $v.Length;$i++){ NCell ($row+$i) $col $v[$i] }; return (ColL $col)+$row+':'+(ColL $col)+($row+$v.Length-1) }
function WriteBlock([int]$row,[int]$col,[double[][]]$m){ for($r=0;$r -lt $m.Length;$r++){ for($c=0;$c -lt $m[$r].Length;$c++){ NCell ($row+$r) ($col+$c) $m[$r][$c] } }; return (ColL $col)+$row+':'+(ColL ($col+$m[0].Length-1))+($row+$m.Length-1) }

TCell 1 8 "DATA BLOCKS (inputs for range/matrix functions)" 1
$RET   = WriteCol 3 8  @(0.012,-0.008,0.015,0.004,-0.011,0.020,-0.006,0.009,-0.003,0.014)  # H3:H12
$BENCH = WriteCol 3 9  @(0.010,-0.005,0.012,0.003,-0.009,0.016,-0.004,0.007,-0.002,0.011)  # I3:I12
$W3    = WriteCol 3 11 @(0.5,0.3,0.2)                 # K3:K5
$MU3   = WriteCol 3 12 @(0.08,0.12,0.18)             # L3:L5
$COV3  = WriteBlock 3 13 @(@(0.04,0.01,0.00),@(0.01,0.06,0.01),@(0.00,0.01,0.09))  # M3:O5
$DD3   = WriteCol 3 17 @(0.10,-0.20,0.05)            # Q3:Q5  (MaxDrawdown = 0.20)
$TWR3  = WriteCol 3 18 @(0.10,0.05,-0.05)            # R3:R5  (TWR = 0.09975)
$AR3   = WriteCol 3 19 @(0.01,0.02,0.03)             # S3:S5  (AnnReturn x10 = 0.20)
$CF2   = WriteCol 3 21 @(-1000.0,1100.0)             # U3:U4
$TIM2  = WriteCol 3 22 @(0.0,1.0)                    # V3:V4  (IRR=0.09531, NPV@5%=51.28)
$MAT4  = WriteCol 3 24 @(0.5,1.0,1.5,2.0)            # X3:X6
$ZERO4 = WriteCol 3 25 @(0.03,0.032,0.034,0.035)    # Y3:Y6
$NAV5  = WriteCol 3 27 @(100.0,110.0,105.0,130.0,120.0) # AA3:AA7 (HWM=130)
$POS3  = WriteCol 3 29 @(10.0,20.0,5.0)              # AC3:AC5
$PRC3  = WriteCol 3 30 @(100.0,50.0,200.0)           # AD3:AD5 (PortValue=3000)
$BPW   = WriteCol 3 32 @(0.5,0.5)                    # AF3:AF4
$BBW   = WriteCol 3 33 @(0.4,0.6)                    # AG3:AG4
$BPR   = WriteCol 3 34 @(0.10,0.04)                  # AH3:AH4
$BBR   = WriteCol 3 35 @(0.08,0.03)                  # AI3:AI4
$BMT   = WriteCol 3 37 @(0.5,1.0,1.5)                # AK3:AK5  black-model payment times
$BMZ   = WriteCol 3 38 @(0.03,0.032,0.034)           # AL3:AL5  zeros
$BMF   = WriteCol 3 39 @(0.035,0.036,0.037)          # AM3:AM5  forwards
$BMA   = WriteCol 3 40 @(0.5,0.5,0.5)                # AN3:AN5  accruals
$PVCF  = WriteCol 3 42 @(2.0,2.0,2.0)                # AP3:AP5
$PVT   = WriteCol 3 43 @(0.25,0.5,0.75)              # AQ3:AQ5
$MDCF  = WriteCol 3 45 @(100.0)                      # AS3
$MDD   = WriteCol 3 46 @(182.5)                      # AT3

# ── Title / dashboard / headers ──
TCell 1 1 "Aleksej.Finance Excel Add-In - Full Test Suite (152 UDFs)" 1
TCell 2 1 "Load the add-in, then press Ctrl+Alt+F9. Dashboard below shows PASS/FAIL. (Async + array rows may need a 2nd recalc.)"
TCell 3 1 "RESULT:" 1
FCell 3 2 'COUNTIF(F7:F400,"PASS")&" / "&(COUNTIF(F7:F400,"PASS")+COUNTIF(F7:F400,"FAIL"))&" PASS"' 1
$HDR=6
TCell $HDR 1 "Category" 1; TCell $HDR 2 "Function" 1; TCell $HDR 3 "Result" 1; TCell $HDR 4 "Expected" 1; TCell $HDR 5 "Tol" 1; TCell $HDR 6 "Status" 1

# ── Case engine ──
$script:row = 7
function Status([int]$r,[string]$check){
  switch($check){
    'num'   { return "IF(ISNUMBER(C$r),IF(ABS(C$r-D$r)<=E$r,`"PASS`",`"FAIL`"),`"FAIL`")" }
    'pos'   { return "IF(AND(ISNUMBER(C$r),C$r>0),`"PASS`",`"FAIL`")" }
    'neg'   { return "IF(AND(ISNUMBER(C$r),C$r<0),`"PASS`",`"FAIL`")" }
    'in01'  { return "IF(AND(ISNUMBER(C$r),C$r>=0,C$r<=1),`"PASS`",`"FAIL`")" }
    'isnum' { return "IF(ISNUMBER(C$r),`"PASS`",`"FAIL`")" }
    'text'  { return "IF(EXACT(C$r,D$r),`"PASS`",`"FAIL`")" }
  }
}
# Case: cat, name, formula(without =), expected(number or text or note), tol, check
function Case([string]$cat,[string]$name,[string]$formula,$expected,$tol,[string]$check){
  $r=$script:row
  TCell $r 1 $cat; TCell $r 2 $name; FCell $r 3 $formula
  if($check -eq 'num'){ NCell $r 4 $expected; NCell $r 5 $tol }
  elseif($check -eq 'text'){ TCell $r 4 $expected }
  else { TCell $r 4 $expected }   # expected here is a human note
  FCell $r 6 (Status $r $check)
  $script:row++
}
function Section([string]$title){ $r=$script:row; TCell $r 1 $title 1; $script:row++ }

# ===================== OPTIONS — Black-Scholes =====================
Section "OPTIONS - Black-Scholes"
Case "Options" "BS_CALL"  "BS_CALL(100,100,1,0.05,0.2)"  10.4506 0.001 num
Case "Options" "BS_PUT"   "BS_PUT(100,100,1,0.05,0.2)"   5.5735  0.001 num
Case "Options" "BS_DELTA" "BS_DELTA(100,100,1,0.05,0.2,FALSE)" 0.6368 0.001 num
Case "Options" "BS_GAMMA" "BS_GAMMA(100,100,1,0.05,0.2)" 0.01876 0.001 num
Case "Options" "BS_VEGA"  "BS_VEGA(100,100,1,0.05,0.2)"  0.3752  0.001 num
Case "Options" "BS_THETA" "BS_THETA(100,100,1,0.05,0.2,FALSE)" "negative" 0 neg
Case "Options" "BS_RHO"   "BS_RHO(100,100,1,0.05,0.2,FALSE)"   "positive" 0 pos
Case "Options" "BS_IV"    "BS_IV(10.4506,100,100,1,0.05,FALSE)" 0.20 0.001 num
Case "Options" "BS_VANNA" "BS_VANNA(100,100,1,0.05,0.2)" "number" 0 isnum
Case "Options" "BS_CHARM" "BS_CHARM(100,100,1,0.05,0.2)" "number" 0 isnum
Case "Options" "BS_VOLGA" "BS_VOLGA(100,100,1,0.05,0.2)" "number" 0 isnum
Case "Options" "BS_SPEED" "BS_SPEED(100,100,1,0.05,0.2)" "number" 0 isnum
Case "Options" "BS_ZOMMA" "BS_ZOMMA(100,100,1,0.05,0.2)" "number" 0 isnum

# ===================== OPTIONS — Binomial =====================
Section "OPTIONS - Binomial Tree"
Case "Options" "BT_PRICE" "BT_PRICE(100,100,1,0.05,0.2,300,FALSE,FALSE)" 10.4506 0.1 num
Case "Options" "BT_DELTA" "BT_DELTA(100,100,1,0.05,0.2,300,FALSE,FALSE)" "in[0,1]" 0 in01
Case "Options" "BT_GAMMA" "BT_GAMMA(100,100,1,0.05,0.2,300,FALSE)" "positive" 0 pos

# ===================== OPTIONS — Exotic =====================
Section "OPTIONS - Exotic"
Case "Options" "EX_BINARY_CASH"  "EX_BINARY_CASH(100,100,1,0.05,0.2,10,FALSE)" "in[0,10]" 0 pos
Case "Options" "EX_BINARY_ASSET" "EX_BINARY_ASSET(100,100,1,0.05,0.2,FALSE)" "positive" 0 pos
Case "Options" "EX_BARRIER_CALL" "EX_BARRIER_CALL(100,100,90,1,0.05,0.2,FALSE,FALSE)" "positive" 0 pos
Case "Options" "EX_BARRIER_PUT"  "EX_BARRIER_PUT(100,100,90,1,0.05,0.2,FALSE,FALSE)" "positive" 0 pos
Case "Options" "EX_ASIAN_GEO"    "EX_ASIAN_GEO(100,100,1,0.05,0.2,FALSE)" "positive" 0 pos
Case "Options" "EX_ASIAN_ARITH"  "EX_ASIAN_ARITH(100,100,1,0.05,0.2,12,20000,FALSE,42)" "positive (stochastic)" 0 pos
Case "Options" "EX_LOOKBACK_CALL" "EX_LOOKBACK_CALL(100,100,1,0.05,0.2)" "positive" 0 pos
Case "Options" "EX_LOOKBACK_PUT"  "EX_LOOKBACK_PUT(100,100,1,0.05,0.2)" "positive" 0 pos

# ===================== OPTIONS — Garman-Kohlhagen =====================
Section "OPTIONS - Garman-Kohlhagen (FX)"
Case "Options" "GK_CALL" "GK_CALL(1.2,1.25,1,0.05,0.03,0.15)" "positive" 0 pos
Case "Options" "GK_PUT"  "GK_PUT(1.2,1.25,1,0.05,0.03,0.15)" "positive" 0 pos
Case "Options" "GK_DELTA" "GK_DELTA(1.2,1.25,1,0.05,0.03,0.15,FALSE)" "in[0,1]" 0 in01
Case "Options" "GK_GAMMA" "GK_GAMMA(1.2,1.25,1,0.05,0.03,0.15)" "positive" 0 pos
Case "Options" "GK_VEGA"  "GK_VEGA(1.2,1.25,1,0.05,0.03,0.15)" "positive" 0 pos
Case "Options" "GK_THETA" "GK_THETA(1.2,1.25,1,0.05,0.03,0.15,FALSE)" "number" 0 isnum
Case "Options" "GK_RHO"   "GK_RHO(1.2,1.25,1,0.05,0.03,0.15,FALSE)" "number" 0 isnum
Case "Options" "GK_RHO_FOREIGN" "GK_RHO_FOREIGN(1.2,1.25,1,0.05,0.03,0.15,FALSE)" "number" 0 isnum
Case "Options" "GK_IV" "GK_IV(GK_CALL(1.2,1.25,1,0.05,0.03,0.15),1.2,1.25,1,0.05,0.03,FALSE)" 0.15 0.001 num

# ===================== OPTIONS — Options on Futures =====================
Section "OPTIONS - Options on Futures"
Case "Options" "OF_CALL" "OF_CALL(100,100,1,0.05,0.2)" "positive" 0 pos
Case "Options" "OF_PUT"  "OF_PUT(100,100,1,0.05,0.2)" "positive" 0 pos
Case "Options" "OF_CALL_FROM_PUT" "OF_CALL_FROM_PUT(OF_PUT(100,100,1,0.05,0.2),100,100,1,0.05)" "positive" 0 pos
Case "Options" "OF_DELTA" "OF_DELTA(100,100,1,0.05,0.2,FALSE)" "in[0,1]" 0 in01
Case "Options" "OF_GAMMA" "OF_GAMMA(100,100,1,0.05,0.2)" "positive" 0 pos
Case "Options" "OF_VEGA"  "OF_VEGA(100,100,1,0.05,0.2)" "positive" 0 pos
Case "Options" "OF_IV"    "OF_IV(OF_CALL(100,100,1,0.05,0.2),100,100,1,0.05,FALSE)" 0.20 0.001 num

# ===================== BONDS =====================
Section "BONDS"
Case "Bonds" "BOND_PRICE" "BOND_PRICE(1000,0.05,0.05,10)" 1000 0.001 num
Case "Bonds" "BOND_YTM"   "BOND_YTM(1000,1000,0.05,10)" 0.05 0.0001 num
Case "Bonds" "BOND_DURATION" "BOND_DURATION(1000,0.05,0.05,10,2)" "positive" 0 pos
Case "Bonds" "BOND_MOD_DURATION" "BOND_MOD_DURATION(1000,0.05,0.05,10,2)" "positive" 0 pos
Case "Bonds" "BOND_CONVEXITY" "BOND_CONVEXITY(1000,0.05,0.05,10,2)" "positive" 0 pos
Case "Bonds" "BOND_DV01" "BOND_DV01(1000,0.05,0.05,10,2)" "positive" 0 pos
Case "Bonds" "BOND_PRICE_CHANGE" "BOND_PRICE_CHANGE(1000,0.05,0.05,10,0.001,2)" "negative" 0 neg
Case "Bonds" "YC_DF" "YC_DF(0.05,3)" 0.860708 0.0001 num
Case "Bonds" "YC_TO_CONT" "YC_TO_CONT(0.06,2)" "number" 0 isnum
Case "Bonds" "YC_FROM_CONT" "YC_FROM_CONT(0.06,2)" "number" 0 isnum
Case "Bonds" "YC_FWD_RATE" "YC_FWD_RATE(0.05,1,0.05,2)" 0.05 0.0001 num
Case "Bonds" "YC_INTERPOLATE" "YC_INTERPOLATE($MAT4,$ZERO4,1.0)" 0.032 0.0001 num
Case "Bonds" "YC_PAR_YIELD" "YC_PAR_YIELD($MAT4,$ZERO4,2,2)" "positive" 0 pos
Case "Bonds" "MORT_PAYMENT" "MORT_PAYMENT(100000,0.06,30,12)" 599.55 0.1 num
Case "Bonds" "MORT_BALANCE" "MORT_BALANCE(100000,0.06,30,0,12)" 100000 1 num
Case "Bonds" "MORT_TOTAL_INTEREST" "MORT_TOTAL_INTEREST(100000,0.06,30,12)" "positive" 0 pos
Case "Bonds" "MORT_EAR" "MORT_EAR(0.06,12)" 0.061678 0.0001 num

# ===================== DERIVATIVES — Forwards/Futures =====================
Section "DERIVATIVES - Forwards & Futures"
Case "Derivatives" "FWD_PRICE" "FWD_PRICE(100,0.05,1)" 105.1271 0.001 num
Case "Derivatives" "FWD_PRICE_YIELD" "FWD_PRICE_YIELD(100,0.05,0.02,1)" 103.0455 0.001 num
Case "Derivatives" "FWD_PRICE_INCOME" "FWD_PRICE_INCOME(100,5,0.05,1)" "positive" 0 pos
Case "Derivatives" "FWD_FX" "FWD_FX(1.2,0.04,0.01,0.5)" 1.21813 0.001 num
Case "Derivatives" "FWD_COMMODITY" "FWD_COMMODITY(50,0.03,0.02,0.01,1)" "positive" 0 pos
Case "Derivatives" "FWD_VALUE" "FWD_VALUE(110,100,0.05,1)" 9.5123 0.001 num
Case "Derivatives" "FWD_VALUE_SHORT" "FWD_VALUE_SHORT(110,100,0.05,1)" -9.5123 0.001 num
Case "Derivatives" "FWD_PV_INCOME" "FWD_PV_INCOME($PVCF,$PVT,0.05)" "positive" 0 pos

# ===================== DERIVATIVES — FRA =====================
Section "DERIVATIVES - FRA"
Case "Derivatives" "FRA_RATE" "FRA_RATE(0.03,1,0.035,2)" 0.04 0.0001 num
Case "Derivatives" "FRA_RATE_SIMPLE" "FRA_RATE_SIMPLE(0.03,1,0.035,2)" "positive" 0 pos
Case "Derivatives" "FRA_VALUE" "FRA_VALUE(1000000,0.05,0.03,1,0.035,2,TRUE)" "number" 0 isnum
Case "Derivatives" "FRA_SETTLEMENT" "FRA_SETTLEMENT(1000000,0.05,0.06,1,2,TRUE)" "number" 0 isnum
Case "Derivatives" "FRA_DV01" "FRA_DV01(1000000,0.05,0.03,1,0.035,2,TRUE)" "number" 0 isnum

# ===================== DERIVATIVES — IRS =====================
Section "DERIVATIVES - Interest Rate Swap"
Case "Derivatives" "IRS_VALUE" "IRS_VALUE(100,0.05,$MAT4,$ZERO4,0.026,0.25,0.029,TRUE)" "number" 0 isnum
Case "Derivatives" "IRS_PAR_RATE" "IRS_PAR_RATE($MAT4,$ZERO4,2)" "positive" 0 pos
Case "Derivatives" "IRS_FIXED_LEG" "IRS_FIXED_LEG(100,0.05,$MAT4,$ZERO4,2,FALSE)" "positive" 0 pos
Case "Derivatives" "IRS_FLOAT_LEG" "IRS_FLOAT_LEG(100,0.026,0.25,0.03)" "positive" 0 pos
Case "Derivatives" "IRS_DV01" "IRS_DV01(100,0.05,$MAT4,$ZERO4,0.026,0.25,0.029,TRUE)" "number" 0 isnum

# ===================== DERIVATIVES — Black Model =====================
Section "DERIVATIVES - Black Model (Caps/Floors/Swaptions)"
Case "Derivatives" "BM_CAPLET" "BM_CAPLET(100,0.04,0.04,1,0.03,0.2,0.5,FALSE)" "positive" 0 pos
Case "Derivatives" "BM_CAP" "BM_CAP(100,0.04,0.2,$BMT,$BMZ,$BMF,$BMA)" "positive" 0 pos
Case "Derivatives" "BM_FLOOR" "BM_FLOOR(100,0.04,0.2,$BMT,$BMZ,$BMF,$BMA)" "positive" 0 pos
Case "Derivatives" "BM_FWD_SWAP_RATE" "BM_FWD_SWAP_RATE($BMT,$BMZ,$BMA)" "positive" 0 pos
Case "Derivatives" "BM_SWAPTION" "BM_SWAPTION(100,0.035,1,0.1,$BMT,$BMZ,$BMA,TRUE)" "positive" 0 pos

# ===================== DERIVATIVES — Short Rate =====================
Section "DERIVATIVES - Short-Rate Models"
Case "Derivatives" "SR_VASICEK_PRICE" "SR_VASICEK_PRICE(0.03,5,0.3,0.05,0.02)" "in(0,1]" 0 in01
Case "Derivatives" "SR_VASICEK_YIELD" "SR_VASICEK_YIELD(0.03,5,0.3,0.05,0.02)" "number" 0 isnum
Case "Derivatives" "SR_VASICEK_LRYIELD" "SR_VASICEK_LRYIELD(0.3,0.05,0.02)" "number" 0 isnum
Case "Derivatives" "SR_VASICEK_OPTION" "SR_VASICEK_OPTION(0.03,1,5,0.8,0.3,0.05,0.02,FALSE)" "positive" 0 pos
Case "Derivatives" "SR_CIR_PRICE" "SR_CIR_PRICE(0.03,5,0.3,0.05,0.02)" "in(0,1]" 0 in01
Case "Derivatives" "SR_CIR_YIELD" "SR_CIR_YIELD(0.03,5,0.3,0.05,0.02)" "number" 0 isnum
Case "Derivatives" "SR_CIR_LRYIELD" "SR_CIR_LRYIELD(0.3,0.05,0.02)" "number" 0 isnum

# ===================== CREDIT =====================
Section "CREDIT"
Case "Credit" "CR_MERTON_EQUITY" "CR_MERTON_EQUITY(100,80,1,0.05,0.2)" "positive" 0 pos
Case "Credit" "CR_MERTON_DEBT" "CR_MERTON_DEBT(100,80,1,0.05,0.2)" "positive" 0 pos
Case "Credit" "CR_DEFAULT_PROB" "CR_DEFAULT_PROB(100,80,1,0.05,0.2)" "in[0,1]" 0 in01
Case "Credit" "CR_DIST_TO_DEFAULT" "CR_DIST_TO_DEFAULT(100,80,1,0.05,0.2)" "positive" 0 pos
Case "Credit" "CR_CREDIT_SPREAD" "CR_CREDIT_SPREAD(100,80,1,0.05,0.2)" "positive" 0 pos
Case "Credit" "CR_SURVIVAL_PROB" "CR_SURVIVAL_PROB(0.02,5)" 0.904837 0.0001 num
Case "Credit" "CR_HAZARD_FROM_SPREAD" "CR_HAZARD_FROM_SPREAD(0.012,0.4)" 0.02 0.0001 num
Case "Credit" "CR_CDS_SPREAD" "CR_CDS_SPREAD(0.02,0.03,5,0.4,4)" "positive" 0 pos
Case "Credit" "CR_CDS_MTM" "CR_CDS_MTM(0.005,0.03,0.03,10000000,5,0.4,4)" "number" 0 isnum

# ===================== PORTFOLIO (Markowitz) =====================
Section "PORTFOLIO - Markowitz"
Case "Portfolio" "PORT_RETURN" "PORT_RETURN($W3,$MU3)" 0.112 0.0001 num
Case "Portfolio" "PORT_VOL" "PORT_VOL($W3,$COV3)" "positive" 0 pos
Case "Portfolio" "PORT_SHARPE" "PORT_SHARPE($W3,$MU3,$COV3,0.02)" "number" 0 isnum
# array-returning funcs handled in the ARRAYS section below

# ===================== RISK =====================
Section "RISK"
Case "Risk" "SHARPE_RATIO" "SHARPE_RATIO($RET,0.02)" "number" 0 isnum
Case "Risk" "VAR_HISTORICAL" "VAR_HISTORICAL($RET,0.95)" "number" 0 isnum
Case "Risk" "VAR_CVAR" "VAR_CVAR($RET,0.95)" "number" 0 isnum
Case "Risk" "VAR_PARAMETRIC" "VAR_PARAMETRIC($RET,0.95)" "number" 0 isnum
Case "Risk" "ANN_RETURN" "ANN_RETURN($AR3,10)" 0.20 0.0001 num
Case "Risk" "ANN_VOL" "ANN_VOL($RET,252)" "positive" 0 pos
Case "Risk" "MAX_DRAWDOWN" "MAX_DRAWDOWN($DD3)" 0.20 0.0001 num
Case "Risk" "RISK_SORTINO" "RISK_SORTINO($RET,0.02,252)" "number" 0 isnum
Case "Risk" "RISK_CALMAR" "RISK_CALMAR($RET,252)" "number" 0 isnum
Case "Risk" "RISK_BETA" "RISK_BETA($RET,$RET)" 1.0 0.0001 num
Case "Risk" "RISK_ALPHA" "RISK_ALPHA($RET,$BENCH,0.02,252)" "number" 0 isnum
Case "Risk" "RISK_TREYNOR" "RISK_TREYNOR($RET,$BENCH,0.02,252)" "number" 0 isnum
Case "Risk" "RISK_TE" "RISK_TE($RET,$BENCH,252)" "positive" 0 pos
Case "Risk" "RISK_IR" "RISK_IR($RET,$BENCH,252)" "number" 0 isnum
Case "Risk" "VOL_EWMA_LATEST" "VOL_EWMA_LATEST($RET,0.94)" "positive" 0 pos
Case "Risk" "VOL_GARCH_LONGRUN" "VOL_GARCH_LONGRUN(0.000002,0.05,0.9)" 0.00004 0.000001 num
Case "Risk" "VOL_GARCH_FORECAST" "VOL_GARCH_FORECAST(0.0002,0.000002,0.05,0.9,10)" "positive" 0 pos

# ===================== FEES =====================
Section "FEES"
Case "Fees" "FEE_MGMT" "FEE_MGMT(100000000,0.02,90,365)" 493150.68 1 num
Case "Fees" "FEE_PERF" "FEE_PERF(120,100,100,0.2,0)" 4.0 0.0001 num
Case "Fees" "FEE_HWM" "FEE_HWM($NAV5)" 130 0.0001 num
Case "Fees" "FEE_EXPENSE_DRAG" "FEE_EXPENSE_DRAG(0.1,0.01,5)" "positive" 0 pos
Case "Fees" "FEE_NET_RETURN" "FEE_NET_RETURN(0.1,0.01)" 0.089109 0.0001 num
Case "Fees" "FEE_CARRIED_INT" "FEE_CARRIED_INT(1280,1000,0.08,1,0.2)" 40 0.0001 num
Case "Fees" "FEE_TRANSACTION_COST" "FEE_TRANSACTION_COST(1000000,0.001,4)" 1200 0.0001 num

# ===================== ATTRIBUTION =====================
Section "ATTRIBUTION"
Case "Attribution" "ATTR_TWR" "ATTR_TWR($TWR3)" 0.09725 0.0001 num
Case "Attribution" "ATTR_MDIETZ" "ATTR_MDIETZ(1000,1200,$MDCF,$MDD,365)" 0.095238 0.001 num
Case "Attribution" "ATTR_IRR" "ATTR_IRR($CF2,$TIM2,0.1)" 0.09531 0.0001 num
Case "Attribution" "ATTR_NPV" "ATTR_NPV($CF2,$TIM2,0.05)" 46.35 0.5 num
Case "Attribution" "ATTR_ALLOC" "ATTR_ALLOC(0.5,0.4,0.05,0.04)" "number" 0 isnum
Case "Attribution" "ATTR_SELECT" "ATTR_SELECT(0.4,0.1,0.08)" 0.008 0.0001 num
Case "Attribution" "ATTR_INTERACT" "ATTR_INTERACT(0.5,0.4,0.1,0.08)" 0.002 0.0001 num
Case "Attribution" "ATTR_BHB_ALLOC" "ATTR_BHB_ALLOC($BPW,$BBW,$BPR,$BBR)" "number" 0 isnum
Case "Attribution" "ATTR_BHB_SELECT" "ATTR_BHB_SELECT($BPW,$BBW,$BPR,$BBR)" "number" 0 isnum
Case "Attribution" "ATTR_BHB_INTERACT" "ATTR_BHB_INTERACT($BPW,$BBW,$BPR,$BBR)" "number" 0 isnum
Case "Attribution" "ATTR_ACTIVE_RETURN" "ATTR_ACTIVE_RETURN(0.07,0.05)" 0.02 0.0001 num

# ===================== EQUITY =====================
Section "EQUITY"
Case "Equity" "EQ_MKTCAP" "EQ_MKTCAP(10000,100)" 1000000 0.001 num
Case "Equity" "EQ_EV" "EQ_EV(1000000,300000,100000)" 1200000 0.001 num
Case "Equity" "EQ_PORT_VALUE" "EQ_PORT_VALUE($POS3,$PRC3)" 3000 0.001 num
Case "Equity" "EQ_PE" "EQ_PE(100,5)" 20 0.0001 num
Case "Equity" "EQ_PB" "EQ_PB(50,20)" 2.5 0.0001 num
Case "Equity" "EQ_PS" "EQ_PS(2000,500)" 4.0 0.0001 num
Case "Equity" "EQ_EVTOEBITDA" "EQ_EVTOEBITDA(1000,100)" 10 0.0001 num
Case "Equity" "EQ_DIV_YIELD" "EQ_DIV_YIELD(2,50)" 0.04 0.0001 num
Case "Equity" "EQ_UNREAL_PNL" "EQ_UNREAL_PNL(10,100,120)" 200 0.0001 num
Case "Equity" "EQ_REAL_PNL" "EQ_REAL_PNL(10,100,90)" -100 0.0001 num
Case "Equity" "EQ_KELLY" "EQ_KELLY(0.1,0.2)" 2.5 0.0001 num
Case "Equity" "EQ_HALF_KELLY" "EQ_HALF_KELLY(0.1,0.2)" 1.25 0.0001 num

# ===================== ASYNC (Monte Carlo) =====================
Section "ASYNC - Monte Carlo (shows #N/A until calc settles - recalc again)"
Case "Options" "MC_EUROPEAN" "MC_EUROPEAN(100,100,1,0.05,0.2,50000,50,FALSE,42)" 10.45 0.5 num
Case "Options" "MC_AMERICAN" "MC_AMERICAN(100,100,1,0.05,0.2,20000,50,TRUE,42)" "positive (>= euro put)" 0 pos

# ===================== ARRAYS (weight vectors; SUM must equal 1, no spilling needed) =====================
Section "ARRAYS - returned weight vector; SUM(result) must equal 1"
Case "Portfolio" "PORT_MIN_VAR"      "SUM(PORT_MIN_VAR($COV3))" 1 0.0001 num
Case "Portfolio" "PORT_RISK_PARITY"  "SUM(PORT_RISK_PARITY($COV3))" 1 0.0001 num
Case "Portfolio" "PORT_MAX_SHARPE"   "SUM(PORT_MAX_SHARPE($MU3,$COV3,0.02))" 1 0.0001 num
Case "Portfolio" "PORT_RISK_CONTRIB" "SUM(PORT_RISK_CONTRIB($W3,$COV3))" 1 0.0001 num

# ===================== ERROR HANDLING & NORMALIZATION =====================
Section "ERROR HANDLING & NORMALIZATION"
Case "Errors" "BS_CALL sigma=0" "BS_CALL(100,100,1,0.05,0)" "sigma (volatility) must be greater than 0." 0 text
Case "Errors" "BS_CALL S=-5" "BS_CALL(-5,100,1,0.05,0.2)" "S must be greater than 0." 0 text
Case "Errors" "BS_CALL sigma=20% text" "BS_CALL(100,100,1,0.05,`"20%`")" 10.4506 0.001 num

# ── Build worksheet XML ──
$sb = New-Object System.Text.StringBuilder
foreach($r in ($cells.Keys | Sort-Object {[int]$_})){
  [void]$sb.Append('<row r="'+$r+'">')
  foreach($cell in ($cells[$r] | Sort-Object {$_.C})){ [void]$sb.Append($cell.X) }
  [void]$sb.Append('</row>')
}
$cf = '<conditionalFormatting sqref="F7:F400"><cfRule type="cellIs" dxfId="0" priority="1" operator="equal"><formula>"PASS"</formula></cfRule><cfRule type="cellIs" dxfId="1" priority="2" operator="equal"><formula>"FAIL"</formula></cfRule></conditionalFormatting>'
$sheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols><col min="1" max="1" width="14" customWidth="1"/><col min="2" max="2" width="22" customWidth="1"/><col min="3" max="3" width="16" customWidth="1"/><col min="4" max="4" width="34" customWidth="1"/><col min="5" max="5" width="9" customWidth="1"/><col min="6" max="6" width="9" customWidth="1"/></cols><sheetData>'+$sb.ToString()+'</sheetData>'+$cf+'</worksheet>'

$parts = [ordered]@{
 '[Content_Types].xml'='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'
 '_rels/.rels'='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
 'xl/workbook.xml'='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Tests" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="0" fullCalcOnLoad="1"/></workbook>'
 'xl/_rels/workbook.xml.rels'='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
 'xl/styles.xml'='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs><dxfs count="2"><dxf><fill><patternFill><bgColor rgb="FFC6EFCE"/></patternFill></fill></dxf><dxf><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs></styleSheet>'
 'xl/worksheets/sheet1.xml'=$sheetXml
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$enc = New-Object System.Text.UTF8Encoding($false)
function Write-Xlsx($outPath){
  $fs = [System.IO.File]::Open($outPath,[System.IO.FileMode]::Create)
  $zip = New-Object System.IO.Compression.ZipArchive($fs,[System.IO.Compression.ZipArchiveMode]::Create)
  foreach($k in $parts.Keys){ $e=$zip.CreateEntry($k,[System.IO.Compression.CompressionLevel]::Optimal); $s=$e.Open(); $b=$enc.GetBytes($parts[$k]); $s.Write($b,0,$b.Length); $s.Close() }
  $zip.Dispose(); $fs.Close()
}
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$local = Join-Path $here 'Aleksej.Finance-Tests.xlsx'
Write-Xlsx $local
Write-Output ("Cases written. Last data row: " + ($script:row))
Write-Output ("Created: " + $local)