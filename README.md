# Excel Finance Add-In

Excel add-in for financial calculations built with ExcelDNA. Options pricing (Black-Scholes and Greeks), bond math, and risk metrics as native Excel functions.

**Download:** grab `Aleksej.Finance.ExcelAddin.xll` from the [latest release](https://github.com/aleksejcupic/aleksej-finance-excel/releases/latest). A single self-contained 64-bit file. No other downloads needed.

---

## Functions

### Options (Black-Scholes)

| Function | Description |
|---|---|
| `=BS_PUT(S, K, T, r, sigma)` | European put price |
| `=BS_CALL(S, K, T, r, sigma)` | European call price |
| `=BS_DELTA(S, K, T, r, sigma, isPut)` | Delta |
| `=BS_GAMMA(S, K, T, r, sigma)` | Gamma |
| `=BS_VEGA(S, K, T, r, sigma)` | Vega (per 1% vol move) |
| `=BS_THETA(S, K, T, r, sigma, isPut)` | Daily theta |
| `=BS_RHO(S, K, T, r, sigma, isPut)` | Rho (per 1% rate move) |
| `=BS_IV(marketPrice, S, K, T, r, isPut)` | Implied volatility |

### Bonds

| Function | Description |
|---|---|
| `=BOND_PRICE(face, couponRate, ytm, years)` | Bond present value |
| `=BOND_YTM(price, face, couponRate, years)` | Yield to maturity |
| `=BOND_DURATION(face, couponRate, ytm, years)` | Macaulay duration |
| `=BOND_MOD_DURATION(face, couponRate, ytm, years)` | Modified duration |
| `=BOND_CONVEXITY(face, couponRate, ytm, years)` | Convexity |
| `=BOND_DV01(face, couponRate, ytm, years)` | Dollar value of 1bp |

### Risk metrics

| Function | Description |
|---|---|
| `=SHARPE_RATIO(dailyReturns, riskFreeRate)` | Annualised Sharpe ratio |
| `=VAR_HISTORICAL(dailyReturns, confidence)` | Historical VaR |
| `=VAR_CVAR(dailyReturns, confidence)` | Conditional VaR (Expected Shortfall) |
| `=VAR_PARAMETRIC(dailyReturns, confidence)` | Parametric VaR (normal distribution) |
| `=ANN_RETURN(dailyReturns)` | Annualised return |
| `=ANN_VOL(dailyReturns)` | Annualised volatility |
| `=MAX_DRAWDOWN(dailyReturns)` | Maximum drawdown |

---

## Architecture

The function implementations live in [`Aleksej.Finance`](https://www.nuget.org/packages/Aleksej.Finance) ([source](https://github.com/aleksejcupic/aleksej-finance)), a separate NuGet library. This add-in is a thin ExcelDNA wrapper that exposes those calculations as Excel UDFs.

```
Aleksej.Finance                       (NuGet: math logic)
        ↑
Aleksej.Finance.Excel                 (ExcelDNA: Excel wrapper)
        ↓
Aleksej.Finance.ExcelAddin.xll             (single self-contained 64-bit add-in)
```

---

## Install

1. Download `Aleksej.Finance.ExcelAddin.xll` from GitHub Releases
2. Open Excel → File → Options → Add-Ins → Manage: Excel Add-ins → Go
3. Click Browse, select the `.xll` file, click OK

---

## Build from source

Requires the .NET SDK (8.0+) and Windows. The add-in targets **.NET Framework 4.8** (built into Windows), so the build needs no extra runtime; `Microsoft.NETFramework.ReferenceAssemblies` is restored automatically.

```bash
git clone https://github.com/aleksejcupic/aleksej-finance-excel
cd aleksej-finance-excel
dotnet build --configuration Release
```

The single self-contained add-in will be at `bin/Release/net48/Aleksej.Finance.ExcelAddin.xll`: one file, no other downloads, no runtime to install.

---

## Author

Aleksej Cupic  
[aleksejcupic.com](https://aleksejcupic.com)
