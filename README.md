# Excel Finance Add-In

Excel add-in for financial calculations built with ExcelDNA. Options pricing (Black-Scholes and Greeks), bond math, and risk metrics as native Excel functions.

**Download:** *(GitHub Releases — coming soon)*

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

The function implementations live in [`AleksejCupic.FinancialMath`](https://github.com/aleksejcupic/financial-math) — a separate NuGet library. This add-in is a thin ExcelDNA wrapper that exposes those calculations as Excel UDFs.

```
AleksejCupic.FinancialMath  (NuGet — math logic)
        ↑
ExcelFinanceAddin            (ExcelDNA — Excel wrapper)
        ↓
ExcelFinanceAddin-packed.xll (distributed via GitHub Releases)
```

---

## Install

1. Download `ExcelFinanceAddin-packed.xll` from GitHub Releases
2. Open Excel → File → Options → Add-Ins → Manage: Excel Add-ins → Go
3. Click Browse, select the `.xll` file, click OK

---

## Build from source

Requires .NET 6 SDK and Windows.

```bash
git clone https://github.com/aleksejcupic/excel-finance-addin
cd excel-finance-addin
dotnet build --configuration Release
```

The `.xll` will be in `bin/Release/net6.0-windows/`.

---

## Author

Aleksej Cupic  
[aleksejcupic.com](https://aleksejcupic.com)
