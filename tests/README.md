# Tests

Three layers, because only some of this can run without Excel:

| # | What | Covers | Runs |
|---|------|--------|------|
| 1 | **Unit tests** (`Aleksej.Finance.Excel.UnitTests`) | the add-in's own logic — `In.*` validators, `%`-normalization, `RangeHelper` | ✅ headless, **in CI**, with coverage |
| 2 | **Excel integration** (`Aleksej.Finance.Excel.Tests`) | every UDF computed through the real loaded plugin | ⚠️ **local only** — drives real Excel |
| 3 | **Self-checking workbook** (`workbook/`) | eyeball PASS/FAIL for all 152 UDFs | 🖐️ manual |

The financial **math** itself is already tested (xUnit + codecov) in the `Aleksej.Finance`
library repo, so these focus on the Excel layer on top of it.

## 1. Unit tests — automated, CI, coverage

Pure C# tests of the validation/normalization layer. No Excel.

```powershell
dotnet test tests/Aleksej.Finance.Excel.UnitTests -c Release /p:Platform=x64 --collect:"XPlat Code Coverage"
```
Runs on every push via `.github/workflows/ci.yml` and uploads coverage to Codecov
(set a `CODECOV_TOKEN` repo secret).

## 2. Excel integration suite — local only

Uses ExcelDNA's official `ExcelDna.Testing` harness: launches real Excel, loads the add-in,
evaluates each UDF as a formula, asserts the result (scalar / range / array / async / errors).

```powershell
dotnet test tests/Aleksej.Finance.Excel.Tests -c Debug /p:Platform=x64
```
or run from Visual Studio Test Explorer (Excel opens, runs, closes).

**Why not in CI?** It needs 64-bit desktop Excel, which GitHub-hosted runners don't have. CI
only *compiles* it (to catch breakage); running it requires a machine with Office (yours, or a
self-hosted Windows runner).

## 3. Self-checking workbook — manual

`Aleksej.Finance-Tests.xlsx` lives in this folder (`workbook/`) ready to open. Load the add-in
in Excel, open the workbook, press **Ctrl+Alt+F9**, and read the dashboard (`152 / 152 PASS`).
Async (Monte Carlo) rows may need a second recalc.

To regenerate it after changing functions, re-run the generator:

```powershell
pwsh tests/workbook/Build-Tests.ps1
```

---

**Versions:** the integration suite is pinned to `ExcelDna.Testing 1.9.0` + `xunit 2.4.1` to
match the add-in's embedded `ExcelDna.Integration 1.9.0`. Expected values are reused from the
math library's xUnit suite; where no single golden value exists, a structural check
(positive / in [0,1] / is-a-number) is used. Error-message tests assume the default error mode
(descriptive text, not `#NUM!`).
