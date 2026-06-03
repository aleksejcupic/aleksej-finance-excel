using System;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Aleksej.Finance.Excel.Settings;

namespace Aleksej.Finance.Excel.Ribbon;

/// <summary>
/// Custom "Finance Tools" ribbon tab for the Excel Finance Add-In.
/// Provides settings for default parameters, function category toggles, and help links.
/// Settings are persisted across sessions via UserSettings.
/// </summary>
[ComVisible(true)]
public class FinanceRibbon : ExcelRibbon
{
    private IRibbonUI? _ribbon;
    private UserSettings _cfg = UserSettings.Load();

    public override string GetCustomUI(string ribbonId) => RibbonXml;

    // ── Ribbon lifecycle ──────────────────────────────────────────────────────

    public void OnRibbonLoad(IRibbonUI ribbon)
    {
        _ribbon = ribbon;
        _cfg    = UserSettings.Load();
    }

    // ── Default parameter controls ────────────────────────────────────────────

    public string GetRiskFreeRate(IRibbonControl _)   => _cfg.DefaultRiskFreeRate.ToString("G");
    public string GetLambda(IRibbonControl _)         => _cfg.DefaultLambda.ToString("G");

    public void OnRiskFreeRateChange(IRibbonControl _, string text)
    {
        if (double.TryParse(text, out double v)) { _cfg.DefaultRiskFreeRate = v; _cfg.Save(); UserSettings.Invalidate(); }
    }

    public void OnLambdaChange(IRibbonControl _, string text)
    {
        if (double.TryParse(text, out double v)) { _cfg.DefaultLambda = v; _cfg.Save(); UserSettings.Invalidate(); }
    }

    public int GetTradingDaysIndex(IRibbonControl _) => _cfg.DefaultTradingDays switch { 260 => 1, 365 => 2, _ => 0 };
    public int GetConfidenceIndex(IRibbonControl _)  => _cfg.DefaultConfidence switch { 0.99 => 1, 0.999 => 2, _ => 0 };
    public int GetFrequencyIndex(IRibbonControl _)   => _cfg.DefaultFrequency switch { 1 => 0, 4 => 2, 12 => 3, _ => 1 };

    public void OnTradingDaysChange(IRibbonControl _, string selectedId, int selectedIndex)
    {
        _cfg.DefaultTradingDays = selectedIndex switch { 1 => 260, 2 => 365, _ => 252 };
        _cfg.Save();
        UserSettings.Invalidate();
    }

    public void OnConfidenceChange(IRibbonControl _, string selectedId, int selectedIndex)
    {
        _cfg.DefaultConfidence = selectedIndex switch { 1 => 0.99, 2 => 0.999, _ => 0.95 };
        _cfg.Save();
        UserSettings.Invalidate();
    }

    public void OnFrequencyChange(IRibbonControl _, string selectedId, int selectedIndex)
    {
        _cfg.DefaultFrequency = selectedIndex switch { 0 => 1, 2 => 4, 3 => 12, _ => 2 };
        _cfg.Save();
        UserSettings.Invalidate();
    }

    // ── Category toggle checkboxes ────────────────────────────────────────────

    public bool GetEnableOptions(IRibbonControl _)          => _cfg.EnableOptions;
    public bool GetEnableBonds(IRibbonControl _)            => _cfg.EnableBonds;
    public bool GetEnableDerivatives(IRibbonControl _)      => _cfg.EnableDerivatives;
    public bool GetEnableCredit(IRibbonControl _)           => _cfg.EnableCredit;
    public bool GetEnablePortfolioRisk(IRibbonControl _)    => _cfg.EnablePortfolioRisk;
    public bool GetEnableFeesAttribution(IRibbonControl _)  => _cfg.EnableFeesAttribution;
    public bool GetEnableLiveData(IRibbonControl _)         => _cfg.EnableLiveData;

    public void OnEnableOptions(IRibbonControl _, bool pressed)         { _cfg.EnableOptions          = pressed; _cfg.Save(); UserSettings.Invalidate(); }
    public void OnEnableBonds(IRibbonControl _, bool pressed)           { _cfg.EnableBonds            = pressed; _cfg.Save(); UserSettings.Invalidate(); }
    public void OnEnableDerivatives(IRibbonControl _, bool pressed)     { _cfg.EnableDerivatives      = pressed; _cfg.Save(); UserSettings.Invalidate(); }
    public void OnEnableCredit(IRibbonControl _, bool pressed)          { _cfg.EnableCredit           = pressed; _cfg.Save(); UserSettings.Invalidate(); }
    public void OnEnablePortfolioRisk(IRibbonControl _, bool pressed)   { _cfg.EnablePortfolioRisk    = pressed; _cfg.Save(); UserSettings.Invalidate(); }
    public void OnEnableFeesAttribution(IRibbonControl _, bool pressed) { _cfg.EnableFeesAttribution  = pressed; _cfg.Save(); UserSettings.Invalidate(); }
    public void OnEnableLiveData(IRibbonControl _, bool pressed)        { _cfg.EnableLiveData         = pressed; _cfg.Save(); UserSettings.Invalidate(); }

    // ── Error display mode ────────────────────────────────────────────────────

    public bool GetErrorsAsExcelError(IRibbonControl _) => _cfg.ErrorsAsExcelError;

    public void OnErrorsAsExcelError(IRibbonControl _, bool pressed) { _cfg.ErrorsAsExcelError = pressed; _cfg.Save(); UserSettings.Invalidate(); }

    // ── Utility buttons ───────────────────────────────────────────────────────

    public void OnResetDefaults(IRibbonControl _)
    {
        UserSettings.ResetToDefaults();
        UserSettings.Invalidate();
        _cfg = UserSettings.Load();
        _ribbon?.Invalidate();
    }

    public void OnOpenDocs(IRibbonControl _)
    {
        try { Process.Start(new ProcessStartInfo("https://github.com/aleksejcupic/aleksej-finance-excel") { UseShellExecute = true }); }
        catch { /* ignore if browser unavailable */ }
    }

    public void OnAbout(IRibbonControl _)
    {
        // Use ExcelDNA's native alert (no Windows.Forms dependency required)
        XlCall.Excel(XlCall.xlcAlert,
            $"Excel Finance Add-In  v{AddInVersion}\n\nPowered by Aleksej.Finance\n\n" +
            "Author: Aleksej Cupic\naleksejcupic.com", 2);
    }

    /// <summary>Live version label shown in the ribbon (e.g. "v1.0.0").</summary>
    public string GetVersionLabel(IRibbonControl _) => $"v{AddInVersion}";

    /// <summary>
    /// The add-in version, read from the assembly's informational version (set by
    /// &lt;Version&gt; in the csproj). Build metadata after '+' is stripped.
    /// </summary>
    private static string AddInVersion
    {
        get
        {
            Assembly asm = typeof(FinanceRibbon).Assembly;
            string? info = asm.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
            if (!string.IsNullOrEmpty(info))
            {
                int plus = info.IndexOf('+');
                return plus >= 0 ? info[..plus] : info;
            }
            return asm.GetName().Version?.ToString() ?? "1.0.0";
        }
    }

    // ── Ribbon XML ────────────────────────────────────────────────────────────

    private const string RibbonXml = @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
  <ribbon>
    <tabs>
      <tab id='tabFinance' label='Finance Tools'>

        <!-- Default Parameters group -->
        <group id='grpDefaults' label='Default Parameters'>
          <editBox id='ebRiskFreeRate'   label='Risk-Free Rate'
                   getText='GetRiskFreeRate' onChange='OnRiskFreeRateChange'
                   sizeString='0.0500' screentip='Annual risk-free rate (e.g. 0.05 = 5%)'/>
          <comboBox id='cbTradingDays'   label='Trading Days'
                    getSelectedItemIndex='GetTradingDaysIndex' onAction='OnTradingDaysChange'
                    screentip='Trading days per year for annualising returns and volatility'>
            <item id='td252' label='252 (US Equity)'/>
            <item id='td260' label='260 (Europe)'/>
            <item id='td365' label='365 (Calendar)'/>
          </comboBox>
          <comboBox id='cbConfidence'    label='VaR Confidence'
                    getSelectedItemIndex='GetConfidenceIndex' onAction='OnConfidenceChange'
                    screentip='Default confidence level for VaR calculations'>
            <item id='c95'  label='95%'/>
            <item id='c99'  label='99%'/>
            <item id='c999' label='99.9%'/>
          </comboBox>
          <comboBox id='cbFrequency'     label='Coupon Frequency'
                    getSelectedItemIndex='GetFrequencyIndex' onAction='OnFrequencyChange'
                    screentip='Default coupon payment frequency per year'>
            <item id='freq1'  label='Annual (1)'/>
            <item id='freq2'  label='Semi-Annual (2)'/>
            <item id='freq4'  label='Quarterly (4)'/>
            <item id='freq12' label='Monthly (12)'/>
          </comboBox>
          <editBox id='ebLambda'         label='EWMA Lambda'
                   getText='GetLambda' onChange='OnLambdaChange'
                   sizeString='0.9400' screentip='EWMA decay factor (0.94 = RiskMetrics standard)'/>
          <checkBox id='chkErrorsAsExcel' label='Errors as Excel errors (#NUM!/#VALUE!)'
                    getPressed='GetErrorsAsExcelError' onAction='OnErrorsAsExcelError'
                    screentip='When on, invalid inputs return #NUM!/#VALUE! so IFERROR can catch them. When off, a descriptive message is shown in the cell.'/>
        </group>

        <!-- Function Category Toggles group -->
        <group id='grpCategories' label='Active Function Categories'>
          <checkBox id='chkOptions'         label='Options (BS_, GK_, BT_, MC_, EX_, OF_)'
                    getPressed='GetEnableOptions'         onAction='OnEnableOptions'/>
          <checkBox id='chkBonds'           label='Bonds (BOND_, YC_, MORT_)'
                    getPressed='GetEnableBonds'           onAction='OnEnableBonds'/>
          <checkBox id='chkDerivatives'     label='Derivatives (FWD_, FRA_, IRS_, BM_, SR_)'
                    getPressed='GetEnableDerivatives'     onAction='OnEnableDerivatives'/>
          <checkBox id='chkCredit'          label='Credit (CR_)'
                    getPressed='GetEnableCredit'          onAction='OnEnableCredit'/>
          <checkBox id='chkPortfolioRisk'   label='Portfolio &amp; Risk (PORT_, RISK_, VOL_)'
                    getPressed='GetEnablePortfolioRisk'   onAction='OnEnablePortfolioRisk'/>
          <checkBox id='chkFeesAttr'        label='Fees &amp; Attribution (FEE_, ATTR_, EQ_)'
                    getPressed='GetEnableFeesAttribution' onAction='OnEnableFeesAttribution'/>
          <checkBox id='chkLiveData'        label='Live Data (EDGAR_, TREASURY_)'
                    getPressed='GetEnableLiveData'        onAction='OnEnableLiveData'/>
        </group>

        <!-- Help and Docs group -->
        <group id='grpHelp' label='Help &amp; Docs'>
          <button id='btnDocs'   label='Documentation ↗' onAction='OnOpenDocs'
                  screentip='Open the online function reference in your browser'/>
          <button id='btnReset'  label='Reset to Defaults' onAction='OnResetDefaults'
                  screentip='Restore all settings to their factory defaults'/>
          <button id='btnAbout'  label='About' onAction='OnAbout'
                  screentip='Version information and author details'/>
          <labelControl id='lblVersion' getLabel='GetVersionLabel'/>
        </group>

      </tab>
    </tabs>
  </ribbon>
</customUI>";
}
