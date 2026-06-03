namespace Aleksej.Finance.Excel.Constants;

/// <summary>Names, descriptions, help topics, and argument text for Black's model (BM_*) functions.</summary>
internal static class BlackModelConstants
{
    public const string Help = Cat.HelpBase + "/derivatives/black-model";

    public const string CapletName = "BM_CAPLET";
    public const string CapletDesc = "Single caplet (or floorlet) price using Black's model. Pays max(F-K,0)*delta*notional at reset. Set isFloor=TRUE for floorlet.";

    public const string CapName = "BM_CAP";
    public const string CapDesc = "Interest rate cap price — sum of caplets over the payment schedule. All arrays must be the same length.";

    public const string FloorName = "BM_FLOOR";
    public const string FloorDesc = "Interest rate floor price — sum of floorlets. Same inputs as BM_CAP.";

    public const string FwdSwapRateName = "BM_FWD_SWAP_RATE";
    public const string FwdSwapRateDesc = "Forward swap rate R = (P(t0) - P(tn)) / sum(delta_i * P(ti)). This is the ATM swaption strike.";

    public const string SwaptionName = "BM_SWAPTION";
    public const string SwaptionDesc = "Swaption price via Black's model. Payer = right to pay fixed K; Receiver = right to receive fixed K.";

    // Argument descriptions (non-common)
    public const string Notional        = "Notional principal";
    public const string ForwardRate     = "Forward interest rate for the period (annual)";
    public const string StrikeCapFloor  = "Cap/floor strike rate (annual)";
    public const string StrikeCap       = "Cap strike rate";
    public const string StrikeFloor     = "Floor strike rate";
    public const string StrikeSwap      = "Fixed strike rate of the underlying swap";
    public const string TCaplet         = "Time to start of accrual period (option expiry)";
    public const string TSwaption       = "Swaption expiry in years (= swap start date)";
    public const string RZero           = "Zero rate to T (continuously compounded)";
    public const string SigmaForward    = "Black volatility of the forward rate";
    public const string SigmaFlat       = "Flat Black volatility";
    public const string SigmaSwap       = "Black vol of the forward swap rate";
    public const string AccrualFraction = "Accrual period length in years (e.g. 0.5 for semi-annual)";
    public const string IsFloor         = "TRUE for floorlet, FALSE for caplet (default)";
    public const string PaymentTimesReset = "Reset/start times of each caplet (range)";
    public const string PaymentTimesResetShort = "Reset times (range)";
    public const string PaymentTimesSwapExpiry = "Swap payment times in years (range, t0 = swaption expiry)";
    public const string PaymentTimesSwapCoupon = "Swap coupon payment times (range, starting at T)";
    public const string ZeroRates       = "Zero rates at each payment time (range)";
    public const string ZeroRatesShort  = "Zero rates (range)";
    public const string ForwardRates    = "Forward rates for each period (range)";
    public const string ForwardRatesShort = "Forward rates (range)";
    public const string AccrualFracs    = "Accrual fractions for each period (range)";
    public const string AccrualFracsShort = "Accrual fractions (range)";
    public const string IsPayer         = "TRUE = payer swaption (right to pay fixed), FALSE = receiver";

    public const double AccrualFractionDefault = 0.5;
    public const bool IsPayerDefault = true;
}
