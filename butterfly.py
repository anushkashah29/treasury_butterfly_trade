"""
Treasury Butterfly Mean-Reversion Backtest  --  Long Belly Only
Period   : Jun 2016 - Jun 2026
Z-window : 252-day rolling (from full 1976-present history in treasury_rates.csv)

Long belly  | Enter Z >= +2.0  | Exit Z <= 0.0   | Long $100M 5Y / Short $118.4M 2Y / Short $26.5M 10Y

Sizing    : DV01-neutral  $100M 5Y belly
             DV01_BELLY = MD_5Y * $100 * 100  = 4.5 * 100 * 100 = $45,000/bp
             DV01_WING  = DV01_BELLY / 2       = $22,500/bp  (each wing gets half)
             N2  = $22,500 / (1.9 * $100) * $1M = $118.42M
             N10 = $22,500 / (8.5 * $100) * $1M = $26.47M
Financing : SOFR (EFFR proxy pre-Apr-2018)
Min hold  : 30 days  |  One trade at a time (no concurrent positions)

Output files
  butterfly_trades_long.csv   -- long belly trade log
  butterfly_trades.png        -- 2-panel chart
"""

import os
import requests
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.gridspec as gridspec

FRED_API_KEY  = os.environ.get("FRED_API_KEY", "0a7bcf57e2f38f50b9f230c76a4ff850")
FRED_BASE_URL = "https://api.stlouisfed.org/fred/series/observations"

INPUT_FILE  = "treasury_rates.csv"
OUT_LONG    = "butterfly_trades_long.csv"
CHART_FILE  = "butterfly_trades.png"

WINDOW       = 252
ENTRY_Z_LONG =  2.0    # enter long belly when Z >= +2.0
EXIT_Z_LONG  =  0.0    # exit  long belly when Z <= 0.0 (full mean reversion)
MIN_HOLD     = 30
START_DATE   = "2016-06-01"
END_DATE     = "2026-06-30"

# DV01-neutral sizing (fixed par-bond modified durations)
BELLY_N = 100e6
MD = {"2Y": 1.9, "5Y": 4.5, "10Y": 8.5}
DV01_PER_MM = {k: v * 100.0 for k, v in MD.items()}
DV01_BELLY  = DV01_PER_MM["5Y"] * (BELLY_N / 1e6)   # $45,000 / bp
DV01_WING   = DV01_BELLY / 2                          # $22,500 / bp (each wing)
N2          = DV01_WING / DV01_PER_MM["2Y"]  * 1e6   # $118,421,053
N10         = DV01_WING / DV01_PER_MM["10Y"] * 1e6   # $26,470,588


# -- FRED helpers --------------------------------------------------------------

def fetch_fred(series_id, start, end):
    r = requests.get(FRED_BASE_URL, params={
        "series_id": series_id, "observation_start": start,
        "observation_end": end, "api_key": FRED_API_KEY, "file_type": "json",
    }, timeout=30)
    r.raise_for_status()
    obs = r.json()["observations"]
    s = pd.Series(
        {o["date"]: (float(o["value"]) if o["value"] != "." else np.nan) for o in obs},
        name=series_id, dtype="float64",
    )
    s.index = pd.to_datetime(s.index)
    return s


def get_repo_rate(start, end):
    """SOFR from Apr-2018; EFFR (DFF) before that. Returns decimal (e.g. 0.05)."""
    print("  Fetching SOFR...")
    sofr = fetch_fred("SOFR", start, end)
    print("  Fetching EFFR (pre-SOFR proxy)...")
    effr = fetch_fred("DFF", start, end)
    combined = effr.copy()
    combined.update(sofr)       # SOFR overwrites EFFR where available
    combined = combined / 100.0
    return combined.sort_index()


# -- backtest state machine ----------------------------------------------------

def run_backtest(df, entry_z=ENTRY_Z_LONG, exit_z=EXIT_Z_LONG, min_hold=MIN_HOLD):
    """
    Long belly only -- one trade at a time, no concurrent positions.
    Enter Z >= entry_z, exit Z <= exit_z (after min_hold days).
    Uses pre-computed z_score_252 from treasury_rates.csv (full 1976-present history).
    """
    df = df.copy()
    if "z_score_252" in df.columns:
        df["z_score"]   = df["z_score_252"]
        df["roll_mean"] = df["roll_mean_252"]
        df["roll_std"]  = df["roll_std_252"]
    else:
        df["roll_mean"] = df["butterfly_spread"].rolling(WINDOW).mean()
        df["roll_std"]  = df["butterfly_spread"].rolling(WINDOW).std()
        df["z_score"]   = (df["butterfly_spread"] - df["roll_mean"]) / df["roll_std"]

    raw_trades = []
    in_trade   = False
    entry_idx  = None
    hold_days  = 0

    for i in range(len(df)):
        z = df.at[i, "z_score"]
        if pd.isna(z):
            continue

        if not in_trade:
            if z >= entry_z:
                in_trade = True; entry_idx = i; hold_days = 0
        else:
            hold_days += 1
            if hold_days >= min_hold and z <= exit_z:
                e = df.iloc[entry_idx]
                x = df.iloc[i]
                raw_trades.append({
                    "entry_date":      e["date"],
                    "entry_DGS2":      round(e["DGS2"],  3),
                    "entry_DGS5":      round(e["DGS5"],  3),
                    "entry_DGS10":     round(e["DGS10"], 3),
                    "entry_butterfly": round(e["butterfly_spread"], 4),
                    "entry_z_score":   round(e["z_score"], 3),
                    "exit_date":       x["date"],
                    "exit_DGS2":       round(x["DGS2"],  3),
                    "exit_DGS5":       round(x["DGS5"],  3),
                    "exit_DGS10":      round(x["DGS10"], 3),
                    "exit_butterfly":  round(x["butterfly_spread"], 4),
                    "exit_z_score":    round(x["z_score"], 3),
                    "hold_days":       hold_days,
                    "_entry_idx":      entry_idx,
                    "_exit_idx":       i,
                })
                in_trade  = False
                entry_idx = None
                hold_days = 0

    return df, raw_trades


# -- P&L calculation -----------------------------------------------------------

def calc_pnl(raw_trades, df, repo):
    """
    Long belly P&L:
      MtM      = -DV01_WING * butterfly_chg_bps   (profit when butterfly falls)
      leg_5y   = -DV01_BELLY * d5  (long 5Y: lose when yield rises)
      leg_2y   = +DV01_WING  * d2  (short 2Y: gain when yield rises)
      leg_10y  = +DV01_WING  * d10 (short 10Y: gain when yield rises)
      carry/d  = [(Y5-SOFR)*BELLY_N + (SOFR-Y2)*N2 + (SOFR-Y10)*N10] / 360
    """
    results = []

    for t in raw_trades:
        d_butter_pct = t["exit_butterfly"] - t["entry_butterfly"]
        mtm_pnl = -DV01_WING * 100 * d_butter_pct

        d2  = round((t["exit_DGS2"]  - t["entry_DGS2"])  * 100, 1)
        d5  = round((t["exit_DGS5"]  - t["entry_DGS5"])  * 100, 1)
        d10 = round((t["exit_DGS10"] - t["entry_DGS10"]) * 100, 1)

        leg5  = int(round(-DV01_BELLY * d5))
        leg2  = int(round(+DV01_WING  * d2))
        leg10 = int(round(+DV01_WING  * d10))

        entry_dt = pd.Timestamp(t["entry_date"])
        exit_dt  = pd.Timestamp(t["exit_date"])
        mask     = (df["date"] >= entry_dt) & (df["date"] < exit_dt)
        period   = df[mask].set_index("date")
        r   = repo.reindex(period.index).ffill().fillna(0)
        y5  = period["DGS5"]  / 100
        y2  = period["DGS2"]  / 100
        y10 = period["DGS10"] / 100

        daily_carry = (
            (y5 - r) * BELLY_N + (r - y2) * N2 + (r - y10) * N10
        ) / 360

        carry_pnl     = float(daily_carry.sum())
        total_pnl     = mtm_pnl + carry_pnl
        avg_repo_pct  = float(r.mean() * 100)
        avg_daily_car = float(daily_carry.mean())

        t_out = {k: v for k, v in t.items() if not k.startswith("_")}
        t_out.update({
            "delta_dgs2_bps":       d2,
            "delta_dgs5_bps":       d5,
            "delta_dgs10_bps":      d10,
            "leg_5y_pnl_usd":       leg5,
            "leg_2y_pnl_usd":       leg2,
            "leg_10y_pnl_usd":      leg10,
            "legs_sum_pnl_usd":     leg5 + leg2 + leg10,
            "md_2y_assumed":        MD["2Y"],
            "md_5y_assumed":        MD["5Y"],
            "md_10y_assumed":       MD["10Y"],
            "belly_5y_notional_mm": round(BELLY_N / 1e6, 2),
            "wing_2y_notional_mm":  round(N2  / 1e6, 2),
            "wing_10y_notional_mm": round(N10 / 1e6, 2),
            "belly_dv01_per_bp":    int(round(DV01_BELLY)),
            "wing_2y_dv01_per_bp":  int(round(DV01_WING)),
            "wing_10y_dv01_per_bp": int(round(DV01_WING)),
            "avg_repo_rate_pct":    round(avg_repo_pct, 3),
            "avg_daily_carry_usd":  int(round(avg_daily_car)),
            "butterfly_chg_bps":    round(d_butter_pct * 100, 1),
            "mtm_pnl_usd":          int(round(mtm_pnl)),
            "carry_pnl_usd":        int(round(carry_pnl)),
            "total_pnl_usd":        int(round(total_pnl)),
            "profitable":           "Y" if total_pnl > 0 else "N",
        })
        results.append(t_out)

    return pd.DataFrame(results)


# -- chart ---------------------------------------------------------------------

def make_chart(df, long_df):
    fig = plt.figure(figsize=(16, 16))
    gs  = gridspec.GridSpec(3, 1, figure=fig, hspace=0.52, height_ratios=[1.8, 1.2, 1.2])

    dates = df["date"]
    mean  = df["roll_mean"]
    std   = df["roll_std"]

    # Panel 1: Spread + rolling bands + trade shading
    ax1 = fig.add_subplot(gs[0])
    upper = mean + 2.0 * std

    ax1.plot(dates, df["butterfly_spread"], color="grey", lw=0.75, label="Butterfly", zorder=2)
    ax1.plot(dates, mean,  color="#1f77b4", lw=1.1, label=f"{WINDOW}d Mean", zorder=3)
    ax1.plot(dates, upper, color="darkred", lw=0.9, ls="--", label="+2.0s Long entry", zorder=3)
    ax1.plot(dates, mean,  color="red",     lw=0.6, ls=":",  alpha=0.7, label="0.0 Long exit (mean)", zorder=3)

    for _, tr in long_df.iterrows():
        clr = "#2ca02c" if tr["profitable"] == "Y" else "#d62728"
        ax1.axvspan(pd.Timestamp(tr["entry_date"]), pd.Timestamp(tr["exit_date"]), alpha=0.13, color=clr)
        ax1.scatter(pd.Timestamp(tr["entry_date"]), tr["entry_butterfly"], marker="^", s=65, color="#2ca02c", zorder=5)
        ax1.scatter(pd.Timestamp(tr["exit_date"]),  tr["exit_butterfly"],  marker="x", s=55, color="#d62728", zorder=5)

    ax1.set_title(
        f"Treasury Butterfly  --  Long Belly | DV01-Neutral $100M | {WINDOW}d Z-Score | Jun 2016-Jun 2026\n"
        f"[^ green = entry  x red = exit  green shading = win  red shading = loss]",
        fontsize=10, fontweight="bold",
    )
    ax1.set_ylabel("Butterfly Spread (%)")
    ax1.legend(fontsize=7, loc="upper right", ncol=3)
    ax1.grid(True, alpha=0.3)
    ax1.xaxis.set_major_locator(mdates.YearLocator(1))
    ax1.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    plt.setp(ax1.xaxis.get_majorticklabels(), rotation=30)

    # Panel 2: Z-score with entry/exit bands
    ax2 = fig.add_subplot(gs[1])
    ax2.plot(dates, df["z_score"], color="#555555", lw=0.85)
    ax2.axhline(ENTRY_Z_LONG, color="darkred", lw=1.0, ls="--", label=f"+{ENTRY_Z_LONG} Long entry")
    ax2.axhline(EXIT_Z_LONG,  color="red",     lw=0.8, ls=":",  label=f"{EXIT_Z_LONG} Long exit")
    ax2.axhline(0, color="black", lw=0.6, alpha=0.4)
    ax2.fill_between(dates, df["z_score"], ENTRY_Z_LONG,
                     where=(df["z_score"] >= ENTRY_Z_LONG), color="red", alpha=0.15)
    for _, tr in long_df.iterrows():
        ax2.axvspan(pd.Timestamp(tr["entry_date"]), pd.Timestamp(tr["exit_date"]), alpha=0.07, color="green")
    ax2.set_title(f"{WINDOW}d Rolling Z-Score (from full 1976-present history in treasury_rates.csv)",
                  fontsize=10, fontweight="bold")
    ax2.set_ylabel("Z-Score")
    ax2.legend(fontsize=7, loc="upper right", ncol=2)
    ax2.grid(True, alpha=0.3)
    ax2.xaxis.set_major_locator(mdates.YearLocator(1))
    ax2.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    plt.setp(ax2.xaxis.get_majorticklabels(), rotation=30)

    # Panel 3: Cumulative P&L
    ax3 = fig.add_subplot(gs[2])
    long_x   = pd.to_datetime(long_df["exit_date"])
    cum_long = long_df["total_pnl_usd"].cumsum() / 1e6
    bar_colors = ["#2ca02c" if p == "Y" else "#d62728" for p in long_df["profitable"]]

    ax3.bar(long_x, long_df["total_pnl_usd"] / 1e6, width=12, alpha=0.6, color=bar_colors)
    ax3.step(long_x, cum_long, color="black", lw=1.6, label="Cumul. P&L", where="post")
    ax3.axhline(0, color="black", lw=0.7)
    ax3.set_title(
        "Cumulative P&L by Trade Close Date (USD millions)\n"
        "[Green bars = wins  Red bars = losses]",
        fontsize=10, fontweight="bold",
    )
    ax3.set_ylabel("USD millions")
    ax3.legend(fontsize=8, loc="upper left")
    ax3.grid(True, alpha=0.3)
    ax3.xaxis.set_major_locator(mdates.YearLocator(1))
    ax3.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    plt.setp(ax3.xaxis.get_majorticklabels(), rotation=30)

    plt.savefig(CHART_FILE, dpi=150, bbox_inches="tight")
    plt.close()
    print(f"Chart saved to {CHART_FILE}")


# -- main ----------------------------------------------------------------------

def main():
    df = pd.read_csv(INPUT_FILE, parse_dates=["date"])
    df = df[(df["date"] >= START_DATE) & (df["date"] <= END_DATE)].copy()
    df.sort_values("date", inplace=True)
    df.reset_index(drop=True, inplace=True)
    print(f"Loaded {len(df)} trading days ({START_DATE} to {END_DATE})")

    print("Fetching repo rates...")
    repo = get_repo_rate(START_DATE, END_DATE)

    print("Running backtest...")
    df, raw_trades = run_backtest(df)
    print(f"  Long belly trades: {len(raw_trades)}")

    if not raw_trades:
        print("No trades found. Exiting.")
        return

    long_df = calc_pnl(raw_trades, df, repo)

    wins = (long_df["profitable"] == "Y").sum()
    n    = len(long_df)
    print(f"\n{'='*65}")
    print(f"  Treasury Butterfly | DV01-Neutral | $100M 5Y belly")
    print(f"  Window  : {WINDOW}d Z-score (full 1976-present history)")
    print(f"  Long    : enter Z >= +{ENTRY_Z_LONG}  |  exit Z <= {EXIT_Z_LONG}")
    print(f"  Min hold: {MIN_HOLD} days  |  One trade at a time")
    print(f"{'='*65}")
    print(f"\n  Long Belly  (Long 5Y / Short 2Y, 10Y)")
    print(f"  Trades     : {n}  |  Win rate: {wins}/{n} ({wins/n*100:.0f}%)")
    print(f"  Total MtM  : ${long_df['mtm_pnl_usd'].sum()/1e6:.3f}M")
    print(f"  Total Carry: ${long_df['carry_pnl_usd'].sum()/1e6:.3f}M")
    print(f"  Total P&L  : ${long_df['total_pnl_usd'].sum()/1e6:.3f}M")
    print(f"{'='*65}")

    cols = [
        "entry_date", "entry_DGS2", "entry_DGS5", "entry_DGS10",
        "entry_butterfly", "entry_z_score",
        "exit_date", "exit_DGS2", "exit_DGS5", "exit_DGS10",
        "exit_butterfly", "exit_z_score",
        "hold_days",
        "delta_dgs2_bps", "delta_dgs5_bps", "delta_dgs10_bps",
        "leg_5y_pnl_usd", "leg_2y_pnl_usd", "leg_10y_pnl_usd", "legs_sum_pnl_usd",
        "md_2y_assumed", "md_5y_assumed", "md_10y_assumed",
        "belly_5y_notional_mm", "wing_2y_notional_mm", "wing_10y_notional_mm",
        "belly_dv01_per_bp", "wing_2y_dv01_per_bp", "wing_10y_dv01_per_bp",
        "avg_repo_rate_pct", "avg_daily_carry_usd",
        "butterfly_chg_bps",
        "mtm_pnl_usd", "carry_pnl_usd", "total_pnl_usd", "profitable",
    ]

    long_df[cols].to_csv(OUT_LONG, index=False)
    print(f"\nLong trades saved: {OUT_LONG}")

    make_chart(df, long_df)

    pd.set_option("display.max_columns", None)
    pd.set_option("display.width", 280)
    pd.set_option("display.float_format", "{:.3f}".format)
    print("\n--- Long Belly Trade Details ---")
    print(long_df[cols].to_string(index=False))


if __name__ == "__main__":
    main()
