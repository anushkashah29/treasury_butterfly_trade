"""
Parameter sweep for the Treasury Butterfly Long Belly strategy.

Sweeps entry/exit Z-score thresholds (min hold fixed at 30 days, same as
butterfly.py) to explore the trade-count vs P&L tradeoff: does loosening the
signal to generate more trades actually raise total P&L, or does per-trade
edge decay faster than volume grows?

Reuses run_backtest / calc_pnl / get_repo_rate from butterfly.py unchanged --
only the entry_z / exit_z thresholds vary per combo.

Output files
  param_sweep_results.csv   -- one row per (entry_z, exit_z) combo
  param_sweep_scatter.png   -- trade count vs total P&L, colored by win rate
  param_sweep_heatmap.png   -- entry_z x exit_z grid, color = total P&L, cell label = trade count
"""

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

from butterfly import (
    INPUT_FILE, START_DATE, END_DATE, MIN_HOLD,
    ENTRY_Z_LONG, EXIT_Z_LONG,
    run_backtest, calc_pnl, get_repo_rate,
)

ENTRY_Z_GRID = np.round(np.arange(1.0, 3.01, 0.25), 2)
EXIT_Z_GRID  = np.round(np.arange(-1.5, 1.51, 0.25), 2)

OUT_CSV      = "param_sweep_results.csv"
SCATTER_FILE = "param_sweep_scatter.png"
HEATMAP_FILE = "param_sweep_heatmap.png"


def load_data():
    df = pd.read_csv(INPUT_FILE, parse_dates=["date"])
    df = df[(df["date"] >= START_DATE) & (df["date"] <= END_DATE)].copy()
    df.sort_values("date", inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df


def sweep(df, repo):
    rows = []
    for entry_z in ENTRY_Z_GRID:
        for exit_z in EXIT_Z_GRID:
            if exit_z >= entry_z:
                continue

            _, raw_trades = run_backtest(df, entry_z=entry_z, exit_z=exit_z, min_hold=MIN_HOLD)

            if not raw_trades:
                rows.append({
                    "entry_z": entry_z, "exit_z": exit_z, "min_hold": MIN_HOLD,
                    "n_trades": 0, "win_rate": np.nan,
                    "total_mtm": 0, "total_carry": 0, "total_pnl": 0,
                    "avg_pnl_per_trade": np.nan,
                })
                continue

            trades_df = calc_pnl(raw_trades, df, repo)
            n = len(trades_df)
            wins = int((trades_df["profitable"] == "Y").sum())

            rows.append({
                "entry_z": entry_z,
                "exit_z": exit_z,
                "min_hold": MIN_HOLD,
                "n_trades": n,
                "win_rate": wins / n,
                "total_mtm": trades_df["mtm_pnl_usd"].sum(),
                "total_carry": trades_df["carry_pnl_usd"].sum(),
                "total_pnl": trades_df["total_pnl_usd"].sum(),
                "avg_pnl_per_trade": trades_df["total_pnl_usd"].mean(),
            })

    return pd.DataFrame(rows)


def plot_scatter(results):
    valid = results[results["n_trades"] > 0]

    fig, ax = plt.subplots(figsize=(10, 7))
    sc = ax.scatter(
        valid["n_trades"], valid["total_pnl"] / 1e6,
        c=valid["win_rate"], cmap="RdYlGn", vmin=0, vmax=1,
        s=60, edgecolor="black", linewidth=0.4, zorder=3,
    )
    cbar = plt.colorbar(sc, ax=ax)
    cbar.set_label("Win Rate")

    baseline = results[(results["entry_z"] == ENTRY_Z_LONG) & (results["exit_z"] == EXIT_Z_LONG)]
    if not baseline.empty and baseline["n_trades"].iloc[0] > 0:
        ax.scatter(
            baseline["n_trades"], baseline["total_pnl"] / 1e6,
            marker="*", s=450, color="blue", edgecolor="black", linewidth=0.6,
            zorder=5, label=f"Current strategy (Z={ENTRY_Z_LONG}/{EXIT_Z_LONG})",
        )
        ax.legend(loc="best")

    ax.axhline(0, color="black", lw=0.8, ls="--", alpha=0.5)
    ax.set_xlabel("Number of Trades")
    ax.set_ylabel("Total P&L ($M)")
    ax.set_title(
        "Trade Count vs Total P&L Across Entry/Exit Z-Score Thresholds\n"
        f"(min hold = {MIN_HOLD}d, {START_DATE} to {END_DATE})",
        fontsize=12, fontweight="bold",
    )
    ax.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig(SCATTER_FILE, dpi=150)
    plt.close()
    print(f"Saved {SCATTER_FILE}")


def plot_heatmap(results):
    pivot_pnl = results.pivot(index="entry_z", columns="exit_z", values="total_pnl") / 1e6
    pivot_n   = results.pivot(index="entry_z", columns="exit_z", values="n_trades")

    data = np.ma.masked_invalid(pivot_pnl.values)
    vmax = np.nanmax(np.abs(pivot_pnl.values))

    cmap = plt.get_cmap("RdYlGn").copy()
    cmap.set_bad("#eeeeee")

    fig, ax = plt.subplots(figsize=(14, 8))
    im = ax.imshow(data, cmap=cmap, aspect="auto", vmin=-vmax, vmax=vmax)

    ax.set_xticks(range(len(pivot_pnl.columns)))
    ax.set_xticklabels(pivot_pnl.columns, rotation=45)
    ax.set_yticks(range(len(pivot_pnl.index)))
    ax.set_yticklabels(pivot_pnl.index)
    ax.set_xlabel("Exit Z-Score")
    ax.set_ylabel("Entry Z-Score")
    ax.set_title(
        f"Total P&L ($M) by Entry/Exit Z-Score  (cell label = trade count, blank = invalid/no trades)\n"
        f"min hold = {MIN_HOLD}d, {START_DATE} to {END_DATE}",
        fontsize=12, fontweight="bold",
    )

    for i in range(pivot_n.shape[0]):
        for j in range(pivot_n.shape[1]):
            n = pivot_n.values[i, j]
            if not np.isnan(n):
                ax.text(j, i, f"{int(n)}", ha="center", va="center", fontsize=7)

    cbar = plt.colorbar(im, ax=ax)
    cbar.set_label("Total P&L ($M)")
    plt.tight_layout()
    plt.savefig(HEATMAP_FILE, dpi=150)
    plt.close()
    print(f"Saved {HEATMAP_FILE}")


def main():
    df = load_data()
    print(f"Loaded {len(df)} trading days ({START_DATE} to {END_DATE})")

    print("Fetching repo rates...")
    repo = get_repo_rate(START_DATE, END_DATE)

    print(f"Sweeping {len(ENTRY_Z_GRID)} entry x {len(EXIT_Z_GRID)} exit thresholds "
          f"(min hold fixed at {MIN_HOLD}d)...")
    results = sweep(df, repo)
    results.to_csv(OUT_CSV, index=False)
    print(f"Saved {len(results)} combos to {OUT_CSV}")

    valid = results[results["n_trades"] > 0]
    pd.set_option("display.width", 200)
    pd.set_option("display.float_format", "{:.3f}".format)

    print("\nTop 10 combos by total P&L:")
    print(valid.sort_values("total_pnl", ascending=False).head(10).to_string(index=False))

    baseline = results[(results["entry_z"] == ENTRY_Z_LONG) & (results["exit_z"] == EXIT_Z_LONG)]
    if not baseline.empty:
        print(f"\nCurrent strategy (Z={ENTRY_Z_LONG}/{EXIT_Z_LONG}) for comparison:")
        print(baseline.to_string(index=False))

    plot_scatter(results)
    plot_heatmap(results)


if __name__ == "__main__":
    main()
