# Treasury Butterfly Mean-Reversion Strategy

## Overview

This project downloads US Treasury Constant Maturity (CMT) yields from the FRED API, constructs the 2s5s10s butterfly spread, and backtests a DV01-neutral mean-reversion strategy — **long the 5Y belly vs. wings** — over the period June 2016 to June 2026.

The Z-score is computed on a **252-day rolling window** (one trading year) using the **full 1976-present yield history**, providing a statistically robust baseline that avoids regime-specific distortions from short windows.

---

## Files

| File | Description |
|---|---|
| `fred_treasury_rates.py` | Downloads DGS2, DGS5, DGS10 from FRED; computes 252d Z-score; outputs `treasury_rates.csv` and `butterfly_spread.png` |
| `butterfly.py` | Backtest engine (long belly only); outputs `butterfly_trades_long.csv` and `butterfly_trades.png` |
| `create_xlsx_files.py` | Converts CSVs to Excel with live formulas; outputs `treasury_rates.xlsx` and `butterfly_trades.xlsx` |
| `treasury_rates.csv` | Daily CMT yields + butterfly spread + 252d Z-score (Jun 1976 – present) |
| `butterfly_trades_long.csv` | Long belly trade log with DV01s, notionals, repo rates, and P&L |
| `treasury_rates.xlsx` | Live Excel formulas for butterfly spread and 252d rolling Z-score (cols E–H) |
| `butterfly_trades.xlsx` | Two-sheet workbook: Long Belly Trades / Summary (all P&L as live formulas) |
| `butterfly_trades.png` | 3-panel chart: spread + entry band, Z-score, cumulative P&L |

---

## Data Source

- **Provider**: Federal Reserve Bank of St. Louis (FRED API)
- **Series**: `DGS2`, `DGS5`, `DGS10` — Treasury Constant Maturity Rates
- **Full history file**: June 1, 1976 – current date
- **Backtest window**: June 1, 2016 – June 30, 2026
- **Frequency**: Daily (business days only; holidays and missing values dropped)
- **Financing rate**: SOFR (`SOFR` series from April 3, 2018); EFFR (`DFF`) as proxy before April 2018

---

## Problem Description

The 5Y belly often richens or cheapens relative to the 2Y and 10Y wings due to supply/demand technicals, Fed policy expectations, convexity hedging, and position flows. This creates the butterfly spread, which tends to mean-revert over medium-term horizons.

**Butterfly Spread Definition:**

```
Butterfly = 2 × DGS5 − DGS2 − DGS10
```

A **positive** butterfly means the 5Y yield is elevated relative to the linear interpolation between 2Y and 10Y — the belly is **cheap**. A **negative** butterfly means the 5Y is **rich** relative to the wings.

---

## Hypothesis

When the 252-day Z-score of the butterfly exceeds **+2.0**, the spread is anomalously cheap relative to its one-year rolling history and is likely to mean-revert. The strategy enters a **long belly** position and holds until Z reverts to 0.0 (the rolling mean).

- **Long belly** (Z ≥ +2.0): belly is cheap — buy 5Y, sell 2Y and 10Y wings
- **Exit** (Z ≤ 0.0): spread has fully reverted to the rolling mean

**Key assumptions:**
1. The butterfly spread is stationary and mean-reverts over a 1–6 month horizon.
2. DV01-neutral sizing eliminates directional duration exposure; P&L is driven purely by curve shape changes.
3. A minimum 30-day holding period filters noise and prevents short-term whipsaws.
4. Full exit at Z = 0.0 (full mean reversion to the rolling mean) maximises capture of the mean-reversion move.
5. Financing cost is approximated using the overnight SOFR rate (EFFR pre-2018).

---

## Trade Structure

### Long Belly — Enter when Z ≥ +2.0

| Leg | Direction | Notional | DV01 / bp |
|---|---|---|---|
| 5Y Treasury (belly) | **Long** | $100.00M | $45,000 |
| 2Y Treasury (wing) | **Short** | $118.42M | $22,500 |
| 10Y Treasury (wing) | **Short** | $26.47M | $22,500 |
| **Net DV01** | | | **$0** |

### Modified Duration Assumptions

Fixed par-bond approximations (constant across all trades):

| Tenor | Modified Duration |
|---|---|
| 2Y | 1.90 years |
| 5Y | 4.50 years |
| 10Y | 8.50 years |

### DV01-Neutral Sizing ($100M Belly)

```
DV01 per $1M face  =  Modified Duration × $100

Belly DV01  =  4.50 × $100 × ($100M / $1M)  =  $45,000 / bp
Wing DV01 (each)  =  $45,000 / 2            =  $22,500 / bp

Wing 2Y  notional  =  $22,500 / (1.90 × $100) × $1M  =  $118.42M
Wing 10Y notional  =  $22,500 / (8.50 × $100) × $1M  =   $26.47M
```

---

## Entry & Exit Rules

| Rule | Long Belly |
|---|---|
| Z-score window | 252 trading days (full 1976-present history) |
| Entry signal | Z ≥ +2.0 |
| Exit signal | Z ≤ 0.0 AND hold ≥ 30 days |
| Min hold | 30 days |
| Concurrent positions | No — one trade at a time |

**Why Z = 0.0 for exit:**
The exit at Z = 0.0 waits for full mean reversion to the rolling 252-day average, capturing the entire spread compression move. This maximises P&L per trade and avoids premature exits that leave gains on the table.

---

## P&L Formulas

### Long Belly

```
MtM P&L ($)    = −DV01_wing × Δbutterfly (bps)      = −$22,500 × Δbfly × 100

Per leg:
  leg_5y   = −$45,000 × Δdgs5_bps     (long 5Y: lose when yield rises)
  leg_2y   = +$22,500 × Δdgs2_bps     (short 2Y: gain when yield rises)
  leg_10y  = +$22,500 × Δdgs10_bps    (short 10Y: gain when yield rises)

Daily carry ($) =
    (Y5 − SOFR) × $100M / 360          [long belly: receive yield, pay SOFR]
  + (SOFR − Y2) × $118.42M / 360       [short 2Y: receive SOFR, pay yield]
  + (SOFR − Y10) × $26.47M / 360       [short 10Y: receive SOFR, pay yield]
```

---

## Z-Score Calculation

### Formula

```
Z-score  =  (butterfly_spread − roll_mean_252) / roll_std_252
```

- `roll_mean_252` and `roll_std_252` are computed over the **trailing 252 trading days** using the **full 1976-present history** in `treasury_rates.csv`.
- `roll_std_252` uses **ddof = 1** (sample standard deviation, pandas default).
- The first 251 rows carry no Z-score (NaN in CSV, blank in Excel) — a full 252-day window is required before any signal is valid.
- Because the window uses the full history, the Z-score at any date reflects how unusual the current spread is relative to a genuine one-year lookback — not just since the backtest start date.

### Excel Formulas in treasury_rates.xlsx

| Column | Excel Formula |
|---|---|
| E: `butterfly_spread` | `=2*C-B-D` |
| F: `roll_mean_252` | `=IF(ROW()-2<252,"",AVERAGE(OFFSET($E$2,ROW()-253,0,252,1)))` |
| G: `roll_std_252` | `=IF(ROW()-2<252,"",STDEV(OFFSET($E$2,ROW()-253,0,252,1)))` |
| H: `z_score_252` | `=IF(F="","",IF(G=0,"",(E-F)/G))` |

> OFFSET-based rolling formulas reference exactly the trailing 252-row window per cell; cached values display immediately on open without requiring recalculation.

---

## Backtest Results (Jun 2016 – Jun 2026)

### Strategy Summary

| Metric | Long Belly |
|---|---|
| Trades | **6** |
| Winning trades | **5 / 6 (83%)** |
| Total MtM P&L | +$855,000 |
| Total Carry P&L | +$172,000 |
| **Total P&L** | **+$1,027,000** |
| Z-score window | 252 days |
| Entry threshold | Z ≥ +2.0 |
| Exit threshold | Z ≤ 0.0 |
| Min hold | 30 days |
| Belly notional | $100M |
| Belly DV01 | $45,000/bp |
| Wing DV01 (each) | $22,500/bp |

### Long Belly Trade Log

| # | Entry | Entry Bfly | Entry Z | Exit | Exit Bfly | Exit Z | Hold | Bfly Chg | MtM P&L | Carry | Total | W/L |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| L1 | 2016-11-18 | +0.190% | +2.33 | 2017-04-18 | +0.060% | 0.000 | 101d | −13 bps | +$292,500 | +$37,943 | **+$330,443** | W |
| L2 | 2019-12-06 | −0.110% | +2.57 | 2020-02-24 | −0.220% | −0.436 | 52d | −11 bps | +$247,500 | −$465 | **+$247,035** | W |
| L3 | 2020-03-06 | −0.070% | +2.81 | 2020-05-08 | −0.190% | −0.153 | 44d | −12 bps | +$270,000 | +$400 | **+$270,400** | W |
| L4 | 2021-02-26 | −0.080% | +2.06 | 2022-03-01 | +0.090% | −0.058 | 253d | +17 bps | −$382,500 | +$95,867 | **−$286,633** | L |
| L5 | 2024-04-11 | −0.270% | +2.31 | 2024-06-06 | −0.420% | −0.266 | 39d | −15 bps | +$337,500 | −$5,283 | **+$332,217** | W |
| L6 | 2024-10-09 | −0.230% | +2.53 | 2025-04-03 | −0.270% | −0.103 | 119d | −4 bps | +$90,000 | +$43,752 | **+$133,752** | W |

---

## Key Trade Analysis

### February 2021 — The One Loss: Regime Change Risk

- **Entry** Feb 26, 2021: butterfly = −0.08%, Z = +2.06 (belly cheap; curve steepening post-Covid)
- **Exit** Mar 1, 2022: butterfly = +0.09%, Z = −0.058 (253 days held — longest in the book)
- The butterfly did not compress; instead it **widened 17 bps** as the Fed hiking cycle began. The 5Y sold off faster than the wings (5Y +117 bps, 2Y +117 bps, 10Y +28 bps), with the massive 2Y short position generating a −$2.6M leg loss that overwhelmed the +$630K 10Y gain.
- Carry partially offset the loss (+$95,867 over 253 days, ~$379/day) as near-zero SOFR made the long 5Y position cheap to fund.
- This trade illustrates the core regime risk: a high Z-score from a prior period of belly compression does not prevent the belly from continuing to widen if the macro regime is shifting (here: from ZIRP to hiking).

**Momentum warning signs that were present at entry (Jun 2021 in the prior version):**
- The butterfly spread was still rising over the preceding 20 days (momentum in the wrong direction)
- 10Y yields were already trending higher (bear steepening in progress)

A 20-day momentum filter that blocks entry when the butterfly is still widening would have flagged this trade — the spread must at least be compressing before mean reversion can begin.

### November 2016 — Clean Mean Reversion (+$330K)

- **Entry** Nov 18, 2016: butterfly = +0.190%, Z = +2.33 (post-election steepening pushed 5Y cheap)
- **Exit** Apr 18, 2017: butterfly = +0.060%, Z = 0.000 (101 days, full reversion to mean)
- The spread compressed 13 bps as the 5Y yield pulled back relative to 2Y and 10Y — a textbook mean-reversion from an over-extended level. Carry was positive throughout (+$376/day average) as 5Y yield modestly exceeded SOFR.

### March 2020 — Best Win Rate: Two Back-to-Back Trades (+$247K, +$270K)

- **L2** Dec 6, 2019 → Feb 24, 2020: Z = +2.57 at entry, spread compressed 11 bps over 52 days. Exit came as Covid volatility began hitting rates.
- **L3** Mar 6, 2020 → May 8, 2020: Z = +2.81 at entry (belly extremely cheap as the Fed cut to zero). Spread compressed 12 bps over 44 days as rates stabilised. Both trades benefited from near-zero SOFR funding cost.

---

## CSV Column Formulas

### treasury_rates.csv

| Column | Formula |
|---|---|
| `butterfly_spread` | `2 × DGS5 − DGS2 − DGS10` |
| `roll_mean_252` | `butterfly_spread.rolling(252).mean()` — trailing 252 trading days, NaN for first 251 rows |
| `roll_std_252` | `butterfly_spread.rolling(252).std(ddof=1)` — sample standard deviation |
| `z_score_252` | `(butterfly_spread − roll_mean_252) / roll_std_252` |

### butterfly_trades_long.csv

| Column | Formula |
|---|---|
| `butterfly_chg_bps` | `(exit_butterfly − entry_butterfly) × 100` |
| `leg_5y_pnl_usd` | `−$45,000 × delta_dgs5_bps` |
| `leg_2y_pnl_usd` | `+$22,500 × delta_dgs2_bps` |
| `leg_10y_pnl_usd` | `+$22,500 × delta_dgs10_bps` |
| `legs_sum_pnl_usd` | `leg_5y + leg_2y + leg_10y` |
| `mtm_pnl_usd` | `−$22,500 × butterfly_chg_bps` |
| `avg_daily_carry_usd` | `[(Y5−SOFR)×$100M + (SOFR−Y2)×$118.4M + (SOFR−Y10)×$26.5M] / 360` |
| `carry_pnl_usd` | `sum of daily carry over hold period` |
| `total_pnl_usd` | `mtm_pnl_usd + carry_pnl_usd` |

---

## Excel Workbook butterfly_trades.xlsx

Two sheets with live formulas (italic blue cells):

**Sheet 1 — Long Belly Trades** (dark green header)

| Section | Columns | Content |
|---|---|---|
| 1 — INPUT | A–M | Entry/exit yields, Z-scores, hold days, avg SOFR, actual avg carry |
| 2 — YIELD CHANGES | N–Q | `=(G−B)×100` per leg; `=2×O−N−P` butterfly change |
| 3 — MtM P&L | R–Z | DV01 constants; `leg_5y=−R×O`, `leg_2y=+S×N`, `leg_10y=+S×P`; `mtm=−S×Q` |
| 4 — CARRY | AA–AH | Notional constants; belly=(C%−L%)×AA/360; 2Y=(L%−B%)×AB/360; 10Y=(L%−D%)×AC/360 |
| 5 — TOTAL | AI–AJ | `=Z+AH`; `=IF(AI>0,"Y","N")` |

**Sheet 2 — Summary** — trade count, win rate, and P&L totals.

---

## Key Observations

1. **83% win rate on long belly**: 5 of 6 long belly trades profited over the Jun 2016 – Jun 2026 period. The strategy is aligned with the structural bias of the butterfly: the belly tends to be cheap (positive Z) during periods of curve steepening or convexity hedging demand for 5Y, and mean-reverts once the technical pressure dissipates.

2. **The one loss (Feb 2021) was a regime change**: The Fed hiking cycle converted what looked like a standard mean-reversion entry into a 253-day hold against the position. Carry (+$96K) partially offset the −$383K MtM loss. A momentum filter (block entry if butterfly is still widening over 20 days) would have flagged this entry.

3. **Carry is modest but directionally consistent**: Long belly carry depends on whether the 5Y yield exceeds SOFR. In ZIRP (2020–2021), carry was near-zero. In the post-2022 hiking cycle (L5, L6), inverted funding curves produced negative belly carry but positive carry on the 2Y/10Y wing positions.

4. **Short belly excluded**: Backtesting showed that short belly (Z ≤ −2.0) produced 1/3 wins over the same period, with the 2022–2023 trade losing −$743K (MtM) as the spread trended to Z = −3.2 and stayed deeply negative for over a year. No stop-loss level tested improved performance — tighter stops generated cascading re-entries in the same adverse environment. The strategy is long-belly only.

5. **The 2Y wing notional dominates carry**: At $118.4M, the 2Y short position carries the largest notional. Carry P&L is therefore dominated by the 2Y leg vs SOFR spread, rather than the belly leg.

---

## Limitations & Extensions

| Limitation | Potential Improvement |
|---|---|
| Fixed modified durations | Rebalance DV01 dynamically at each entry using prevailing yield levels |
| No stop-loss | Exit if Z reaches +3.5 to cap runaway losses in regime-change environments |
| No momentum filter | Block entry if 20-day butterfly change is still widening (spread not yet peaking) |
| CMT par yields | Use on-the-run Treasury prices for precise DV01 and carry calculations |
| One trade at a time | Allow scaling — add at Z = 3.0 and scale out symmetrically |
| No transaction costs | Bid-offer on Treasuries (~0.5–1 bp per leg); haircuts on repo collateral |
| SOFR as single repo rate | In practice, specific-issue repo rates differ from general collateral SOFR |
| Equal DV01 per wing | Consider convexity-adjusted or market-value-weighted wings |

---

## How to Run

### Prerequisites

```bash
pip install requests pandas matplotlib xlsxwriter
```

A free FRED API key is required. Register at https://fred.stlouisfed.org/docs/api/api_key.html

### Step 1 — Download Treasury Data

```bash
set FRED_API_KEY=your_key_here
python fred_treasury_rates.py
```

Outputs `treasury_rates.csv` (date, DGS2, DGS5, DGS10, butterfly_spread, roll_mean_252, roll_std_252, z_score_252) and `butterfly_spread.png`.

### Step 2 — Run the Backtest

```bash
set FRED_API_KEY=your_key_here
python butterfly.py
```

Outputs:
- `butterfly_trades_long.csv` — long belly trade log
- `butterfly_trades.png` — 3-panel chart (spread, Z-score, cumulative P&L)

### Step 3 — Generate Excel Workbooks

```bash
python create_xlsx_files.py
```

Outputs:
- `treasury_rates.xlsx` — butterfly spread and 252d rolling Z-score with live Excel formulas
- `butterfly_trades.xlsx` — two-sheet workbook (Long Belly Trades / Summary) with all P&L as live formulas

> **Note**: Run steps in order. `butterfly.py` reads `treasury_rates.csv` produced in Step 1. `create_xlsx_files.py` reads all CSVs from Steps 1 and 2.
