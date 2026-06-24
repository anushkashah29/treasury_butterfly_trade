"""
Downloads 2Y, 5Y, 10Y Constant Maturity Treasury rates from FRED API.

Output columns in treasury_rates.csv:
  date              : trading date (YYYY-MM-DD)
  DGS2              : 2-Year CMT yield (%)
  DGS5              : 5-Year CMT yield (%)
  DGS10             : 10-Year CMT yield (%)
  butterfly_spread  : 2*DGS5 - DGS2 - DGS10  (%)
  roll_mean_252     : 252-day rolling mean of butterfly_spread
  roll_std_252      : 252-day rolling sample std of butterfly_spread  (ddof=1)
  z_score_252       : (butterfly_spread - roll_mean_252) / roll_std_252
"""

import os
import requests
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import date

FRED_API_KEY = os.environ.get("FRED_API_KEY", "")
FRED_BASE_URL = "https://api.stlouisfed.org/fred/series/observations"

SERIES = ["DGS2", "DGS5", "DGS10"]
START_DATE = "1976-06-01"
OUTPUT_FILE = "treasury_rates.csv"
CHART_FILE = "butterfly_spread.png"


def fetch_series(series_id: str, api_key: str, start: str, end: str) -> pd.Series:
    params = {
        "series_id": series_id,
        "observation_start": start,
        "observation_end": end,
        "api_key": api_key,
        "file_type": "json",
    }
    response = requests.get(FRED_BASE_URL, params=params, timeout=30)
    response.raise_for_status()
    observations = response.json()["observations"]
    data = {
        obs["date"]: float(obs["value"]) if obs["value"] != "." else None
        for obs in observations
    }
    return pd.Series(data, name=series_id, dtype="float64")


def plot_butterfly(df: pd.DataFrame):
    fig, ax = plt.subplots(figsize=(14, 6))

    ax.plot(df["date"], df["butterfly_spread"], color="#1f77b4", linewidth=0.8, alpha=0.9)
    ax.axhline(0, color="black", linewidth=0.8, linestyle="--", alpha=0.5)
    ax.fill_between(df["date"], df["butterfly_spread"], 0,
                    where=df["butterfly_spread"] >= 0, alpha=0.15, color="green", label="Positive")
    ax.fill_between(df["date"], df["butterfly_spread"], 0,
                    where=df["butterfly_spread"] < 0, alpha=0.15, color="red", label="Negative")

    ax.xaxis.set_major_locator(mdates.YearLocator(5))
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    plt.xticks(rotation=45)

    ax.set_title("Treasury Butterfly Spread (2×DGS5 − DGS2 − DGS10)", fontsize=14, fontweight="bold")
    ax.set_xlabel("Date")
    ax.set_ylabel("Spread (bps / %)")
    ax.legend(loc="upper right")
    ax.grid(True, alpha=0.3)

    plt.tight_layout()
    plt.savefig(CHART_FILE, dpi=150)
    plt.close()
    print(f"Chart saved to {CHART_FILE}")


def main():
    api_key = FRED_API_KEY
    if not api_key:
        api_key = input("Enter your FRED API key: ").strip()
    if not api_key:
        raise ValueError("FRED API key is required.")

    end_date = date.today().isoformat()
    print(f"Fetching treasury rates from {START_DATE} to {end_date}...")

    series_data = {}
    for series_id in SERIES:
        print(f"  Fetching {series_id}...")
        series_data[series_id] = fetch_series(series_id, api_key, START_DATE, end_date)

    df = pd.DataFrame(series_data)
    df.index.name = "date"
    df.reset_index(inplace=True)
    df["date"] = pd.to_datetime(df["date"])

    # Drop holiday/weekend rows where any rate is missing
    df.dropna(subset=["DGS2", "DGS5", "DGS10"], inplace=True)
    df.reset_index(drop=True, inplace=True)

    # Butterfly spread: 2*DGS5 - DGS2 - DGS10
    df["butterfly_spread"] = 2 * df["DGS5"] - df["DGS2"] - df["DGS10"]

    # 252-day rolling Z-score (one trading year)
    # Formula: z = (butterfly_spread - roll_mean_252) / roll_std_252
    # roll_std_252 uses ddof=1 (sample std, pandas default)
    # Requires exactly 252 observations -- first 251 rows are NaN
    df["roll_mean_252"] = df["butterfly_spread"].rolling(252).mean().round(4)
    df["roll_std_252"]  = df["butterfly_spread"].rolling(252).std().round(4)
    df["z_score_252"]   = (
        (df["butterfly_spread"] - df["roll_mean_252"]) / df["roll_std_252"]
    ).round(3)

    df["date"] = df["date"].dt.strftime("%Y-%m-%d")
    df.to_csv(OUTPUT_FILE, index=False)
    print(f"Saved {len(df)} rows to {OUTPUT_FILE}")
    print(df.tail())

    df["date"] = pd.to_datetime(df["date"])
    plot_butterfly(df)


if __name__ == "__main__":
    main()
