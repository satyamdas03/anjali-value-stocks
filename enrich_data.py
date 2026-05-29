"""
Enrich US Stock data with TTM, QoQ growth, PEG calculations.
Fixed version using column-based indexing.
"""
import yfinance as yf
import pandas as pd
import numpy as np
import warnings
import time
import os

warnings.filterwarnings('ignore')

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_PATH = os.path.join(SCRIPT_DIR, 'US_Stock_Data_Raw.csv')
OUTPUT_CSV = os.path.join(SCRIPT_DIR, 'US_Stock_Data_Enriched.csv')

df = pd.read_csv(CSV_PATH)
print(f"Loaded {len(df)} stocks")

REVENUE_ROWS = ['TotalRevenue', 'Total Revenue', 'Net Revenue']
INCOME_ROWS = ['NetIncome', 'Net Income', 'Net Income Common Stockholders', 'Net Income Continuous Operations']

for idx, row in df.iterrows():
    ticker = row['Ticker']
    if not ticker or pd.isna(ticker):
        continue

    try:
        stock = yf.Ticker(ticker)

        # --- TTM Growth (annual financials) ---
        fin = stock.financials
        if not fin.empty and fin.shape[1] >= 2:
            cols = fin.columns
            latest = cols[0]
            prior = cols[1]

            rev_row = next((r for r in REVENUE_ROWS if r in fin.index), None)
            ni_row = next((r for r in INCOME_ROWS if r in fin.index), None)

            if rev_row:
                rev_curr = fin.loc[rev_row, latest]
                rev_prev = fin.loc[rev_row, prior]
                if pd.notna(rev_curr) and pd.notna(rev_prev) and rev_prev != 0:
                    df.at[idx, 'Sales TTM 1Yr Growth'] = round((rev_curr / rev_prev - 1) * 100, 1)

            if ni_row:
                ni_curr = fin.loc[ni_row, latest]
                ni_prev = fin.loc[ni_row, prior]
                if pd.notna(ni_curr) and pd.notna(ni_prev) and ni_prev != 0:
                    df.at[idx, 'NetProfit TTM 1Yr Growth'] = round((ni_curr / ni_prev - 1) * 100, 1)

        # --- QoQ Growth (quarterly financials) ---
        qfin = stock.quarterly_financials
        if not qfin.empty and qfin.shape[1] >= 2:
            qcols = qfin.columns
            qlatest = qcols[0]
            qprior = qcols[1]

            rev_row = next((r for r in REVENUE_ROWS if r in qfin.index), None)
            ni_row = next((r for r in INCOME_ROWS if r in qfin.index), None)

            if rev_row:
                qrev_curr = qfin.loc[rev_row, qlatest]
                qrev_prev = qfin.loc[rev_row, qprior]
                if pd.notna(qrev_curr) and pd.notna(qrev_prev) and qrev_prev != 0:
                    df.at[idx, 'QoQ Sales Growth'] = round((qrev_curr / qrev_prev - 1) * 100, 1)

            if ni_row:
                qni_curr = qfin.loc[ni_row, qlatest]
                qni_prev = qfin.loc[ni_row, qprior]
                if pd.notna(qni_curr) and pd.notna(qni_prev) and qni_prev != 0:
                    df.at[idx, 'QoQ Profit Growth'] = round((qni_curr / qni_prev - 1) * 100, 1)

        # --- PEG Calculations ---
        pe = row.get('PE Ratio')
        fwd_pe = row.get('TtmFuturePE')
        ttm_profit_g = df.at[idx, 'NetProfit TTM 1Yr Growth']
        earnings_g_pct = row.get('NetProfit YoY Growth')  # Already in %

        # Use YoY earnings growth for PEG if TTM not available
        profit_for_peg = ttm_profit_g if pd.notna(ttm_profit_g) else earnings_g_pct

        if pd.notna(pe) and pd.notna(profit_for_peg) and profit_for_peg != 0:
            df.at[idx, 'TTM PEG'] = round(abs(pe / profit_for_peg), 2)
            # For negative growth, PEG is negative
            if profit_for_peg < 0:
                df.at[idx, 'TTM PEG'] = round(pe / profit_for_peg, 2)

        if pd.notna(fwd_pe) and pd.notna(profit_for_peg) and profit_for_peg != 0:
            df.at[idx, 'Ttm FuturePEG'] = round(abs(fwd_pe / profit_for_peg), 2)
            if profit_for_peg < 0:
                df.at[idx, 'Ttm FuturePEG'] = round(fwd_pe / profit_for_peg, 2)

        # Fill missing YoY growth from .info if still missing
        info = stock.info
        if pd.isna(df.at[idx, 'QoQ Sales Growth']):
            rev_q_g = info.get('revenueQuarterlyGrowth')
            if rev_q_g is not None:
                df.at[idx, 'QoQ Sales Growth'] = round(rev_q_g * 100, 1)

        if pd.isna(df.at[idx, 'QoQ Profit Growth']):
            earn_q_g = info.get('earningsQuarterlyGrowth')
            if earn_q_g is not None:
                df.at[idx, 'QoQ Profit Growth'] = round(earn_q_g * 100, 1)

    except Exception as e:
        print(f"  Error enriching {ticker}: {e}")

    if (idx + 1) % 10 == 0:
        print(f"  Processed {idx + 1}/{len(df)} stocks")
        df.to_csv(OUTPUT_CSV, index=False)

    time.sleep(0.5)

df.to_csv(OUTPUT_CSV, index=False)
print(f"\nSaved enriched data to {OUTPUT_CSV}")

# Print completeness
for col in df.columns:
    non_null = df[col].notna().sum()
    print(f"  {col}: {non_null}/{len(df)}")