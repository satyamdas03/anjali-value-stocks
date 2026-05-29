"""
Collect Indian NIFTY 100 stock data from yfinance.
Matches the exact pattern of collect_data.py for US stocks.

Output: Indian_Stock_Data.csv — ~100 stocks x 33 columns with all metrics calculated.
"""
import yfinance as yf
import pandas as pd
import numpy as np
import warnings
import time
import os
import sys
from datetime import datetime

warnings.filterwarnings("ignore")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_CSV = os.path.join(SCRIPT_DIR, "Indian_Stock_Data.csv")
CHECKPOINT_CSV = os.path.join(SCRIPT_DIR, "Indian_Stock_Data_checkpoint.csv")
PROGRESS_FILE = os.path.join(SCRIPT_DIR, "collect_india_progress.txt")

REVENUE_ROWS = ["TotalRevenue", "Total Revenue", "Net Revenue"]
INCOME_ROWS = [
    "NetIncome", "Net Income", "Net Income Common Stockholders",
    "Net Income Continuous Operations",
]

COLUMNS = [
    "Index Name", "Ticker", "Sector", "Sub Sector",
    "Sales YoY Growth", "NetProfit YoY Growth",
    "Sales TTM 1Yr Growth", "NetProfit TTM 1Yr Growth",
    "QoQ Sales Growth", "QoQ Profit Growth",
    "3M Return", "6M Return", "1Yr Return", "2Yr Return",
    "PE Ratio", "Future PE", "TTM PEG", "Future PEG",
    "PB Ratio", "EV/Sales", "EV/EBITDA",
    "Market Cap (Billions)", "Revenue (Billions)", "TTM Revenue (Billions)",
    "QtrStd", "YrStd", "Qtr Beta", "Yr Beta",
    "_Loss_Profit_YoY", "_Loss_Profit_TTM", "_Loss_Profit_QoQ",
    "DII Quarter", "DII 1Yr", "FII Quarter", "FII 1Yr",
]


def fetch_nifty100_tickers():
    """Return hardcoded NSE ticker list extracted from sample data.
    Wikipedia has no NIFTY 100 constituents table."""
    print("Using hardcoded NSE ticker list (250 stocks from sample data)...")
    return _fallback_tickers()


def _fallback_tickers():
    """Hardcoded 250 NSE stocks extracted from sample data (LM 250)."""
    tickers = [
        "PAYTM","UPL","ABFRL","FACT","GMRAIRPORT","IDEA","BHARTIHEXA","MRPL","HINDPETRO","TATACHEM",
        "PRESTIGE","IOC","FLUOROCHEM","GODREJIND","BPCL","POONAWALLA","JSWSTEEL","LTF","ADANIPOWER","SRF",
        "COROMANDEL","MFSL","ABCAPITAL","GRASIM","TATACOMM","BAYERCROP","NHPC","APLAPOLLO","DALBHARAT","IDFCFIRSTB",
        "AXISBANK","ONGC","AIAENG","SJVN","TATATECH","JSL","TECHM","RVNL","SYNGENE","PGHH",
        "ASIANPAINT","GLAND","ADANIGREEN","SHREECEM","ASTRAL","AMBUJACEM","SBICARD","EXIDEIND","BRITANNIA","ICICIPRULI",
        "APOLLOTYRE","INDUSINDBK","GUJGASLTD","TITAN","SUNTV","GRINDWELL","RELIANCE","MPHASIS","SCHAEFFLER","HINDUNILVR",
        "DEEPAKNTR","JUBLFOOD","WIPRO","DABUR","ITC","IGL","POWERGRID","KPRMILL","BHEL","TORNTPOWER",
        "LTTS","TATAELXSI","OIL","HINDZINC","ZFCVINDIA","DRREDDY","LTIM","NLCINDIA","SUPREMEIND","LINDEINDIA",
        "MAXHEALTH","CARBORUNIV","UNITDSPR","BAJAJ-AUTO","TCS","COFORGE","CONCOR","IRFC","SUNDARMFIN","INFY",
        "COALINDIA","BANKBARODA","SUNDRMFAST","HCLTECH","NESTLEIND","BANDHANBNK","GICRE","M&MFIN","M&M","TIMKEN",
        "CRISIL","LT","BERGEPAINT","IRCTC","EMAMILTD","TATAPOWER","POLYCAB","LICI","3MINDIA","KOTAKBANK",
        "FEDERALBNK","SBIN","MARICO","SAIL","TATAINVEST","PETRONET","PATANJALI","ULTRACEMCO","BAJAJFINSV","HDFCLIFE",
        "CANBK","MRF","DMART","NTPC","TIINDIA","SKFINDIA","RECLTD","JINDALSTEL","AUBANK","BAJFINANCE",
        "LICHSGFIN","MUTHOOTFIN","TATACONSUM","ICICIBANK","INDIGO","PFC","KEI","OBEROIRLTY","SBILIFE","PIIND",
        "ABBOTINDIA","IPCALAB","STARHEALTH","HONAUT","EICHERMOT","MEDANTA","HAVELLS","PHOENIXLTD","KALYANKJIL","UNOMINDA",
        "THERMAX","ASHOKLEY","PAGEIND","METROBRAND","SONACOMS","SHRIRAMFIN","MARUTI","PERSISTENT","HEROMOTOCO","ATGL",
        "IRB","JSWENERGY","GLAXO","BHARTIARTL","COLPAL","ADANIPORTS","SUNPHARMA","SOLARINDS","NMDC","JSWINFRA",
        "MSUMI","GODREJCP","IOB","ICICIGI","HDFCBANK","VBL","CHOLAFIN","UNIONBANK","CIPLA","INDHOTEL",
        "TVSMOTOR","LLOYDSME","TORNTPHARM","BDL","PIDILITIND","BOSCHLTD","VOLTAS","FORTIS","ZYDUSLIFE","AJANTPHARM",
        "ESCORTS","HDFCAMC","BAJAJHLDNG","CUMMINSIND","MANKIND","OFSS","BEL","ENDURANCE","HUDCO","DIVISLAB",
        "SIEMENS","ACC","ABB","CGPOWER","IREDA","LODHA","HAL","IDBI","MAHABANK","INDIANB",
        "ALKEM","NIACL","BANKINDIA","BHARATFORG","NAM-INDIA","KPITTECH","AUROPHARMA","HINDALCO","BALKRISIND","JIOFIN",
        "DLF","JKCEMENT","BIOCON","APOLLOHOSP","ADANIENSOL","NAUKRI","DIXON","TATASTEEL","MOTHERSON","MAZDOCK",
        "UBL","LUPIN","INDUSTOWER","DELHIVERY","GODREJPROP","ADANIENT","YESBANK","NYKAA","TRENT","GAIL",
        "COCHINSHIP","BSE","VEDL","TATAMOTORS","POWERINDIA","PNB","SUZLON","POLICYBZR","ETERNAL","AWL",
    ]
    return list(dict.fromkeys(tickers))


def load_progress():
    """Load last completed ticker index from progress file."""
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE) as f:
            return int(f.read().strip())
    return 0


def save_progress(idx):
    with open(PROGRESS_FILE, "w") as f:
        f.write(str(idx))


def try_financials_row(fin, row_names, col_name):
    """Try multiple row name variants to extract a value from financials."""
    row = next((r for r in row_names if r in fin.index), None)
    if row:
        val = fin.loc[row, col_name]
        if pd.notna(val):
            return val
    return None


def safe_round(val, decimals=2):
    """Round a value, returning None if NaN."""
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return None
    return round(float(val), decimals)


def collect_stock_data(tickers):
    """Main data collection for all tickers, with checkpoint resume."""
    start_idx = load_progress()
    print(f"Starting from ticker index {start_idx} (already processed)")

    # Pre-fetch NSE NIFTY 50 (^NSEI) data once for all beta calculations
    print("Fetching ^NSEI benchmark data...")
    nsei = yf.Ticker("^NSEI")
    nsei_hist_3mo = nsei.history(period="3mo")
    nsei_hist_1y = nsei.history(period="1y")
    nsei_qtr_returns = nsei_hist_3mo["Close"].pct_change().dropna()
    nsei_yr_returns = nsei_hist_1y["Close"].pct_change().dropna()

    # Load existing checkpoint if resuming
    if start_idx > 0 and os.path.exists(CHECKPOINT_CSV):
        df = pd.read_csv(CHECKPOINT_CSV)
        print(f"  Resumed from checkpoint with {len(df)} rows")
    else:
        df = pd.DataFrame(columns=COLUMNS)

    total = len(tickers)
    processed_since_checkpoint = 0

    for i in range(start_idx, total):
        ticker_raw = tickers[i]
        ticker = f"{ticker_raw}.NS"
        row = {col: None for col in COLUMNS}
        row["Ticker"] = ticker_raw
        row["Index Name"] = "LM 250"

        try:
            stock = yf.Ticker(ticker)
            info = stock.info or {}

            row["Sector"] = info.get("sector", "")
            row["Sub Sector"] = info.get("industry", "")

            # --- Price History & Returns ---
            hist_3mo = stock.history(period="3mo")
            hist_6mo = stock.history(period="6mo")
            hist_1y = stock.history(period="1y")
            hist_2y = stock.history(period="2y")

            if len(hist_1y) > 1:
                price_now = hist_1y["Close"].iloc[-1]
                price_3m = hist_3mo["Close"].iloc[0] if len(hist_3mo) > 0 else None
                price_6m = hist_6mo["Close"].iloc[0] if len(hist_6mo) > 0 else None
                price_1y = hist_1y["Close"].iloc[0]
                price_2y = hist_2y["Close"].iloc[0] if len(hist_2y) > 0 else None

                row["3M Return"] = safe_round(((price_now / price_3m) - 1) * 100) if price_3m else None
                row["6M Return"] = safe_round(((price_now / price_6m) - 1) * 100) if price_6m else None
                row["1Yr Return"] = safe_round(((price_now / price_1y) - 1) * 100) if price_1y else None
                row["2Yr Return"] = safe_round(((price_now / price_2y) - 1) * 100) if price_2y else None
            else:
                row["3M Return"] = row["6M Return"] = row["1Yr Return"] = row["2Yr Return"] = None

            # --- Risk Metrics (Std Dev) ---
            if len(hist_1y) > 20:
                daily_returns = hist_1y["Close"].pct_change().dropna()
                qtr_daily = daily_returns.tail(63)
                row["QtrStd"] = safe_round(qtr_daily.std() * np.sqrt(252) * 100)
                row["YrStd"] = safe_round(daily_returns.std() * np.sqrt(252) * 100)
            else:
                row["QtrStd"] = row["YrStd"] = None

            # --- Beta Calculation against ^NSEI ---
            if len(hist_3mo) > 10 and len(nsei_qtr_returns) > 10:
                stock_qtr_ret = hist_3mo["Close"].pct_change().dropna()
                common = stock_qtr_ret.index.intersection(nsei_qtr_returns.index)
                if len(common) > 10:
                    cov = np.cov(stock_qtr_ret.loc[common], nsei_qtr_returns.loc[common])[0][1]
                    var = np.var(nsei_qtr_returns.loc[common])
                    row["Qtr Beta"] = safe_round(cov / var, 4) if var > 0 else None

            if len(hist_1y) > 20 and len(nsei_yr_returns) > 20:
                stock_yr_ret = hist_1y["Close"].pct_change().dropna()
                common = stock_yr_ret.index.intersection(nsei_yr_returns.index)
                if len(common) > 20:
                    cov = np.cov(stock_yr_ret.loc[common], nsei_yr_returns.loc[common])[0][1]
                    var = np.var(nsei_yr_returns.loc[common])
                    row["Yr Beta"] = safe_round(cov / var, 4) if var > 0 else None

            # --- Valuation from .info ---
            pe = info.get("trailingPE")
            pb = info.get("priceToBook")
            row["PE Ratio"] = safe_round(pe) if pe else None
            row["PB Ratio"] = safe_round(pb) if pb and pb > 0 else None
            row["EV/Sales"] = safe_round(info.get("enterpriseToRevenue"))
            ev_ebitda = info.get("enterpriseToEbitda")
            if ev_ebitda is not None and ev_ebitda > 0:
                row["EV/EBITDA"] = safe_round(ev_ebitda)
            else:
                row["EV/EBITDA"] = None

            # --- Size Metrics (.info based) ---
            mc = info.get("marketCap")
            row["Market Cap (Billions)"] = safe_round(mc / 1e9) if mc else None
            rev = info.get("totalRevenue")
            row["Revenue (Billions)"] = safe_round(rev / 1e9) if rev else None

            # --- YoY Growth from .info ---
            rev_g = info.get("revenueGrowth")
            earn_g = info.get("earningsGrowth")
            row["Sales YoY Growth"] = safe_round(rev_g * 100) if rev_g is not None else None
            row["NetProfit YoY Growth"] = safe_round(earn_g * 100) if earn_g is not None else None

            # --- TTM Growth from Annual Financials ---
            fin = stock.financials
            if not fin.empty and fin.shape[1] >= 2:
                latest_col = fin.columns[0]
                prior_col = fin.columns[1]

                rev_curr = try_financials_row(fin, REVENUE_ROWS, latest_col)
                rev_prev = try_financials_row(fin, REVENUE_ROWS, prior_col)
                if rev_curr and rev_prev and rev_prev != 0:
                    row["Sales TTM 1Yr Growth"] = safe_round((rev_curr / rev_prev - 1) * 100, 1)

                ni_curr = try_financials_row(fin, INCOME_ROWS, latest_col)
                ni_prev = try_financials_row(fin, INCOME_ROWS, prior_col)
                if ni_curr and ni_prev and ni_prev != 0:
                    row["NetProfit TTM 1Yr Growth"] = safe_round((ni_curr / ni_prev - 1) * 100, 1)
                # TTM Revenue from annual financials
                if rev_curr:
                    row["TTM Revenue (Billions)"] = safe_round(rev_curr / 1e9)

            # --- QoQ Growth from Quarterly Financials ---
            qfin = stock.quarterly_financials
            if not qfin.empty and qfin.shape[1] >= 2:
                qlatest = qfin.columns[0]
                qprior = qfin.columns[1]

                qrev_curr = try_financials_row(qfin, REVENUE_ROWS, qlatest)
                qrev_prev = try_financials_row(qfin, REVENUE_ROWS, qprior)
                if qrev_curr and qrev_prev and qrev_prev != 0:
                    row["QoQ Sales Growth"] = safe_round((qrev_curr / qrev_prev - 1) * 100, 1)
                elif row["QoQ Sales Growth"] is None:
                    rqg = info.get("revenueQuarterlyGrowth")
                    row["QoQ Sales Growth"] = safe_round(rqg * 100) if rqg is not None else None

                qni_curr = try_financials_row(qfin, INCOME_ROWS, qlatest)
                qni_prev = try_financials_row(qfin, INCOME_ROWS, qprior)
                if qni_curr and qni_prev and qni_prev != 0:
                    row["QoQ Profit Growth"] = safe_round((qni_curr / qni_prev - 1) * 100, 1)
                elif row["QoQ Profit Growth"] is None:
                    eqg = info.get("earningsQuarterlyGrowth")
                    row["QoQ Profit Growth"] = safe_round(eqg * 100) if eqg is not None else None

            # --- Cap extreme QoQ Profit Growth ---
            if row["QoQ Profit Growth"] is not None:
                if row["QoQ Profit Growth"] > 500:
                    row["QoQ Profit Growth"] = 500.0
                elif row["QoQ Profit Growth"] < -500:
                    row["QoQ Profit Growth"] = -500.0

            # --- Future PE = Current PE * (1 + TTM Profit Growth / 100) ---
            if row["PE Ratio"] is not None and row["NetProfit TTM 1Yr Growth"] is not None:
                row["Future PE"] = safe_round(row["PE Ratio"] * (1 + row["NetProfit TTM 1Yr Growth"] / 100))

            # --- PEG = PE / NetProfit_TTM_Growth ---
            if row["PE Ratio"] is not None and row["NetProfit TTM 1Yr Growth"] is not None and row["NetProfit TTM 1Yr Growth"] != 0:
                row["TTM PEG"] = safe_round(row["PE Ratio"] / row["NetProfit TTM 1Yr Growth"])

            if row["Future PE"] is not None and row["NetProfit TTM 1Yr Growth"] is not None and row["NetProfit TTM 1Yr Growth"] != 0:
                row["Future PEG"] = safe_round(row["Future PE"] / row["NetProfit TTM 1Yr Growth"])

            # --- Loss Flags (profit columns only) ---
            row["_Loss_Profit_YoY"] = (row["NetProfit YoY Growth"] is not None and row["NetProfit YoY Growth"] < 0)
            row["_Loss_Profit_TTM"] = (row["NetProfit TTM 1Yr Growth"] is not None and row["NetProfit TTM 1Yr Growth"] < 0)
            row["_Loss_Profit_QoQ"] = (row["QoQ Profit Growth"] is not None and row["QoQ Profit Growth"] < 0)

            # --- India-specific columns (empty for now) ---
            row["DII Quarter"] = None
            row["DII 1Yr"] = None
            row["FII Quarter"] = None
            row["FII 1Yr"] = None

        except Exception as e:
            print(f"  [{i+1}/{total}] {ticker}: ERROR — {e}")

        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
        processed_since_checkpoint += 1

        if (i + 1) % 10 == 0 or i == total - 1:
            pct = (i + 1) / total * 100
            print(f"  [{i+1}/{total}] {ticker}: PE={row.get('PE Ratio')}, "
                  f"1Yr={row.get('1Yr Return')}, Beta={row.get('Yr Beta')}  ({pct:.0f}%)")

        # Checkpoint every 50 stocks
        if processed_since_checkpoint >= 50 or i == total - 1:
            df.to_csv(CHECKPOINT_CSV, index=False)
            save_progress(i + 1)
            print(f"  -- Checkpoint saved at {i+1}/{total} ({datetime.now().strftime('%H:%M:%S')}) --")
            processed_since_checkpoint = 0

        time.sleep(0.3)

    # Final save
    df.to_csv(OUTPUT_CSV, index=False)
    # Clean up temp files
    for f in [CHECKPOINT_CSV, PROGRESS_FILE]:
        if os.path.exists(f):
            os.remove(f)

    print(f"\nDone! {len(df)} stocks saved to {OUTPUT_CSV}")
    _print_summary(df)
    return df


def _print_summary(df):
    """Print data completeness summary."""
    print("\n=== Column Fill Rates ===")
    for col in COLUMNS:
        if col.startswith("_"):
            continue
        n = df[col].notna().sum()
        print(f"  {col}: {n}/{len(df)} ({100*n//len(df)}%)")

    # Top/bottom highlights
    print("\n=== Top 5 by 1Yr Return ===")
    top5 = df.nlargest(5, "1Yr Return")[["Ticker", "1Yr Return"]]
    for _, r in top5.iterrows():
        print(f"  {r['Ticker']}: +{r['1Yr Return']}%")

    print("\n=== Lowest 5 PE ===")
    bot5 = df.nsmallest(5, "PE Ratio")[["Ticker", "PE Ratio"]]
    for _, r in bot5.iterrows():
        print(f"  {r['Ticker']}: {r['PE Ratio']}")


if __name__ == "__main__":
    tickers = fetch_nifty100_tickers()
    # Deduplicate
    tickers = list(dict.fromkeys(tickers))
    print(f"Collecting data for {len(tickers)} stocks...\n")
    collect_stock_data(tickers)
