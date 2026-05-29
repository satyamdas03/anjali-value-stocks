"""
AnjaliValueStocks scheduler.
Runs full fundamental refresh every 6 hours and lightweight price refresh every 30 minutes.
Logs to scheduler.log with rotation (keeps last 10 backups).
"""
import os
import sys
import subprocess
import logging
import logging.handlers
import time
from datetime import datetime

from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.interval import IntervalTrigger

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_PATH = os.path.join(SCRIPT_DIR, "scheduler.log")
PRICE_REFRESH_SCRIPT = os.path.join(SCRIPT_DIR, "price_refresh.py")

# ---------------------------------------------------------
# Logging setup (rotating, keeps 10 backups)
# ---------------------------------------------------------
def setup_logging():
    logger = logging.getLogger("anjali_scheduler")
    logger.setLevel(logging.INFO)
    # Prevent duplicate handlers if module is reloaded
    logger.handlers = []

    fmt = "%(asctime)s - %(levelname)s - %(message)s"
    formatter = logging.Formatter(fmt)

    file_handler = logging.handlers.RotatingFileHandler(
        LOG_PATH, maxBytes=2 * 1024 * 1024, backupCount=10
    )
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)

    return logger


logger = setup_logging()

# ---------------------------------------------------------
# Price refresh inline script (written to disk on startup)
# ---------------------------------------------------------
PRICE_REFRESH_SOURCE = '''\
"""
Lightweight price refresh for all tickers.
Updates returns, standard deviation, and beta only.
Outputs: Price_Refresh_SP500.csv, Price_Refresh_NSE.csv
"""
import os
import sys
import time
from datetime import datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

import yfinance as yf
import pandas as pd
import numpy as np

try:
    from collect_data import fetch_sp500_tickers
except Exception as e:
    print(f"Warning: could not import collect_data.fetch_sp500_tickers: {e}")
    def fetch_sp500_tickers():
        return []

try:
    from collect_indian_data import fetch_nifty100_tickers
except Exception as e:
    print(f"Warning: could not import collect_indian_data.fetch_nifty100_tickers: {e}")
    def fetch_nifty100_tickers():
        return []


def safe_round(val, decimals=2):
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return None
    return round(float(val), decimals)


def refresh_group(tickers, bench_ticker, out_csv):
    """Fetch 3M return, QtrStd, YrStd, Qtr Beta, Yr Beta for a ticker list."""
    bench = yf.Ticker(bench_ticker)
    bench_hist_3mo = bench.history(period="3mo")
    bench_hist_1y = bench.history(period="1y")
    bench_qtr_ret = bench_hist_3mo["Close"].pct_change().dropna()
    bench_yr_ret = bench_hist_1y["Close"].pct_change().dropna()

    rows = []
    total = len(tickers)
    for i, ticker in enumerate(tickers):
        row = {
            "Ticker": ticker,
            "Timestamp": datetime.utcnow().isoformat(),
        }
        try:
            stock = yf.Ticker(ticker)
            hist_3mo = stock.history(period="3mo")
            hist_1y = stock.history(period="1y")

            if len(hist_1y) > 1:
                price_now = hist_1y["Close"].iloc[-1]
                price_3m = hist_3mo["Close"].iloc[0] if len(hist_3mo) > 0 else None
                row["3M Return"] = safe_round(
                    ((price_now / price_3m) - 1) * 100
                ) if price_3m else None

            if len(hist_1y) > 20:
                daily_returns = hist_1y["Close"].pct_change().dropna()
                qtr_daily = daily_returns.tail(63)
                row["QtrStd"] = safe_round(qtr_daily.std() * np.sqrt(252) * 100)
                row["YrStd"] = safe_round(daily_returns.std() * np.sqrt(252) * 100)

            if len(hist_3mo) > 10 and len(bench_qtr_ret) > 10:
                stock_qtr_ret = hist_3mo["Close"].pct_change().dropna()
                common = stock_qtr_ret.index.intersection(bench_qtr_ret.index)
                if len(common) > 10:
                    cov = np.cov(stock_qtr_ret.loc[common], bench_qtr_ret.loc[common])[0][1]
                    var = np.var(bench_qtr_ret.loc[common])
                    row["Qtr Beta"] = safe_round(cov / var, 4) if var > 0 else None

            if len(hist_1y) > 20 and len(bench_yr_ret) > 20:
                stock_yr_ret = hist_1y["Close"].pct_change().dropna()
                common = stock_yr_ret.index.intersection(bench_yr_ret.index)
                if len(common) > 20:
                    cov = np.cov(stock_yr_ret.loc[common], bench_yr_ret.loc[common])[0][1]
                    var = np.var(bench_yr_ret.loc[common])
                    row["Yr Beta"] = safe_round(cov / var, 4) if var > 0 else None

        except Exception as e:
            row["Error"] = str(e)

        rows.append(row)
        if (i + 1) % 50 == 0:
            print(f"  [{i+1}/{total}] refreshed")
        time.sleep(0.1)

    df = pd.DataFrame(rows)
    df.to_csv(out_csv, index=False)
    print(f"Saved {out_csv} ({len(df)} rows)")


if __name__ == "__main__":
    print("Lightweight price refresh starting...")
    sp500 = fetch_sp500_tickers()
    if sp500:
        refresh_group(sp500, "SPY", os.path.join(SCRIPT_DIR, "Price_Refresh_SP500.csv"))
    else:
        print("No S&P 500 tickers to refresh.")

    indian = fetch_nifty100_tickers()
    if indian:
        indian_ns = [f"{t}.NS" for t in indian]
        refresh_group(indian_ns, "^NSEI", os.path.join(SCRIPT_DIR, "Price_Refresh_NSE.csv"))
    else:
        print("No Indian tickers to refresh.")

    print("Lightweight price refresh complete.")
'''

# ---------------------------------------------------------
# Helpers
# ---------------------------------------------------------
def write_price_refresh_script():
    """Ensure the lightweight price-refresh script exists on disk."""
    with open(PRICE_REFRESH_SCRIPT, "w", encoding="utf-8") as f:
        f.write(PRICE_REFRESH_SOURCE)
    logger.info(f"Ensured {PRICE_REFRESH_SCRIPT} exists")


def run_script(script_name, cwd=None):
    """Run a Python script via subprocess.run and return True on success."""
    path = os.path.join(SCRIPT_DIR, script_name)
    logger.info(f"Starting {script_name}")
    try:
        result = subprocess.run(
            [sys.executable, path],
            cwd=cwd or SCRIPT_DIR,
            capture_output=True,
            text=True,
            check=False,
        )
        if result.returncode != 0:
            logger.error(f"{script_name} failed (exit {result.returncode})")
            if result.stdout:
                logger.error(f"STDOUT: {result.stdout[:2000]}")
            if result.stderr:
                logger.error(f"STDERR: {result.stderr[:2000]}")
            return False

        logger.info(f"{script_name} completed successfully")
        # Log first 30 lines of stdout for visibility
        if result.stdout:
            for line in result.stdout.strip().splitlines()[:30]:
                logger.info(f"  {line}")
        return True
    except Exception as e:
        logger.exception(f"Exception running {script_name}: {e}")
        return False


def full_refresh():
    """Run the full pipeline: US data -> Indian data -> Excel build."""
    logger.info("=== Full refresh started ===")
    ok = run_script("collect_data.py")
    if not ok:
        logger.error("Full refresh chain aborted after collect_data.py")
        return

    ok = run_script("collect_indian_data.py")
    if not ok:
        logger.error("Full refresh chain aborted after collect_indian_data.py")
        return

    ok = run_script("build_excel.py")
    if not ok:
        logger.error("Full refresh chain aborted after build_excel.py")
        return

    logger.info("=== Full refresh completed successfully ===")


def price_refresh():
    """Run the lightweight price-refresh script."""
    logger.info("=== Price refresh started ===")
    ok = run_script("price_refresh.py")
    if not ok:
        logger.error("Price refresh failed, will retry next cycle")
    else:
        logger.info("=== Price refresh completed ===")


# ---------------------------------------------------------
# Main
# ---------------------------------------------------------
def main():
    print("Scheduler started. Full refresh every 6h, price refresh every 30min.")
    logger.info("Scheduler initializing")

    # Ensure the lightweight refresh script is on disk so subprocess can run it
    write_price_refresh_script()

    # On startup, run one full refresh immediately
    full_refresh()

    scheduler = BackgroundScheduler()
    scheduler.add_job(
        full_refresh,
        trigger=IntervalTrigger(hours=6),
        id="full_refresh",
        replace_existing=True,
    )
    scheduler.add_job(
        price_refresh,
        trigger=IntervalTrigger(minutes=30),
        id="price_refresh",
        replace_existing=True,
    )
    scheduler.start()

    logger.info("Scheduler running. Jobs: full_refresh every 6h, price_refresh every 30min.")
    try:
        while True:
            time.sleep(60)
    except (KeyboardInterrupt, SystemExit):
        logger.info("Scheduler shutting down")
        scheduler.shutdown()
        sys.exit(0)


if __name__ == "__main__":
    main()
