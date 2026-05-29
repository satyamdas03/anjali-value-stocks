"""
Scrape DII/FII shareholding data from Screener.in for Indian NSE stocks.
Populates DII Quarter, DII 1Yr, FII Quarter, FII 1Yr columns.

Uses Screener.in search API to resolve ticker → company URL.

Usage:
    python scrape_dii_fii.py
    python scrape_dii_fii.py --update   # Update Indian_Stock_Data.csv in-place

Screener.in provides quarterly shareholding pattern with:
- Promoters %
- FII %
- DII %
- Public %

Our columns are percentage point CHANGES:
- DII Quarter = DII%_latest - DII%_prev_quarter
- DII 1Yr = DII%_latest - DII%_4q_ago
- FII Quarter = FII%_latest - FII%_prev_quarter
- FII 1Yr = FII%_latest - FII%_4q_ago
"""
import os
import sys
import time
import json
import pandas as pd
import numpy as np
import re

try:
    import requests
    from bs4 import BeautifulSoup
except ImportError:
    print("Installing required packages: requests beautifulsoup4")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "requests", "beautifulsoup4"])
    import requests
    from bs4 import BeautifulSoup

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INDIAN_CSV = os.path.join(SCRIPT_DIR, "Indian_Stock_Data.csv")
OUTPUT_CSV = os.path.join(SCRIPT_DIR, "DII_FII_Data.csv")

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.5",
}

# Common ticker-to-Screener-slug mapping for popular stocks
# Screener uses lowercase-hyphenated company names as URL slugs
TICKER_SLUG_MAP = {
    "RELIANCE": "reliance-industries",
    "TCS": "tcs",
    "HDFCBANK": "hdfc-bank",
    "INFY": "infosys",
    "ICICIBANK": "icici-bank",
    "HINDUNILVR": "hindustan-unilever",
    "SBIN": "state-bank-of-india",
    "BHARTIARTL": "bharti-airtel",
    "ITC": "itc",
    "KOTAKBANK": "kotak-mahindra-bank",
    "LT": "larsen-toubro",
    "HCLTECH": "hcl-technologies",
    "ASIANPAINT": "asian-paints",
    "AXISBANK": "axis-bank",
    "BAJFINANCE": "bajaj-finance",
    "MARUTI": "maruti-suzuki-india",
    "SUNPHARMA": "sun-pharmaceutical-industries",
    "TITAN": "titan-company",
    "WIPRO": "wipro",
    "ULTRACEMCO": "ultratech-cement",
    "ADANIENT": "adani-enterprises",
    "ONGC": "oil-natural-gas-corporation",
    "NTPC": "ntpc",
    "POWERGRID": "power-grid-corporation-of-india",
    "TATAMOTORS": "tata-motors",
    "TATASTEEL": "tata-steel",
    "TATACONSUM": "tata-consumer-products",
    "BAJAJFINSV": "bajaj-finserv",
    "HINDALCO": "hindalco-industries",
    "COALINDIA": "coal-india",
    "JSWSTEEL": "jsw-steel",
    "GRASIM": "grasim-industries",
    "BPCL": "bharat-petroleum-corporation",
    "IOC": "indian-oil-corporation",
    "HDFC": "hdfc",
    "ADANIPORTS": "adani-ports-special-economic-zone",
    "TECHM": "tech-mahindra",
    "DIVISLAB": "divis-laboratories",
    "DRREDDY": "dr-reddys-laboratories",
    "CIPLA": "cipla",
    "APOLLOHOSP": "apollo-hospitals-enterprise",
    "EICHERMOT": "eicher-motors",
    "M_M": "mahindra-mahindra",
    "HEROMOTOCO": "hero-motocorp",
    "BAJAJ-AUTO": "bajaj-auto",
    "INDUSINDBK": "indusind-bank",
    "SBILIFE": "sbi-life-insurance-company",
    "HDFCLIFE": "hdfc-life-insurance-company",
    "TATAMTRDVR": "tata-motors",
    "PIDILITIND": "pidilite-industries",
    "DABUR": "dabur-india",
    "BRITANNIA": "britannia-industries",
    "NESTLEIND": "nestle-india",
    "GAIL": "gail-india",
    "PNB": "punjab-national-bank",
    "BANKBARODA": "bank-of-baroda",
    "CANBK": "canara-bank",
    "UNIONBANK": "union-bank-of-india",
    "IDFCFIRSTB": "idfc-first-bank",
    "FEDERALBNK": "federal-bank",
    "MUTHOOTFIN": "muthoot-finance",
    "SHREECEM": "shree-cement",
    "AMBUJACEM": "ambuja-cements",
    "ACC": "acc",
    "DLF": "dlf",
    "GODREJCP": "godrej-consumer-products",
    "HAVELLS": "havells-india",
    "VINATIORGA": "vinati-organics",
    "PAGEIND": "page-industries",
    "BOSCHLTD": "bosch",
    "COLPAL": "colgate-palmolive-india",
    "MARICO": "marico",
    "TORNTPHARM": "torrent-pharmaceuticals",
    "MCDOWELL-N": "united-spirits",
    "IDEA": "vodafone-idea",
    "NHPC": "nhpc",
    "RECLTD": "rec-ltd",
    "PNBHOUSING": "pnb-housing-finance",
    "LTF": "lt-investment-holdings",
    "ADANIGREEN": "adani-green-energy",
    "TATAPOWER": "tata-power-company",
    "YESBANK": "yes-bank",
    "SUZLON": "suzlon-energy",
    "ZOMATO": "zomato",
    "NYKAA": "nykaa",
    "PAYTM": "one97-communications",
}


def search_screener(ticker, session):
    """Resolve ticker to Screener.in URL. Screener uses uppercase ticker directly."""
    # Try uppercase ticker (Screener's primary URL format)
    url = f"https://www.screener.in/company/{ticker.upper()}/consolidated/"
    try:
        resp = session.get(url, timeout=15, allow_redirects=True)
        if resp.status_code == 200:
            return url, ticker.upper()
    except Exception:
        pass

    # Try without /consolidated/
    url = f"https://www.screener.in/company/{ticker.upper()}/"
    try:
        resp = session.get(url, timeout=15, allow_redirects=True)
        if resp.status_code == 200:
            return url, ticker.upper()
    except Exception:
        pass

    return None, None


def fetch_shareholding(ticker_ns, session=None):
    """Fetch DII/FII shareholding percentages from Screener.in for a single ticker.

    Returns dict with keys: DII_latest, DII_prev, DII_4q_ago, FII_latest, FII_prev, FII_4q_ago
    or None if data not found.
    """
    ticker = ticker_ns.replace(".NS", "")

    if session is None:
        session = requests.Session()
        session.headers.update(HEADERS)

    url, slug = search_screener(ticker, session)
    if not url:
        return None

    try:
        resp = session.get(url, timeout=15)
        if resp.status_code != 200:
            # Try without /consolidated/
            url = url.replace("/consolidated/", "/")
            resp = session.get(url, timeout=15)
            if resp.status_code != 200:
                return None

        soup = BeautifulSoup(resp.text, "html.parser")

        # Find the shareholding pattern section
        sh_section = soup.find("section", {"id": "shareholding"})
        if not sh_section:
            sh_section = soup.find("div", string=re.compile("Shareholding", re.IGNORECASE))
            if sh_section:
                sh_section = sh_section.find_parent("section") or sh_section.find_parent("div")

        if not sh_section:
            tables = soup.find_all("table")
            for table in tables:
                header_row = table.find("tr")
                if header_row:
                    header_text = header_row.get_text().lower()
                    if "fii" in header_text or "dii" in header_text or "promoter" in header_text:
                        sh_section = table
                        break

        if not sh_section:
            return None

        # Parse the shareholding table
        tables = sh_section.find_all("table") if sh_section.name != "table" else [sh_section]

        for table in tables:
            rows = table.find_all("tr")
            if len(rows) < 2:
                continue

            header_cells = rows[0].find_all(["th", "td"])
            quarters = [cell.get_text(strip=True) for cell in header_cells[1:]]

            dii_values = {}
            fii_values = {}

            for row in rows[1:]:
                cells = row.find_all(["th", "td"])
                if len(cells) < 2:
                    continue

                category = cells[0].get_text(strip=True).lower()
                values = []
                for cell in cells[1:]:
                    raw = cell.get_text(strip=True).replace(",", "").replace("%", "").strip()
                    try:
                        val = float(raw)
                        values.append(val)
                    except (ValueError, AttributeError):
                        values.append(None)

                # Screener uses "DIIs+" and "FIIs+" (with plus suffix)
                if "dii" in category and "fii" not in category:
                    for i, v in enumerate(values):
                        if v is not None and i < len(quarters):
                            dii_values[quarters[i]] = v
                elif "fii" in category or "foreign" in category:
                    for i, v in enumerate(values):
                        if v is not None and i < len(quarters):
                            fii_values[quarters[i]] = v

            if dii_values or fii_values:
                sorted_dii = sorted(dii_values.items(), key=lambda x: x[0], reverse=True)
                sorted_fii = sorted(fii_values.items(), key=lambda x: x[0], reverse=True)

                result = {}
                if len(sorted_dii) >= 1:
                    result["DII_latest"] = sorted_dii[0][1]
                if len(sorted_dii) >= 2:
                    result["DII_prev"] = sorted_dii[1][1]
                if len(sorted_dii) >= 4:
                    result["DII_4q_ago"] = sorted_dii[3][1]

                if len(sorted_fii) >= 1:
                    result["FII_latest"] = sorted_fii[0][1]
                if len(sorted_fii) >= 2:
                    result["FII_prev"] = sorted_fii[1][1]
                if len(sorted_fii) >= 4:
                    result["FII_4q_ago"] = sorted_fii[3][1]

                return result

        return None

    except Exception as e:
        return None


def compute_changes(data):
    """Compute quarter and 1-year percentage point changes from raw shareholding data."""
    result = {
        "DII Quarter": None,
        "DII 1Yr": None,
        "FII Quarter": None,
        "FII 1Yr": None,
    }

    if data is None:
        return result

    if "DII_latest" in data and "DII_prev" in data:
        val = data["DII_latest"] - data["DII_prev"]
        result["DII Quarter"] = round(val, 2)

    if "DII_latest" in data and "DII_4q_ago" in data:
        val = data["DII_latest"] - data["DII_4q_ago"]
        result["DII 1Yr"] = round(val, 2)

    if "FII_latest" in data and "FII_prev" in data:
        val = data["FII_latest"] - data["FII_prev"]
        result["FII Quarter"] = round(val, 2)

    if "FII_latest" in data and "FII_4q_ago" in data:
        val = data["FII_latest"] - data["FII_4q_ago"]
        result["FII 1Yr"] = round(val, 2)

    return result


def scrape_all_tickers(tickers_ns, update_csv=False):
    """Scrape DII/FII data for all Indian tickers."""
    session = requests.Session()
    session.headers.update(HEADERS)

    results = []
    total = len(tickers_ns)
    found = 0

    for i, ticker in enumerate(tickers_ns):
        ticker_clean = ticker.replace(".NS", "")
        print(f"  [{i+1}/{total}] {ticker_clean}...", end=" ", flush=True)

        data = fetch_shareholding(ticker, session)
        changes = compute_changes(data)
        changes["Ticker"] = ticker

        if changes["DII Quarter"] is not None or changes["FII Quarter"] is not None:
            found += 1
            print(f"DII_Q={changes['DII Quarter']}, FII_Q={changes['FII Quarter']}")
        else:
            print("no data")

        results.append(changes)
        time.sleep(1)  # Rate limit: 1 request per second

    df = pd.DataFrame(results)
    df = df[["Ticker", "DII Quarter", "DII 1Yr", "FII Quarter", "FII 1Yr"]]
    df.to_csv(OUTPUT_CSV, index=False)
    print(f"\nSaved {len(df)} rows to {OUTPUT_CSV}")
    print(f"Found DII/FII data for {found}/{total} tickers")

    if update_csv:
        update_indian_csv(df)

    return df


def update_indian_csv(dii_fii_df):
    """Merge DII/FII data into Indian_Stock_Data.csv."""
    if not os.path.exists(INDIAN_CSV):
        print(f"Warning: {INDIAN_CSV} not found. Cannot update.")
        return

    indian_df = pd.read_csv(INDIAN_CSV)

    for col in ["DII Quarter", "DII 1Yr", "FII Quarter", "FII 1Yr"]:
        if col not in indian_df.columns:
            indian_df[col] = np.nan

    dii_fii_df = dii_fii_df.set_index("Ticker")
    for col in ["DII Quarter", "DII 1Yr", "FII Quarter", "FII 1Yr"]:
        indian_df[col] = indian_df["Ticker"].map(dii_fii_df[col]).fillna(indian_df[col])

    indian_df.to_csv(INDIAN_CSV, index=False)
    print(f"Updated {INDIAN_CSV} with DII/FII data")


if __name__ == "__main__":
    if not os.path.exists(INDIAN_CSV):
        print(f"Error: {INDIAN_CSV} not found. Run collect_indian_data.py first.")
        sys.exit(1)

    indian_df = pd.read_csv(INDIAN_CSV)
    tickers = indian_df["Ticker"].tolist()
    print(f"Loaded {len(tickers)} Indian tickers")

    update_mode = "--update" in sys.argv
    dii_fii_df = scrape_all_tickers(tickers, update_csv=update_mode)

    if not update_mode:
        print(f"\nTo update Indian_Stock_Data.csv, run: python scrape_dii_fii.py --update")