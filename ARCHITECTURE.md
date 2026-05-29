# AnjaliValueStocks — Full Automation Architecture

**Project**: Anjali Value Stocks Screener  
**Markets**: US (S&P 500 + SmallCap 600 + MidCap 400) + India (NSE 250 / LM 250)  
**Output**: Multi-sheet color-coded Excel with composite scoring  
**Automation**: Railway 24/7 via APScheduler  
**Last Updated**: 2026-05-29

---

## 1. Pipeline Architecture

```
┌───────────────────────────────────────────────────────────────────────────────┐
│  scheduler.py (APScheduler) — runs 24/7 on Railway                     │
│  ───────────────────────────────────────────────────────────────────────────────────└──┬──┐
│                                            │  │  │
│    Every 6h          Every 30min           │  │  │
│      ↓                 ↓                  │  │  │
├───────────────────────────────────────────────────────────────────────────────┤  │  │  │
│  Full Refresh      Price Refresh         │  │  │
│  (collect_all)     (lightweight)         │  │  │
│      │                │                  │  │  │
│      │                │                  │  │  │
└────┬──────┬───────┬────────────────────────────────────────────────────────────────────────────────┘  │  │  │
     │    │      │                                              │  │  │
     │    │      │                                              │  │  │
  ┌────────────┐ │  │                                       │  │  │
  │  US Data    │ │  │                                       │  │  │
  │ collect_data │ │  │                                       │  │  │
  └────────────┘ │  │                                       │  │  │
       │          │  │                                       │  │  │
       ↓          │  │                                       │  │  │
  ┌────────────────┐ │  │                                       │  │  │
  │ Indian Data    │ │  │                                       │  │  │
  │collect_indian│ │  │                                       │  │  │
  └────────────────┘ │  │                                       │  │  │
       │          │  │                                       │  │  │
       └───────────┘  │                                       │  │  │
            │              │                                       │  │  │
            ↓              ↓                                       │  │  │
       ┌───────────────────────────────────────────────────────────────────────────────┐  │  │  │
       │  build_excel.py (Excel builder)                             │  │  │
       │    ───────────────────────────────────────────────────────────────────────────────────────────────└──┬──┘  │  │
       │                                                         │     │  │
       │    ┌─────────────────┐  ┌─────────────────┐  ┌────────────────┐ │     │  │
       │    │ S&P 500        │  │ SmallMidCap     │  │ NSE 100     │ │     │  │
       │    │ Sheet 1        │  │ Sheet 2         │  │ Sheet 3     │ │     │  │
       │    └─────────────────┘  └─────────────────┘  └────────────────┘ │     │  │
       │                                                         │     │  │
       │    ↓                                                    │     │  │
       └───────────────────────────────────────────────────────────────────────────────────────────────────────┘     │  │
                                                              ↓     │  │
                                                    ┌───────────────────────────┐  │  │
                                                    │ US_Stock_Analysis_Coloured │  │  │
                                                    │    .xlsx (Output)          │  │  │
                                                    └───────────────────────────┘  │  │
                                                                                    │  │
                                                                                    │  │
                                                                              ┌───────┐  │
                                                                              │ Railway │←──┘
                                                                              │ Deploy  │
                                                                              └───────┘
```

| Component | File | Input | Output | Frequency | Runtime |
|-----------|------|-------|--------|-----------|---------|
| **US Collector** | `collect_data.py` | Wikipedia S&P 500 + yfinance | `US_Stock_Data.csv` | Every 6h | ~25 min |
| **Indian Collector** | `collect_indian_data.py` | Hardcoded 250 NSE + yfinance | `Indian_Stock_Data.csv` | Every 6h | ~5 min |
| **Excel Builder** | `build_excel.py` | 2 CSVs + Wikipedia S&P 400/600 | `US_Stock_Analysis_Coloured.xlsx` | After full refresh | ~50 min |
| **Scheduler** | `scheduler.py` | APScheduler cron | Logs + Excel | 24/7 | Always running |

---

## 2. Data Flow

| Step | Action | Source | Destination |
|------|--------|--------|-------------|
| 1 | Fetch S&P 500 tickers | Wikipedia | `collect_data.py` |
| 2 | Fetch S&P 400/600 tickers | Wikipedia | `build_excel.py` (inline) |
| 3 | Fetch 250 NSE tickers | Hardcoded list | `collect_indian_data.py` |
| 4 | Collect US fundamentals | yfinance | `US_Stock_Data.csv` |
| 5 | Collect Indian fundamentals | yfinance (.NS suffix) | `Indian_Stock_Data.csv` |
| 6 | Collect SmallMid fundamentals | yfinance | `build_excel.py` (memory) |
| 7 | Compute colors + scores | pandas quantiles | Color DataFrame |
| 8 | Write 3-sheet Excel | openpyxl | `.xlsx` file |
| 9 | Health check | Railway | HTTP 200 |

---

## 3. Column Mappings

### US Sheet — 31 Columns

| # | Column | Source | Color Category |
|---|--------|--------|----------------|
| 1 | Ticker | Wikipedia / yfinance | None |
| 2 | Sector | yfinance `.info` | None |
| 3 | Sub Sector | yfinance `.info` | None |
| 4 | Sales YoY Growth | yfinance `.info` | Growth |
| 5 | NetProfit YoY Growth | yfinance `.info` | Growth |
| 6 | Sales TTM 1Yr Growth | yfinance `financials` | Growth |
| 7 | NetProfit TTM 1Yr Growth | yfinance `financials` | Growth |
| 8 | QoQ Sales Growth | yfinance `quarterly_financials` | Growth |
| 9 | QoQ Profit Growth | yfinance `quarterly_financials` | Growth |
| 10 | 3M Return | yfinance `history` | Returns |
| 11 | 6M Return | yfinance `history` | Returns |
| 12 | 1Yr Return | yfinance `history` | Returns |
| 13 | 2Yr Return | yfinance `history` | Returns |
| 14 | PE Ratio | yfinance `.info` | Valuation |
| 15 | Future PE | Calculated | Valuation |
| 16 | TTM PEG | Calculated | Valuation |
| 17 | Future PEG | Calculated | Valuation |
| 18 | PB Ratio | yfinance `.info` | UNCOLORED |
| 19 | EV/Sales | yfinance `.info` | UNCOLORED |
| 20 | EV/EBITDA | yfinance `.info` | UNCOLORED |
| 21 | Market Cap (B) | yfinance `.info` | UNCOLORED |
| 22 | Revenue (B) | yfinance `.info` | UNCOLORED |
| 23 | TTM Revenue (B) | yfinance `financials` | UNCOLORED |
| 24 | QtrStd | yfinance `history` | Risk |
| 25 | YrStd | yfinance `history` | Risk |
| 26 | Qtr Beta | Cov/Var vs SPY | Risk |
| 27 | Yr Beta | Cov/Var vs SPY | Risk |
| 28 | RETURN SCORE | Sum of 4 returns colors | None |
| 29 | GROWTH SCORE | Sum of 4 growth colors | None |
| 30 | VALUATION SCORE | Sum of 4 valuation colors | None |
| 31 | RISK SCORE | Sum of 4 risk colors | None |

### Indian Sheet — 37 Columns

| # | Column | Source | Color Category | Notes |
|---|--------|--------|----------------|-------|
| 1 | Index Name | Hardcoded "LM 250" | None | Matches sample |
| 2 | NseCode | Hardcoded list | None | Ticker without .NS |
| 3 | Sector | yfinance `.info` | None | |
| 4 | Sub Sector | yfinance `.info` | None | |
| 5-8 | Growth (4 cols) | Same as US | Growth | No QoQ in display |
| 9-12 | Returns (4 cols) | Same as US | Returns | |
| 13-16 | Valuation (4 cols) | Same as US | Valuation | |
| 17 | Alpha | Computed | None | RETURN + GROWTH scores |
| 18 | Risk | Computed | None | RISK score |
| 19 | Final Score | Computed | None | Sum of all 4 scores |
| 20-23 | DII/FII (4 cols) | Empty placeholder | UNCOLORED | Not available via yfinance |
| 24-27 | Risk metrics | Same as US | Risk | |
| 28-31 | Scores (4 cols) | Computed | None | |
| 32 | Rebalance Date | Empty | UNCOLORED | Admin column |
| 33 | Future Return | Empty | UNCOLORED | Admin column |
| 34 | Strategy Stocks | Empty | UNCOLORED | Admin column |
| 35 | Stocks List | Empty | UNCOLORED | Admin column |

---

## 4. Color Hierarchies

| Category | Q1 | Q2 | Q3 | Q4 | Q5 | Sort | Special Rules |
|----------|----|----|----|----|-----|------|--------------|
| **Growth** | DR (-1) | LR (-0.5) | White (0) | LG (+0.5) | **DG (+1)** | Ascending | Loss = DR override |
| **Returns** | DR (-1) | LR (-0.5) | White (0) | LG (+0.5) | **DG (+1)** | Ascending | Higher = better |
| **Valuation** | White (0) | **DG (+1)** | LG (+0.5) | LR (-0.5) | DR (-1) | Ascending | NaN = DR |
| **Risk** | LR (-0.5) | White (0) | LG (+0.5) | **DG (+1)** | DR (-1) | Ascending | Moderate-high = sweet spot |

Colors: DG=#34CF58, LG=#99F2BB, White=#FFFFFF, LR=#E3CACA, DR=#F5AEAE

---

## 5. Scoring Formulas

| Score | Components | Formula | Range |
|-------|-----------|---------|-------|
| **RETURN SCORE** | 3M + 6M + 1Yr + 2Yr | Sum of color scores | -4 to +4 |
| **GROWTH SCORE** | Sales YoY + NetProfit YoY + Sales TTM + NetProfit TTM | Sum of color scores | -4 to +4 |
| **VALUATION SCORE** | PE + Future PE + TTM PEG + Future PEG | Sum of color scores | -4 to +4 |
| **RISK SCORE** | QtrStd + YrStd + Qtr Beta + Yr Beta | Sum of color scores | -4 to +4 |
| **Alpha** | RETURN + GROWTH | `RETURN_SCORE + GROWTH_SCORE` | -8 to +8 |
| **Risk** | RISK | `RISK_SCORE` | -4 to +4 |
| **Final Score** | All 4 categories | `RETURN + GROWTH + VALUATION + RISK` | -16 to +16 |

Score map: DG=+1, LG=+0.5, White=0, LR=-0.5, DR=-1

---

## 6. Key Formulas

| Metric | Formula |
|--------|---------|
| **Returns** | `(Price_now / Price_start - 1) × 100` |
| **Annualized Std** | `daily_returns.std() × sqrt(252) × 100` |
| **Beta (US)** | `Cov(stock_daily, SPY_daily) / Var(SPY_daily)` |
| **Beta (India)** | `Cov(stock_daily, ^NSEI_daily) / Var(^NSEI_daily)` |
| **Future PE** | `PE_Ratio × (1 + NetProfit_TTM_Growth / 100)` |
| **TTM PEG** | `PE_Ratio / NetProfit_TTM_Growth` |
| **Future PEG** | `Future_PE / NetProfit_TTM_Growth` |
| **TTM Growth** | `(financials[latest] / financials[prior] - 1) × 100` |
| **QoQ Growth** | `(quarterly[latest] / quarterly[prior] - 1) × 100` |
| **QoQ Profit Cap** | Clamped to `[-500%, +500%]` |

---

## 7. Cron Schedule

| Job | Expression | Meaning | Action |
|-----|-----------|---------|--------|
| **Full Refresh** | `0 0,6,12,18 * * *` | Every 6 hours at :00 | `collect_data.py` + `collect_indian_data.py` + `build_excel.py` |
| **Price Refresh** | `*/30 * * * *` | Every 30 minutes | Placeholder (delegates to full refresh for now) |
| **Health Log** | Every 10 minutes | Internal | Prints alive status to scheduler.log |

---

## 8. CLI Flags

| Flag | Script | Effect |
|------|--------|--------|
| `--sp500-only` | `build_excel.py` | Build only S&P 500 sheet |
| `--smallmid-only` | `build_excel.py` | Build only SmallMidCap sheet (adds to existing Excel) |
| `--india-only` | `build_excel.py` | Build only NSE 100 sheet |
| `--us-only` | `build_excel.py` | Build US + SmallMid, skip Indian |

---

## 9. Environment Variables

| Variable | Default | Purpose |
|----------|---------|---------|
| `PYTHONUNBUFFERED` | `1` | Force stdout flush for Railway logs |
| `TZ` | `Asia/Kolkata` | Timezone for cron scheduling |
| `RAILWAY_ENVIRONMENT` | — | Set by Railway; used to detect production mode |

---

## 10. Railway Deployment

| Step | Command / Action |
|------|-----------------|
| 1. Create Railway project | `railway login` → `railway init` |
| 2. Link GitHub repo | Railway Dashboard → Settings → GitHub Repo |
| 3. Deploy | `railway up` (or auto-deploy on push) |
| 4. Add volume | Railway Dashboard → Volumes → Add `/app/data` |
| 5. View logs | `railway logs` or Dashboard |
| 6. Download Excel | Volume file explorer or `railway ssh` → `cat` |

**Railway Config**: `railway.json`
- Builder: Dockerfile
- Start: `python scheduler.py`
- Restart: ON_FAILURE, max 10 retries

---

## 11. File Inventory

| File | Type | Purpose | Lines |
|------|------|---------|-------|
| `collect_data.py` | Active | US S&P 500 collector | ~350 |
| `collect_indian_data.py` | Active | Indian NSE 250 collector | ~380 |
| `build_excel.py` | Active | Excel builder (3 sheets) | ~450 |
| `scheduler.py` | Active | APScheduler orchestrator | ~150 |
| `requirements.txt` | Config | Python dependencies | 6 |
| `Dockerfile` | Config | Container image | 30 |
| `railway.json` | Config | Railway deployment | 12 |
| `docker-compose.yml` | Config | Local Docker testing | 20 |
| `.gitignore` | Config | Git exclusions | 40 |
| `.dockerignore` | Config | Docker exclusions | 20 |
| `README.md` | Docs | Project overview | 280 |
| `ARCHITECTURE.md` | Docs | This file | — |
| `build_us_stock_sheet.py` | Archived | Old 3-script pipeline | 681 |
| `enrich_data.py` | Archived | Old enricher | 123 |
| `US_Stock_Data.csv` | Data | US 503 stocks | — |
| `Indian_Stock_Data.csv` | Data | Indian 250 stocks | — |
| `US_Stock_Analysis_Coloured.xlsx` | Output | 3-sheet Excel | — |

---

## 12. Known Limitations

| # | Limitation | Impact | Workaround |
|---|-----------|--------|-----------|
| 1 | DII/FII data unavailable via yfinance | Indian sheet has empty institutional columns | Manual update or future NSE scraper |
| 2 | 30-minute price refresh not truly lightweight | Full collection takes ~70 min | Run every 6h instead; true lightweight refresh requires paid data API |
| 3 | Railway free tier sleeps after inactivity | Scheduler stops if no traffic | Use paid Railway plan or external keep-alive ping |
| 4 | NSE tickers are hardcoded | List may drift from actual index | Annual review; no reliable free API for NSE constituents |
| 5 | yfinance rate limits | Occasional 429 errors | 0.3s delay per stock; checkpoint resume |
| 6 | Sector-relative coloring not implemented | All stocks compared globally | Planned enhancement |
| 7 | SmallMidCap data collected inline | Re-collected on every Excel build | Could be separated to own collector |

---

## 13. Tech Stack

| Layer | Technology |
|-------|-----------|
| **Language** | Python 3.11 |
| **Data Source** | yfinance (Yahoo Finance) |
| **Data Processing** | pandas, numpy |
| **Excel Generation** | openpyxl |
| **Scheduling** | APScheduler (BackgroundScheduler) |
| **Web Scraping** | requests + pandas.read_html |
| **Container** | Docker (python:3.11-slim) |
| **Hosting** | Railway |
| **Version Control** | Git + GitHub |

---

*End of Architecture Document*
