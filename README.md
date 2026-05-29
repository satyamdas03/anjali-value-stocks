# Anjali Value Stocks

S&P 500 + SmallCap 600 + MidCap 400 stock screener with multi-sheet quintile color-coded Excel analysis. Index-relative coloring, composite scoring system, loss-aware growth metrics.

Built with `yfinance` + `openpyxl`. Two-script pipeline: collect data once, generate colored Excel.

---

## Pipeline Architecture

```
collect_data.py                    build_excel.py
  Input: Wikipedia S&P 500 list     Input: US_Stock_Data.csv
  Process: yfinance (info,           Process: color_growth/returns/
           history, financials,               valuation/risk + sum scores
           quarterly_financials,              + quintile boundaries
           beta vs SPY)              Output: US_Stock_Analysis_Coloured.xlsx
  Output: US_Stock_Data.csv                  Sheets:
  Checkpoints: every 50 stocks                - S&P 500 Analysis (503 stocks)
  Rate limit: 0.3s per stock                  - SmallMidCap Analysis
                                               (~1000 stocks, index-relative)
```

---

## Output Excel — 34 Columns

| Section | Columns | Color |
|---------|---------|-------|
| **Core** | Ticker, Sector, Sub Sector | None |
| **Growth** | Sales YoY, NetProfit YoY, Sales TTM 1Yr, NetProfit TTM 1Yr, QoQ Sales, QoQ Profit | Growth hierarchy (loss-aware) |
| **Returns** | 3M, 6M, 1Yr, 2Yr | Returns hierarchy |
| **Valuation** | PE Ratio, Future PE, TTM PEG, Future PEG | Valuation hierarchy |
| **Ratios** | PB Ratio, EV/Sales, EV/EBITDA | UNCOLORED |
| **Size** | Market Cap (B), Revenue (B), TTM Revenue (B) | UNCOLORED |
| **Risk** | QtrStd, YrStd, Qtr Beta, Yr Beta | Risk hierarchy |
| **Scores** | RETURN SCORE, GROWTH SCORE, VALUATION SCORE, RISK SCORE | None (values only) |

---

## Color Hierarchies (Corrected — May 29, 2026)

Pattern-matched from `sample data/Satyam May 26.xlsx` (249-stock Indian NSE screener). Exhaustive line-by-line quintile decoding.

Colors: DG=DarkGreen(#34CF58), LG=LightGreen(#99F2BB), White(#FFFFFF), LR=LightRed(#E3CACA), DR=DarkRed(#F5AEAE)

### Growth (6 columns — ascending sort)

| Quintile | Color | Score | Notes |
|----------|-------|-------|-------|
| Q1 (0-20%, worst growth) | DR | -1 | Lowest growth = worst |
| Q2 (20-40%) | LR | -0.5 | |
| Q3 (40-60%) | White | 0 | |
| Q4 (60-80%) | LG | +0.5 | |
| Q5 (80-100%, best growth) | DG | +1 | Highest growth = best |
| Loss-making profit cols | DR | -1 | Overrides quintile position |

Quantile boundaries calculated from profitable stocks only. Loss flag overrides to DR.

### Returns (4 columns — ascending sort)

| Quintile | Color | Score |
|----------|-------|-------|
| Q1 (0-20%, worst returns) | DR | -1 |
| Q2 (20-40%) | LR | -0.5 |
| Q3 (40-60%) | White | 0 |
| Q4 (60-80%) | LG | +0.5 |
| Q5 (80-100%, best returns) | DG | +1 |

Straightforward: higher returns = better. No special rules.

### Valuation (4 columns — ascending sort, lowest PE first)

| Quintile | Color | Score | Meaning |
|----------|-------|-------|---------|
| Q1 (0-20%, lowest PE/PEG) | White | 0 | Ultra-cheap = value trap concern |
| Q2 (20-40%) | DG | +1 | Good value — **sweet spot** |
| Q3 (40-60%) | LG | +0.5 | Fair value |
| Q4 (60-80%) | LR | -0.5 | Expensive |
| Q5 (80-100%, highest PE/PEG) | DR | -1 | Very expensive |
| NaN PE/PEG | DR | -1 | Missing data = penalized |

Key insight from sample data: lowest quintile (cheapest stocks) gets **White/NOT Green** because extremely low PE often signals a value trap (e.g., CHTR PE 3.8). The **sweet spot** is Q2 — reasonable valuations get the +1 DG reward.

### Risk (4 columns — ascending sort, lowest std/beta first)

| Quintile | Color | Score | Meaning |
|----------|-------|-------|---------|
| Q1 (0-20%, safest) | LR | -0.5 | Too safe = missed returns |
| Q2 (20-40%) | White | 0 | Safe |
| Q3 (40-60%) | LG | +0.5 | Moderate risk |
| Q4 (60-80%) | DG | +1 | Mod-high risk — **sweet spot** |
| Q5 (80-100%, riskiest) | DR | -1 | Too risky |

Risk uses a "moderate-is-best" curve. Safest stocks (Q1) get LR penalty because ultra-low volatility often means ultra-low returns. The **sweet spot** is Q4 — moderate-high risk offers better risk-reward.

---

## Scoring System

| Score | Components | Range |
|-------|-----------|-------|
| RETURN SCORE | 3M + 6M + 1Yr + 2Yr | -4 to +4 |
| GROWTH SCORE | Sales YoY + NetProfit YoY + Sales TTM 1Yr + NetProfit TTM 1Yr | -4 to +4 |
| VALUATION SCORE | PE + Future PE + TTM PEG + Future PEG | -4 to +4 |
| RISK SCORE | QtrStd + YrStd + Qtr Beta + Yr Beta | -4 to +4 |

**GROWTH SCORE excludes QoQ** columns (QoQ Sales Growth, QoQ Profit Growth) — QoQ metrics are noisy and seasonal, only 4 longer-term growth columns are used.

Formula: DG=+1, LG=+0.5, White=0, LR=-0.5, DR=-1. Sum across category.

---

## Index-Relative Coloring (SmallMidCap sheet)

The SmallCap 600 and MidCap 400 sheet uses **index-relative coloring**: quintile boundaries are calculated within each index group separately.

- S&P 400 stocks compared ONLY against other S&P 400 stocks
- S&P 600 stocks compared ONLY against other S&P 600 stocks
- No cross-index comparison — avoids small caps always looking "worse" than mid caps on valuation

Both groups share the same color hierarchies and column layout as the S&P 500 sheet.

---

## Data Collection Details

### S&P 500 (collect_data.py)
- Source: Wikipedia S&P 500 list + yfinance
- Metrics: 27 data columns + 3 loss flags = 30 total
- Future PE: `Current PE x (1 + TTM Profit Growth / 100)` (calculated, not yfinance forwardPE)
- PEG: `PE / NetProfit_TTM_Growth` (unified single formula)
- Beta: `Cov(stock_returns, spy_returns) / Var(spy_returns)` on aligned daily returns
- QoQ Profit capped at +/-500%
- Negative PB ratios: excluded (set to NaN)
- Checkpoint: every 50 stocks, auto-resume on restart
- Rate limit: 0.3s per stock (~25 min for full 503)

### SmallCap 600 + MidCap 400 (build_excel.py inline)
- Source: Wikipedia S&P 400 + S&P 600 lists + yfinance
- Same metrics structure as S&P 500
- SPY benchmark pre-fetched once, reused for all beta calculations
- Collection embedded in build_excel.py for simplicity
- Rate limit: 0.3s per stock (~5-8 min for combined ~1000 stocks)

---

## Key Formulas

| Metric | Formula |
|--------|---------|
| **3M/6M/1Yr/2Yr Returns** | `(Price_now / Price_period_start - 1) x 100` |
| **Annualized Std** | `daily_returns.std() x sqrt(252) x 100` |
| **Beta** | `Cov(stock_daily, spy_daily) / Var(spy_daily)` on aligned dates |
| **Future PE** | `PE_Ratio x (1 + NetProfit_TTM_Growth / 100)` |
| **TTM PEG** | `PE_Ratio / NetProfit_TTM_Growth` |
| **Future PEG** | `Future_PE / NetProfit_TTM_Growth` |
| **TTM Growth** | `(financials[latest_col] / financials[prior_col] - 1) x 100` |
| **QoQ Growth** | `(quarterly_fin[latest] / quarterly_fin[prior] - 1) x 100` |

---

## Project Files

| File | Status | Purpose |
|------|--------|---------|
| `collect_data.py` | Active | S&P 500 data pipeline (~350 lines) |
| `build_excel.py` | Active | Excel builder with 4 color functions, 2 sheets (~450 lines) |
| `build_us_stock_sheet.py` | Archived | Old 3-script pipeline (681 lines) — kept for reference |
| `enrich_data.py` | Archived | Old enricher (123 lines) — kept for reference |
| `US_Stock_Data.csv` | Data | 503 S&P 500 stocks x 30 columns |
| `US_Stock_Analysis_Coloured.xlsx` | Output | 2-sheet Excel with full color coding + scores |
| `MAZG Historical Data.csv` | Reference | Mazagon Dock 267-day price history |
| `stock_analysis_coloured (1).xlsx` | Reference | Original NSE 100 Indian screener |
| `sample data/Satyam May 26.xlsx` | Reference | Color pattern ground truth (249 Indian stocks) |

---

## Running

```bash
# Step 1: Collect S&P 500 data (~25 min, checkpoints every 50 stocks)
python collect_data.py

# Step 2: Build Excel (~5-8 min for Small/Mid cap collection)
python build_excel.py

# Output: US_Stock_Analysis_Coloured.xlsx
#   Sheet 1: "S&P 500 Analysis" (503 stocks)
#   Sheet 2: "SmallMidCap Analysis" (~1000 stocks)
```

---

## Bug History (Chronological)

### Session 1 — May 16, 2026 (Initial Deep Dive)
| # | Bug | Status |
|---|-----|--------|
| 1 | QoQ Profit 62,700% for UNH — tiny denominator explodes ratio | FIXED — capped at +/-500% |
| 2 | Negative PB ratios (ABBV -108.95, DELL -68.76, etc.) — share buybacks > book value | FIXED — excluded from coloring |
| 3 | Negative EV/EBITDA for BRK-B (-2.25) — massive cash holdings | Open (rare, low impact) |
| 4 | Empty MMC ghost row — yfinance fetch failed silently | FIXED — try/except produces NaN row |
| 5 | PEG inconsistency across 3 scripts — 3 different formulas | FIXED — single unified formula |
| 6 | Indian NSE sheet not automated — reference-only xlsx | Open (future enhancement) |
| 7 | No sector-relative coloring — all stocks in one quintile pool | Open (future enhancement) |
| 8 | Hardcoded S&P 500 tickers (247 entries with duplicates) | FIXED — Wikipedia dynamic fetch |
| 9 | Duplicate tickers in list (F x2, O x2, SHW x2, PGR x2, MMC x2, STZ/STZ-B) | FIXED — dedup after fetch |

### Session 2 — May 17, 2026 (Exhaustive Stats)
| # | Bug | Status |
|---|-----|--------|
| 10 | UNH 62,700% makes QoQ Profit quintile coloring useless | FIXED — cap at +/-500% |
| 11 | HON: +33.3% sales QoQ but -813.9% profit QoQ — one-time charge distorts colors | FIXED — loss flags |
| 12 | VRTX: TTM profit -838.1% but YoY profit +61.4% — inconsistent signals | Mitigated (both present) |
| 13 | MAZG distribution-volume pattern at peak — institutional selling signal | Analysis only |
| 14 | Dead beta code in build_excel.py (unreachable `higher_is_better=None` branch) | FIXED — rewritten |
| 15 | DII/FII institutional flow absent from US sheet | Open (future enhancement) |

### Session 3 — May 17, 2026 (Refactoring)
- Pipeline: 3-script → 2-script (consolidated collect_data.py + build_excel.py)
- Tickers: 100 hardcoded → 503 Wikipedia dynamic
- Color: Single `get_quintile_fill()` → 4 separate functions (Growth, Returns, Valuation, Risk)
- Added: Scoring system (4 sum-score columns), loss flags, checkpoints
- Growth: 20:20:30:30 DG:LG:LR:DR with loss-making → DR, no White

### Session 4 — May 18, 2026 (Risk Fix)
| # | Bug | Status |
|---|-----|--------|
| 16 | Risk hierarchy wrong direction — ascending (lowest=DR) → descending (highest=DR) | FIXED — color_risk() rewritten |

### Session 5 — May 29, 2026 (Color Correction + SmallMidCap)
- **Full color hierarchy rewrite**: all 4 categories corrected to match sample data patterns
  - Valuation: Q1(cheapest)=White(value trap), Q2=DG(sweet spot)
  - Risk: Q1(safest)=LR(missed returns), Q4=DG(sweet spot)
  - Returns: standard ascending Q1=DR→Q5=DG
  - Growth: standard ascending, loss-aware
- **GROWTH SCORE**: QoQ columns excluded (4 cols only, -4 to +4)
- **New columns**: Market Cap (B), Revenue (B), TTM Revenue (B)
- **New sheet**: SmallCap 600 + MidCap 400 with index-relative coloring
- **New repo**: Created as private GitHub repo

---

## Color Evolution (History)

| Session | Valuation | Risk | Returns | Growth |
|---------|-----------|------|---------|--------|
| May 16-17b (old) | Q1=DR, Q2=DG, Q3=LG, Q4=White, Q5=LR | Q1=DR→Q5=LR (ascending) | Q1=DR(top), Q2=DG, Q3=LG, Q4=White, Q5=LR | DG(20%) LG(20%) LR(30%) DR(30%), No White |
| May 18 | Same | Q1=LR→Q5=DR (descending) | Same | Same |
| May 29 (current) | Q1=White, Q2=DG, Q3=LG, Q4=LR, Q5=DR | Q1=LR, Q2=White, Q3=LG, Q4=DG, Q5=DR | Q1=DR, Q2=LR, Q3=White, Q4=LG, Q5=DG | Q1=DR, Q2=LR, Q3=White, Q4=LG, Q5=DG |

---

## Tech Stack

- **Python 3.13** — all scripts
- **yfinance** — stock data (price history, financials, fundamentals, market cap)
- **openpyxl** — Excel generation with cell-level color formatting
- **pandas/numpy** — data processing, quantile calculations, beta computation
- **requests** — Wikipedia ticker fetching with User-Agent headers
- **GitHub** — private repository hosting

---

## Future Enhancements

- [ ] Indian NSE 100 pipeline (mirror US for Indian stocks)
- [ ] Sector-relative coloring (toggleable)
- [ ] 13F/DII/FII institutional flow data for US sheet
- [ ] Composite ranking column
- [ ] Time-series diff between data collection runs
- [ ] Scheduled refresh (GitHub Actions cron)
- [ ] Dynamic S&P 500 fetching from latest Wikipedia
- [ ] Individual stock deep-dive (MAZG-style) for any ticker
- [ ] CLI flags (--tickers, --output, --sector-relative, --no-smallmid)

---

*Last updated: May 29, 2026. 5 sessions. 16 bugs. 3 pipeline generations.*
