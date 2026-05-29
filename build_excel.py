"""
Build US Stock Analysis Excel with corrected color hierarchies matching sample data patterns.
Also builds SmallCap 600 + MidCap 400 sheet with index-relative coloring.
Also builds NSE 100 sheet from Indian_Stock_Data.csv using the same color system.

Color hierarchies (decoded from sample data, verified against 249 Indian stocks):
  Valuation (PE, Future PE, PEG, Future PEG): Q1=White, Q2=DG, Q3=LG, Q4=LR, Q5=DR. NaN=DR.
    Ultra-cheap = value trap concern (White). Sweet spot in Q2 (DG).
  Risk (Std, Beta): Q1=LR, Q2=White, Q3=LG, Q4=DG, Q5=DR.
    Too safe = missed returns (LR). Sweet spot in Q4 (DG).
  Returns (3M, 6M, 1Yr, 2Yr): Q1=DR, Q2=LR, Q3=White, Q4=LG, Q5=DG. Higher=better.
  Growth (6 columns): Q1=DR, Q2=LR, Q3=White, Q4=LG, Q5=DG. Loss-making profit cols=DR.

Scores: DG=+1, LG=+0.5, White=0, LR=-0.5, DR=-1. Sum per category.
GROWTH SCORE uses only 4 cols (excludes QoQ).
Ratios + Size + DII/FII columns: uncolored.
"""
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os
import warnings

warnings.filterwarnings("ignore")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_PATH = os.path.join(SCRIPT_DIR, "US_Stock_Data.csv")
SMALLMID_CSV_PATH = os.path.join(SCRIPT_DIR, "SmallMidCap_Stock_Data.csv")
OUTPUT_PATH = os.path.join(SCRIPT_DIR, "US_Stock_Analysis_Coloured.xlsx")

# --- Fills ---
FILLS = {
    "DG": PatternFill(start_color="34CF58", end_color="34CF58", fill_type="solid"),
    "LG": PatternFill(start_color="99F2BB", end_color="99F2BB", fill_type="solid"),
    "White": PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid"),
    "LR": PatternFill(start_color="E3CACA", end_color="E3CACA", fill_type="solid"),
    "DR": PatternFill(start_color="F5AEAE", end_color="F5AEAE", fill_type="solid"),
}
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=10, name="Calibri")
DATA_FONT = Font(size=10, name="Calibri")
BORDER = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)

SCORE_MAP = {"DG": 1.0, "LG": 0.5, "White": 0.0, "LR": -0.5, "DR": -1.0}

# --- Column Definitions ---
# (col_key, header_name, format_str)
ALL_COLUMNS = [
    # Core (no color)
    ("Ticker", "Ticker", None),
    ("Sector", "Sector", None),
    ("Sub Sector", "Sub Sector", None),
    # Growth (6 columns)
    ("Sales YoY Growth", "Sales YoY Growth", "0.00"),
    ("NetProfit YoY Growth", "NetProfit YoY Growth", "0.00"),
    ("Sales TTM 1Yr Growth", "Sales TTM 1Yr Growth", "0.00"),
    ("NetProfit TTM 1Yr Growth", "NetProfit TTM 1Yr Growth", "0.00"),
    ("QoQ Sales Growth", "QoQ Sales Growth", "0.00"),
    ("QoQ Profit Growth", "QoQ Profit Growth", "0.00"),
    # Returns (4 columns)
    ("3M Return", "3M Return", "0.00"),
    ("6M Return", "6M Return", "0.00"),
    ("1Yr Return", "1Yr Return", "0.00"),
    ("2Yr Return", "2Yr Return", "0.00"),
    # Valuation (4 columns)
    ("PE Ratio", "PE Ratio", "0.00"),
    ("Future PE", "Future PE", "0.00"),
    ("TTM PEG", "TTM PEG", "0.00"),
    ("Future PEG", "Future PEG", "0.00"),
    # Ratios (3 columns, UNCOLORED)
    ("PB Ratio", "PB Ratio", "0.00"),
    ("EV/Sales", "EV/Sales", "0.00"),
    ("EV/EBITDA", "EV/EBITDA", "0.00"),
    # Size (3 columns, UNCOLORED)
    ("Market Cap (Billions)", "Market Cap (B)", "0.00"),
    ("Revenue (Billions)", "Revenue (B)", "0.00"),
    ("TTM Revenue (Billions)", "TTM Revenue (B)", "0.00"),
    # Risk (4 columns)
    ("QtrStd", "QtrStd", "0.00"),
    ("YrStd", "YrStd", "0.00"),
    ("Qtr Beta", "Qtr Beta", "0.0000"),
    ("Yr Beta", "Yr Beta", "0.0000"),
    # DII/FII (4 columns, UNCOLORED)
    ("DII Quarter", "DII Quarter", "0.00"),
    ("DII 1Yr", "DII 1Yr", "0.00"),
    ("FII Quarter", "FII Quarter", "0.00"),
    ("FII 1Yr", "FII 1Yr", "0.00"),
    # Scores (4 columns, no color)
    ("RETURN SCORE", "RETURN SCORE", "0.00"),
    ("GROWTH SCORE", "GROWTH SCORE", "0.00"),
    ("VALUATION SCORE", "VALUATION SCORE", "0.00"),
    ("RISK SCORE", "RISK SCORE", "0.00"),
]

GROWTH_COLS = [
    "Sales YoY Growth", "NetProfit YoY Growth",
    "Sales TTM 1Yr Growth", "NetProfit TTM 1Yr Growth",
    "QoQ Sales Growth", "QoQ Profit Growth",
]
# GROWTH SCORE only uses 4 cols (excludes QoQ)
GROWTH_SCORE_COLS = [
    "Sales YoY Growth", "NetProfit YoY Growth",
    "Sales TTM 1Yr Growth", "NetProfit TTM 1Yr Growth",
]
RETURN_COLS = ["3M Return", "6M Return", "1Yr Return", "2Yr Return"]
VALUATION_COLS = ["PE Ratio", "Future PE", "TTM PEG", "Future PEG"]
RATIO_COLS = ["PB Ratio", "EV/Sales", "EV/EBITDA"]
SIZE_COLS = ["Market Cap (Billions)", "Revenue (Billions)", "TTM Revenue (Billions)"]
RISK_COLS = ["QtrStd", "YrStd", "Qtr Beta", "Yr Beta"]

# Loss flags -> column mapping
LOSS_FLAG_MAP = {
    "NetProfit YoY Growth": "_Loss_Profit_YoY",
    "NetProfit TTM 1Yr Growth": "_Loss_Profit_TTM",
    "QoQ Profit Growth": "_Loss_Profit_QoQ",
}

UNCOLORED_COLS = RATIO_COLS + SIZE_COLS + ["DII Quarter", "DII 1Yr", "FII Quarter", "FII 1Yr"]

# --- Indian sheet column layout (matches sample data) ---
INDIAN_ALL_COLUMNS = [
    # Core
    ("Index Name", "Index Name", None),
    ("Ticker", "NseCode", None),
    ("Sector", "Sector", None),
    ("Sub Sector", "Sub Sector", None),
    # Growth
    ("Sales YoY Growth", "Sales YoY Growth", "0.00"),
    ("NetProfit YoY Growth", "NetProfit YoY Growth", "0.00"),
    ("Sales TTM 1Yr Growth", "Sales TTM 1Yr Growth", "0.00"),
    ("NetProfit TTM 1Yr Growth", "NetProfit TTM 1Yr Growth", "0.00"),
    # Returns
    ("3M Return", "3M Return", "0.00"),
    ("6M Return", "6M Return", "0.00"),
    ("1Yr Return", "1Yr Return", "0.00"),
    ("2Yr Return", "2Yr Return", "0.00"),
    # Valuation
    ("PE Ratio", "PE Ratio", "0.00"),
    ("Future PE", "Future PE", "0.00"),
    ("TTM PEG", "TTM PEG", "0.00"),
    ("Future PEG", "Future PEG", "0.00"),
    # Composite scores (no color)
    ("Alpha", "Alpha", "0.00"),
    ("Risk", "Risk", "0.00"),
    ("Final Score", "Final Score", "0.00"),
    # Institutional (uncolored, empty for now)
    ("DII Quarter", "DII Quarter", "0.00"),
    ("DII 1Yr", "DII 1Yr", "0.00"),
    ("FII Quarter", "FII Quarter", "0.00"),
    ("FII 1Yr", "FII 1Yr", "0.00"),
    # Risk
    ("QtrStd", "QtrStd", "0.00"),
    ("YrStd", "YrStd", "0.00"),
    ("Qtr Beta", "Qtr Beta", "0.0000"),
    ("Yr Beta", "Yr Beta", "0.0000"),
    # Scores (no color)
    ("RETURN SCORE", "RETURN SCORE", "0.00"),
    ("GROWTH SCORE", "GROWTH SCORE", "0.00"),
    ("VALUATION SCORE", "VALUATION SCORE", "0.00"),
    ("RISK SCORE", "RISK SCORE", "0.00"),
    # Admin (uncolored, empty)
    ("Rebalance Date", "Rebalance Date", None),
    ("Future Return", "Future Return", "0.00"),
    ("Strategy Stocks", "Strategy Stocks", None),
    ("Stocks List", "Stocks List", None),
]


def _quantile_bounds(series, n_groups):
    clean = series.dropna()
    if len(clean) < n_groups:
        return None
    return [clean.quantile(i / n_groups) for i in range(1, n_groups)]


# ============================================================
# COLOR FUNCTIONS (corrected per sample data patterns)
# ============================================================

def color_valuation(df, col):
    """Q1(lowest PE)=White, Q2=DG, Q3=LG, Q4=LR, Q5(highest)=DR. NaN=DR.
    Ultra-cheap = value trap (White). Sweet spot in Q2 (DG)."""
    series = pd.to_numeric(df[col], errors="coerce")
    valid = series.dropna()

    if len(valid) < 5:
        return pd.Series("White", index=df.index)

    q20 = valid.quantile(0.20)
    q40 = valid.quantile(0.40)
    q60 = valid.quantile(0.60)
    q80 = valid.quantile(0.80)

    def assign(v):
        if pd.isna(v):
            return "DR"
        if v <= q20:
            return "White"
        if v <= q40:
            return "DG"
        if v <= q60:
            return "LG"
        if v <= q80:
            return "LR"
        return "DR"

    return series.apply(assign)


def color_risk(df, col):
    """Q1(lowest risk)=LR, Q2=White, Q3=LG, Q4=DG, Q5(highest)=DR.
    Too safe = missed returns (LR). Sweet spot in Q4 (DG)."""
    series = pd.to_numeric(df[col], errors="coerce")
    valid = series.dropna()

    if len(valid) < 5:
        return pd.Series("White", index=df.index)

    q20 = valid.quantile(0.20)
    q40 = valid.quantile(0.40)
    q60 = valid.quantile(0.60)
    q80 = valid.quantile(0.80)

    def assign(v):
        if pd.isna(v):
            return "White"
        if v <= q20:
            return "LR"
        if v <= q40:
            return "White"
        if v <= q60:
            return "LG"
        if v <= q80:
            return "DG"
        return "DR"

    return series.apply(assign)


def color_returns(df, col):
    """Standard ascending: Q1(worst)=DR, Q2=LR, Q3=White, Q4=LG, Q5(best)=DG."""
    series = pd.to_numeric(df[col], errors="coerce")
    valid = series.dropna()

    if len(valid) < 5:
        return pd.Series("White", index=df.index)

    q20 = valid.quantile(0.20)
    q40 = valid.quantile(0.40)
    q60 = valid.quantile(0.60)
    q80 = valid.quantile(0.80)

    def assign(v):
        if pd.isna(v):
            return "White"
        if v <= q20:
            return "DR"
        if v <= q40:
            return "LR"
        if v <= q60:
            return "White"
        if v <= q80:
            return "LG"
        return "DG"

    return series.apply(assign)


def color_growth(df, col):
    """Standard ascending: Q1(worst)=DR, Q2=LR, Q3=White, Q4=LG, Q5(best)=DG.
    Loss-making profit columns -> DR. NaN -> White."""
    series = pd.to_numeric(df[col], errors="coerce")
    loss_flag = LOSS_FLAG_MAP.get(col)

    if loss_flag and loss_flag in df.columns:
        loss_mask = df[loss_flag].fillna(False)
    else:
        loss_mask = pd.Series(False, index=df.index)

    profitable = series[~loss_mask].dropna()

    if len(profitable) < 4:
        def fallback(v, is_loss):
            if pd.isna(v):
                return "White"
            return "DR" if is_loss else "White"
        result = pd.Series("White", index=df.index)
        for idx in df.index:
            result.at[idx] = fallback(
                series.at[idx] if idx in series.index else None,
                loss_mask.at[idx] if idx in loss_mask.index else False)
        return result

    q20 = profitable.quantile(0.20)
    q40 = profitable.quantile(0.40)
    q60 = profitable.quantile(0.60)
    q80 = profitable.quantile(0.80)

    def assign(v, is_loss):
        if pd.isna(v):
            return "White"
        if is_loss:
            return "DR"
        if v <= q20:
            return "DR"
        if v <= q40:
            return "LR"
        if v <= q60:
            return "White"
        if v <= q80:
            return "LG"
        return "DG"

    result = pd.Series("White", index=df.index)
    for idx in df.index:
        result.at[idx] = assign(
            series.at[idx] if idx in series.index else None,
            loss_mask.at[idx] if idx in loss_mask.index else False)
    return result


# ============================================================
# DATA COLLECTION HELPERS (for Small/Mid cap)
# ============================================================

REVENUE_ROWS = ["TotalRevenue", "Total Revenue", "Net Revenue"]
INCOME_ROWS = [
    "NetIncome", "Net Income", "Net Income Common Stockholders",
    "Net Income Continuous Operations",
]


def safe_round(val, decimals=2):
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return None
    return round(float(val), decimals)


def try_financials_row(fin, row_names, col_name):
    row = next((r for r in row_names if r in fin.index), None)
    if row:
        val = fin.loc[row, col_name]
        if pd.notna(val):
            return val
    return None


# ============================================================
# SMALL CAP 600 + MID CAP 400 SHEET
# ============================================================

def build_small_mid_sheet(wb):
    """Build SmallCap 600 + MidCap 400 sheet from CSV with index-relative coloring."""
    print("\n=== Building SmallCap 600 + MidCap 400 Sheet ===\n")

    if not os.path.exists(SMALLMID_CSV_PATH):
        print(f"Warning: {SMALLMID_CSV_PATH} not found. Skipping SmallMidCap sheet.")
        return None

    combined = pd.read_csv(SMALLMID_CSV_PATH)
    combined = combined.drop_duplicates(subset=["Ticker"])
    combined = combined.sort_values(["Index", "Ticker"]).reset_index(drop=True)

    print(f"Loaded {len(combined)} stocks from {SMALLMID_CSV_PATH}")

    # Index-relative coloring: color each index group separately, then merge
    all_colors = []
    all_scores = []

    for idx_label, idx_df in combined.groupby("Index"):
        idx_df = idx_df.reset_index(drop=True)
        print(f"\n--- {idx_label}: {len(idx_df)} stocks ---")

        # Assign colors within index group
        color_idx = assign_colors(idx_df)
        color_idx["_idx"] = idx_df["Index"].values

        scores_idx = compute_scores(idx_df, color_idx)

        all_colors.append(color_idx)
        all_scores.append(scores_idx)

        for sc in ["RETURN SCORE", "GROWTH SCORE", "VALUATION SCORE", "RISK SCORE"]:
            vals = scores_idx[sc]
            print(f"  {sc}: min={vals.min():.1f}, max={vals.max():.1f}, median={vals.median():.1f}")

    final_colors = pd.concat(all_colors, ignore_index=True)
    final_scores = pd.concat(all_scores, ignore_index=True)

    # Clean up for display
    combined_display = combined.drop(columns=[c for c in combined.columns if c.startswith("_Loss_")], errors="ignore")

    # Write sheet
    ws = wb.create_sheet("SmallMidCap Analysis")
    write_sheet(ws, combined_display, final_colors, final_scores)
    set_column_widths(ws)
    ws.freeze_panes = "D2"
    last_col = get_column_letter(len(ALL_COLUMNS))
    ws.auto_filter.ref = f"A1:{last_col}{len(combined_display) + 1}"
    add_legend(ws, len(combined_display) + 3)

    print(f"\nSmallMidCap sheet complete: {len(combined)} stocks")
    return final_scores



# ============================================================
# NSE 100 SHEET
# ============================================================

def build_indian_sheet(wb):
    """Build NSE 100 sheet from Indian_Stock_Data.csv."""
    indian_csv = os.path.join(SCRIPT_DIR, "Indian_Stock_Data.csv")
    if not os.path.exists(indian_csv):
        print(f"Warning: {indian_csv} not found. Skipping NSE 100 sheet.")
        return None

    df = pd.read_csv(indian_csv)
    df = df.drop_duplicates(subset=["Ticker"])
    df = df.sort_values("Ticker").reset_index(drop=True)

    loss_cols = [c for c in df.columns if c.startswith("_Loss_")]
    df_clean = df.drop(columns=loss_cols, errors="ignore")

    print(f"Building NSE 100 sheet: {len(df_clean)} stocks, {len(INDIAN_ALL_COLUMNS)} columns")

    color_df = assign_colors(df)
    scores = compute_scores(df, color_df)

    # Compute composite scores: Alpha = RETURN + GROWTH, Risk = RISK, Final = all 4
    extra_scores = pd.DataFrame(index=df.index)
    extra_scores["Alpha"] = scores["RETURN SCORE"] + scores["GROWTH SCORE"]
    extra_scores["Risk"] = scores["RISK SCORE"]
    extra_scores["Final Score"] = (scores["RETURN SCORE"] + scores["GROWTH SCORE"]
                                      + scores["VALUATION SCORE"] + scores["RISK SCORE"])
    extra_scores = extra_scores.round(2)

    ws = wb.create_sheet("NSE 100 Analysis")
    write_sheet(ws, df_clean, color_df, scores, columns=INDIAN_ALL_COLUMNS, extra_scores=extra_scores)
    set_column_widths(ws, columns=INDIAN_ALL_COLUMNS)
    ws.freeze_panes = "E2"  # Freeze after Index Name + NseCode + Sector
    last_col = get_column_letter(len(INDIAN_ALL_COLUMNS))
    ws.auto_filter.ref = f"A1:{last_col}{len(df_clean) + 1}"
    add_legend(ws, len(df_clean) + 3)

    print("=== NSE 100 Score Ranges ===")
    for sc in ["RETURN SCORE", "GROWTH SCORE", "VALUATION SCORE", "RISK SCORE", "Alpha", "Risk", "Final Score"]:
        if sc in extra_scores.columns:
            vals = extra_scores[sc]
        else:
            vals = scores[sc]
        print(f"  {sc}: min={vals.min():.1f}, max={vals.max():.1f}, median={vals.median():.1f}")

    return scores


# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":
    import sys
    sp500_only = "--sp500-only" in sys.argv
    smallmid_only = "--smallmid-only" in sys.argv
    india_only = "--india-only" in sys.argv
    us_only = "--us-only" in sys.argv

    if smallmid_only:
        from openpyxl import load_workbook
        wb = load_workbook(OUTPUT_PATH)
        build_small_mid_sheet(wb)
        wb.save(OUTPUT_PATH)
        print(f"\nSaved Excel to {OUTPUT_PATH}")
        print("Sheets:", wb.sheetnames)
    elif india_only:
        wb = Workbook()
        wb.remove(wb.active)
        result = build_indian_sheet(wb)
        if result is not None and wb.sheetnames:
            wb.save(OUTPUT_PATH)
            print(f"\nSaved Excel to {OUTPUT_PATH}")
            print("Sheets:", wb.sheetnames)
        else:
            print("\nNo sheets built. Skipping save.")
    else:
        wb = Workbook()
        wb.remove(wb.active)

        build_sp500_sheet(wb)

        if not sp500_only:
            build_small_mid_sheet(wb)
        else:
            print("\nSkipping SmallMidCap sheet (--sp500-only flag set)")

        if not us_only:
            build_indian_sheet(wb)
        else:
            print("\nSkipping NSE 100 sheet (--us-only flag set)")

        wb.save(OUTPUT_PATH)
        print(f"\nSaved Excel to {OUTPUT_PATH}")
        print("Sheets:", wb.sheetnames)
