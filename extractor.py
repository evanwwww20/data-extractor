import pandas as pd
import numpy as np
import re, io
from rapidfuzz import fuzz, process
from dateutil import parser as dateparser

class ExtractionError(Exception):
    pass

# Canonical schema for rent roll
CANON_COLUMNS = ["tenant_name","unit","unit_type","sqft","market_rent","rent"]

# Synonyms / fuzzy labels (extendable)
SYNONYMS = {
    "tenant_name": ["tenant name","resident","name"],
    "unit": ["unit","apt","apartment","door","unit #","unit no"],
    "unit_type": ["unit type","type","floorplan","plan"],
    "sqft": ["sq ft","sqft","area","square feet","sf"],
    "market_rent": ["market rent","market  rent","market\nrent","market $","mrkt rent","sched rent"],
    "rent": ["rent","current rent","actual rent","contract rent","base rent"]
}

# T12 keys
T12_KEYS = {
    "gross_income": ["gross income","total income","effective gross income","egi","revenue total","rental income total"],
    "operating_expenses": ["operating expenses","total operating expenses","opex","op ex","total expenses"],
    # Allow NOI to be computed or taken if present
    "noi": ["net operating income","noi"]
}

CURRENCY_RE = re.compile(r"^\s*\(?\$?\s*-?\d{1,3}(?:,\d{3})*(?:\.\d+)?\)?\s*$")

def _clean_header(s):
    if s is None:
        return ""
    s = str(s)
    s = s.replace("\n"," ").replace("\r"," ")
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s

def _as_number(x):
    if pd.isna(x):
        return np.nan
    s = str(x).strip()
    if s == "":
        return np.nan
    # handle parentheses negatives
    s = s.replace("$","").replace(",","")
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    try:
        return float(s)
    except:
        return np.nan

def _guess_header_row(df: pd.DataFrame, min_non_empty=4):
    for idx in range(min(30, len(df))):
        row = df.iloc[idx]
        non_empty = row.notna().sum()
        # Avoid rows that are clearly titles (single cell filled)
        if non_empty >= min_non_empty:
            return idx
    return 0

def _map_headers(header_vals, synonyms):
    headers = [_clean_header(c) for c in header_vals]
    target_map = {}
    for canon, syns in synonyms.items():
        best_idx = None
        best_score = 0
        for i, h in enumerate(headers):
            if not h:
                continue
            score = max([fuzz.token_set_ratio(h, s) for s in syns])
            if score > best_score:
                best_idx, best_score = i, score
        if best_idx is not None and best_score >= 80:
            target_map[canon] = best_idx
    return target_map

def _detect_data_rows(df, header_row):
    return df.iloc[header_row+1:].reset_index(drop=True)

def _coerce_numeric_cols(df, idxs):
    for i in idxs:
        col = df.columns[i]
        df[col] = df[col].apply(_as_number)
    return df

def extract_from_workbook(content: bytes, filename: str = None):
    """Extract per-unit info from rent roll workbooks with messy headers."""
    try:
        xl = pd.ExcelFile(io.BytesIO(content))
    except Exception as e:
        raise ExtractionError(f"Cannot open workbook: {e}")

    collected = []

    for sheet_name in xl.sheet_names:
        try:
            ws = xl.parse(sheet_name, header=None, dtype=object)
        except Exception:
            continue

        ws = ws.dropna(how="all").reset_index(drop=True).dropna(how="all", axis=1)
        if ws.empty or ws.shape[1] < 3:
            continue

        header_row = _guess_header_row(ws, min_non_empty=4)
        header_vals = ws.iloc[header_row].tolist()
        colmap = _map_headers(header_vals, SYNONYMS)

        data = _detect_data_rows(ws, header_row)
        if data.empty:
            continue

        # Coerce plausible numeric columns
        idxs = [colmap.get("sqft"), colmap.get("market_rent"), colmap.get("rent")]
        idxs = [i for i in idxs if i is not None]
        data = _coerce_numeric_cols(data, idxs)

        out = pd.DataFrame(columns=CANON_COLUMNS)

        out["tenant_name"] = data.iloc[:, colmap["tenant_name"]].astype(str).str.strip() if "tenant_name" in colmap else np.nan
        out["unit"] = data.iloc[:, colmap["unit"]].astype(str).str.strip() if "unit" in colmap else np.nan
        out["unit_type"] = data.iloc[:, colmap["unit_type"]].astype(str).str.strip() if "unit_type" in colmap else np.nan
        out["sqft"] = data.iloc[:, colmap["sqft"]].apply(_as_number) if "sqft" in colmap else np.nan
        out["market_rent"] = data.iloc[:, colmap["market_rent"]].apply(_as_number) if "market_rent" in colmap else np.nan
        out["rent"] = data.iloc[:, colmap["rent"]].apply(_as_number) if "rent" in colmap else np.nan

        # Filter out rows with no useful info
        mask = (~out["unit"].isna() | ~out["tenant_name"].isna()) & (~out["market_rent"].isna() | ~out["rent"].isna())
        out = out[mask]
        if not out.empty:
            out["_sheet"] = sheet_name
            collected.append(out)

    if not collected:
        raise ExtractionError("No usable tables found in rent roll. Adjust synonyms or provide a sample.")

    df_all = pd.concat(collected, ignore_index=True)

    # Remove obvious footer/total rows
    drop_mask = df_all["tenant_name"].fillna("").str.lower().str.contains("total|totals") | df_all["unit"].fillna("").str.lower().str.contains("total|totals")
    df_all = df_all[~drop_mask].reset_index(drop=True)

    meta = {"file": filename, "sheets_parsed": len(collected)}
    return df_all, meta

def summarize_totals(df: pd.DataFrame):
    return {
        "units_count": int(df["unit"].notna().sum()),
        "market_rent_total": float(np.nansum(df["market_rent"])),
        "rent_total": float(np.nansum(df["rent"])),
        "avg_market_rent": float(np.nanmean(df["market_rent"])),
        "avg_rent": float(np.nanmean(df["rent"]))
    }

# ---------- T12 extraction & financial calculations ----------

def _find_best_match(header, candidates):
    header = _clean_header(header)
    if not header:
        return 0
    scores = [fuzz.token_set_ratio(header, c) for c in candidates]
    return max(scores) if scores else 0

def extract_t12_financials(content: bytes, filename: str = None):
    """Extract gross income, operating expenses, and NOI from messy T12 workbooks."""
    try:
        xl = pd.ExcelFile(io.BytesIO(content))
    except Exception as e:
        raise ExtractionError(f"Cannot open T12 workbook: {e}")

    best = {"gross_income": np.nan, "total_operating_expenses": np.nan, "noi": np.nan}

    for sheet_name in xl.sheet_names:
        try:
            ws = xl.parse(sheet_name, header=None, dtype=object)
        except Exception:
            continue
        ws = ws.dropna(how="all").reset_index(drop=True).dropna(how="all", axis=1)
        if ws.empty:
            continue

        # Try to detect a two-column key-value style table OR a row-based totals table
        # Approach: scan all cells for labels matching our T12 keys; take the numeric neighbor right/below.
        for r in range(min(len(ws), 200)):
            for c in range(min(ws.shape[1], 20)):
                val = ws.iat[r, c]
                if val is None or str(val).strip() == "":
                    continue
                label = _clean_header(val)
                if not label:
                    continue

                # Check gross income
                if _find_best_match(label, T12_KEYS["gross_income"]) >= 85:
                    # neighbor to right or below
                    candidates = []
                    if c+1 < ws.shape[1]: candidates.append(ws.iat[r, c+1])
                    if r+1 < len(ws): candidates.append(ws.iat[r+1, c])
                    num = next((x for x in (_as_number(v) for v in candidates) if not pd.isna(x)), np.nan)
                    if not pd.isna(num): best["gross_income"] = num

                # Check operating expenses
                if _find_best_match(label, T12_KEYS["operating_expenses"]) >= 85:
                    candidates = []
                    if c+1 < ws.shape[1]: candidates.append(ws.iat[r, c+1])
                    if r+1 < len(ws): candidates.append(ws.iat[r+1, c])
                    num = next((x for x in (_as_number(v) for v in candidates) if not pd.isna(x)), np.nan)
                    if not pd.isna(num): best["total_operating_expenses"] = num

                # Check NOI
                if _find_best_match(label, T12_KEYS["noi"]) >= 85:
                    candidates = []
                    if c+1 < ws.shape[1]: candidates.append(ws.iat[r, c+1])
                    if r+1 < len(ws): candidates.append(ws.iat[r+1, c])
                    num = next((x for x in (_as_number(v) for v in candidates) if not pd.isna(x)), np.nan)
                    if not pd.isna(num): best["noi"] = num

    # Compute NOI if not present and we have components
    if pd.isna(best["noi"]):
        if not pd.isna(best["gross_income"]) and not pd.isna(best["total_operating_expenses"]):
            best["noi"] = float(best["gross_income"] - best["total_operating_expenses"])

    if pd.isna(best["gross_income"]) and pd.isna(best["total_operating_expenses"]) and pd.isna(best["noi"]):
        raise ExtractionError("Could not locate Gross Income, OpEx, or NOI in T12. Try another file or adjust labels.")

    return best

def compute_cap_rate(noi: float, purchase_price: float) -> float:
    if purchase_price <= 0:
        return 0.0
    return float(noi) / float(purchase_price)

def compute_price_from_cap(noi: float, cap_rate: float) -> float:
    if cap_rate <= 0:
        return 0.0
    return float(noi) / float(cap_rate)

def stats_on_rent_roll(df: pd.DataFrame):
    # Only use rows with a numeric rent
    rents = pd.to_numeric(df["rent"], errors="coerce")
    rents = rents[~pd.isna(rents)]
    if rents.empty:
        return {"mean_rent": None, "median_rent": None, "count": 0}
    return {
        "mean_rent": float(rents.mean()),
        "median_rent": float(rents.median()),
        "count": int(rents.size)
    }

def annual_gross_rent_from_rent_roll(df: pd.DataFrame):
    rents = pd.to_numeric(df["rent"], errors="coerce")
    total_monthly = float(np.nansum(rents))
    return total_monthly * 12.0

def per_unit_expense(annual_expense: float, total_units: int):
    if total_units <= 0:
        return 0.0
    return float(annual_expense) / float(total_units)

def price_per_unit(purchase_price: float, total_units: int):
    if total_units <= 0:
        return 0.0
    return float(purchase_price) / float(total_units)
