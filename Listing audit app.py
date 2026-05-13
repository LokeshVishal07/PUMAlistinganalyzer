"""
PUMA Listing Audit Analyzer  v4.0
===================================
CONFIRMED LOGIC:

SOURCE OF TRUTH (in priority order):
  1. Special Article Override  → ACTIVE/INACTIVE overrides EVERYTHING (ZeCom + inventory + stock)
  2. ZeCom Tracker             → Article-level YES/NO/OFF + Launch Date
  3. Content File              → ONLY maps Article_No (parent) → EAN (child variants)
                                 NOT used for active/inactive decisions
  4. Inventory File            → Stock source of truth per EAN per channel
  5. Marketplace File          → What is actually listed (EAN level) + MP's own stock

ACTIVE CONDITIONS (ALL must be true):
  - ZeCom Tracker = YES
  - Launch Date is in the past (or blank)
  - Special Override = ACTIVE (if uploaded, ignores all above)

INACTIVE CONDITIONS (ANY of these):
  - ZeCom Tracker = NO or OFF
  - Launch Date is in the future
  - Special Override = INACTIVE (if uploaded, ignores all above)

OUTPUT CATEGORIES (per Marketplace per Region):
  1. Active           – ZeCom says ACTIVE + EAN listed on MP + past launch
  2. Inactive         – ZeCom says INACTIVE + EAN listed on MP (should be delisted)
  3. Not Listed (Article Level) – ZeCom ACTIVE, article in content, but ZERO EANs on MP
  4. Not Listed (EAN Level)    – ZeCom ACTIVE, article has SOME variants on MP, specific EAN missing

INVENTORY LOGIC:
  - Inventory file = source of truth for stock
  - Marketplace own stock field also captured and compared
  - If inventory=0 but MP stock>0 → flag as DISCREPANCY
  - If inventory>0 but MP stock=0 → flag as DISCREPANCY
  - PH Lazada: effective_stock = inventory_stock - 1 (floor 0)
  - PH Shopee/Zalora/TikTok: use inventory directly
  - MY/SG: use inventory directly for all channels

MARKETPLACE EAN IDENTIFIERS:
  Lazada  → SellerSKU
  Zalora  → SellerSku
  Shopee  → SKU
  TikTok  → Seller sku
"""

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import warnings
warnings.filterwarnings("ignore")

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="PUMA Listing Audit Analyzer", layout="wide", page_icon="📊")

st.markdown("""
<style>
.main-header{
    background:linear-gradient(135deg,#1a1a2e 0%,#16213e 50%,#0f3460 100%);
    padding:2rem;border-radius:12px;text-align:center;margin-bottom:1.5rem;color:white;
}
.main-header h1{font-size:2rem;font-weight:700;margin:0;letter-spacing:.5px;}
.main-header p{color:#a0aec0;margin-top:.4rem;font-size:.9rem;}
.metric-box{
    background:white;border-radius:8px;padding:.9rem;
    border:1px solid #e2e8f0;text-align:center;margin-bottom:.5rem;
}
.metric-value{font-size:1.7rem;font-weight:800;}
.metric-label{font-size:.75rem;color:#718096;margin-top:.15rem;}
.cg{color:#276221;} .cr{color:#9c0006;} .co{color:#7b5800;} .cp{color:#4a235a;}
.info-box{
    background:#ebf8ff;border:1px solid #bee3f8;border-radius:6px;
    padding:.6rem .9rem;font-size:.82rem;margin:.4rem 0;
}
.warn-box{
    background:#fffbea;border:1px solid #f6e05e;border-radius:6px;
    padding:.6rem .9rem;font-size:.82rem;margin:.4rem 0;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class='main-header'>
  <h1>📊 PUMA Listing Audit Analyzer</h1>
  <p>ZeCom-Driven · Article &amp; Variant Level · MY / SG / PH · Lazada / Shopee / Zalora / TikTok</p>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════
REGIONS      = ["MY", "SG", "PH"]
MARKETPLACES = ["Lazada", "Shopee", "Zalora", "TikTok"]

# PH Lazada buffer: effective stock = raw_inventory - 1, minimum 0
# Shopee/Zalora/TikTok PH and all MY/SG: no buffer
CHANNEL_BUFFER = {
    "PH": {"Lazada": 1, "Shopee": 0, "Zalora": 0, "TikTok": 0},
    "MY": {"Lazada": 0, "Shopee": 0, "Zalora": 0, "TikTok": 0},
    "SG": {"Lazada": 0, "Shopee": 0, "Zalora": 0, "TikTok": 0},
}

# Confirmed EAN column names per marketplace
MP_EAN_COLS = {
    "Lazada" : ["SellerSKU", "Seller SKU", "seller_sku", "SellerSku"],
    "Shopee" : ["SKU", "sku", "Model SKU", "variation_sku", "Seller SKU", "SellerSKU"],
    "Zalora" : ["SellerSku", "Seller Sku", "SellerSKU", "Seller SKU", "seller_sku"],
    "TikTok" : ["Seller sku", "Seller SKU", "SellerSKU", "seller_sku", "SKU"],
}

# Output category names (also used as Excel sheet suffix)
CAT_ACTIVE   = "Active"
CAT_INACTIVE = "Inactive"
CAT_NL_ART   = "Not Listed - Article Level"
CAT_NL_EAN   = "Not Listed - EAN Level"
ALL_CATS     = [CAT_ACTIVE, CAT_INACTIVE, CAT_NL_ART, CAT_NL_EAN]

CAT_STYLE = {
    CAT_ACTIVE  : {"hdr": "#1e6f39", "bg": "#c6efce", "fc": "#276221", "icon": "🟢"},
    CAT_INACTIVE: {"hdr": "#9c0006", "bg": "#ffc7ce", "fc": "#9c0006", "icon": "🔴"},
    CAT_NL_ART  : {"hdr": "#7b5800", "bg": "#ffeb9c", "fc": "#7b5800", "icon": "🟡"},
    CAT_NL_EAN  : {"hdr": "#4a235a", "bg": "#e8d5f5", "fc": "#4a235a", "icon": "🟣"},
}

# ══════════════════════════════════════════════════════════════════════════════
# BULLET-PROOF COLUMN HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def _sc(c) -> str:
    """Any column label (int/float/str/NaN) → clean stripped string."""
    try:
        s = str(c).strip()
        return "" if s in ("nan", "None", "NaT", "") else s
    except Exception:
        return ""

def _cl(c) -> str:
    return _sc(c).lower()

def sanitize_cols(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [_sc(c) or f"_col_{i}" for i, c in enumerate(df.columns)]
    return df

def find_header_row(file, max_scan: int = 25) -> int:
    """Find the row with the most non-empty string cells → that is the header."""
    try:
        prev = pd.read_excel(file, header=None, nrows=max_scan)
        file.seek(0)
        best_r, best_s = 0, -1
        for i, row in prev.iterrows():
            sc = int(row.dropna().apply(
                lambda v: isinstance(v, str) and len(v.strip()) > 0
            ).sum())
            if sc > best_s:
                best_s, best_r = sc, int(i)
        return best_r
    except Exception:
        try:
            file.seek(0)
        except Exception:
            pass
        return 0

def safe_read(file) -> pd.DataFrame:
    """Read Excel/CSV, auto-detect header, always return string-column DataFrame."""
    try:
        hdr = find_header_row(file)
        df  = pd.read_excel(file, header=hdr)
    except Exception:
        try:
            file.seek(0)
            df = pd.read_excel(file, header=0)
        except Exception:
            try:
                file.seek(0)
                df = pd.read_csv(file, header=0)
            except Exception as e:
                st.warning(f"Cannot read file: {e}")
                return pd.DataFrame()
    df = sanitize_cols(df)
    df = df.dropna(how="all").reset_index(drop=True)
    return df

def norm_col(df: pd.DataFrame, candidates: list, new_name: str) -> pd.DataFrame:
    """Rename first matching column to new_name (case-insensitive)."""
    cands = [str(x).strip().lower() for x in candidates]
    for col in df.columns:
        if _cl(col) in cands:
            if col != new_name:
                df = df.rename(columns={col: new_name})
            break
    return df

def cs(s: pd.Series) -> pd.Series:
    """Clean string series: strip, upper, replace NaN/None with empty string."""
    return s.astype(str).str.strip().str.upper().replace(
        {"NAN": "", "NONE": "", "NAT": "", "NA": ""}
    )

def to_num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0).clip(lower=0).astype(int)

def first_col_with(df: pd.DataFrame, keywords: list, exclude: set = None) -> str | None:
    """Return first column name containing any keyword."""
    excl = exclude or set()
    for kw in keywords:
        for c in df.columns:
            if kw in _cl(c) and c not in excl:
                return c
    return None

# ══════════════════════════════════════════════════════════════════════════════
# FILE LOADERS
# ══════════════════════════════════════════════════════════════════════════════

# ── 1. ZeCom Tracker ──────────────────────────────────────────────────────────
def load_zecom(file) -> pd.DataFrame:
    """
    Returns one row per Article_No.
    Columns: Article_No, Tracker_Lazada, Tracker_Shopee, Tracker_Zalora,
             Tracker_TikTok, Launch_Date
    Tracker raw values: YES / NO / OFF / blank
    """
    df = safe_read(file)

    df = norm_col(df, [
        "PIM Article#", "PIM Article", "Article No", "ArticleNo",
        "Color_No", "ColorNo", "Color No", "Article Number",
        "Style Number", "StyleNo", "Style No", "Parent SKU",
    ], "Article_No")

    if "Article_No" not in df.columns:
        st.error(
            "❌ ZeCom Tracker: Cannot find Article No column. "
            "Expected one of: PIM Article#, Color_No, Article No, Style Number"
        )
        return pd.DataFrame(columns=["Article_No"])

    # Tracker columns — find by marketplace name in column header
    reserved = {"Article_No"}
    for mp in MARKETPLACES:
        already = {f"Tracker_{m}" for m in MARKETPLACES}
        cols = [
            c for c in df.columns
            if mp.lower() in _cl(c)
            and c not in already
            and c not in reserved
        ]
        if cols:
            df = df.rename(columns={cols[0]: f"Tracker_{mp}"})

    # Guarantee all tracker columns exist (blank = UNKNOWN)
    for mp in MARKETPLACES:
        if f"Tracker_{mp}" not in df.columns:
            df[f"Tracker_{mp}"] = np.nan

    # Launch date
    skip = reserved | {f"Tracker_{m}" for m in MARKETPLACES}
    date_cols = [
        c for c in df.columns
        if ("launch" in _cl(c) or "go live" in _cl(c) or "golive" in _cl(c)
            or "live date" in _cl(c))
        and c not in skip
    ]
    if date_cols:
        df = df.rename(columns={date_cols[0]: "Launch_Date"})
        df["Launch_Date"] = pd.to_datetime(df["Launch_Date"], errors="coerce")
    else:
        df["Launch_Date"] = pd.NaT

    df["Article_No"] = cs(df["Article_No"])
    df = df[df["Article_No"].str.len() > 0]
    df = df.drop_duplicates(subset=["Article_No"]).reset_index(drop=True)
    return df


# ── 2. Special Article Override ────────────────────────────────────────────────
def load_override(file) -> dict:
    """
    Returns {Article_No → 'ACTIVE' | 'INACTIVE'}.
    Overrides ZeCom + inventory for ALL regions.
    Expected columns: Article_No (or any alias), Status (ACTIVE / INACTIVE).
    """
    df = safe_read(file)
    df = norm_col(df, [
        "PIM Article#", "PIM Article", "Article No", "ArticleNo",
        "Color_No", "ColorNo", "Article Number", "Style Number",
    ], "Article_No")
    df = norm_col(df, [
        "Status", "status", "Override Status", "Listing Status",
        "Active", "Flag",
    ], "Status")

    if "Article_No" not in df.columns or "Status" not in df.columns:
        st.warning("⚠️ Special Override: needs Article_No and Status columns.")
        return {}

    df["Article_No"] = cs(df["Article_No"])
    df["Status"]     = cs(df["Status"])
    df = df[df["Article_No"].str.len() > 0]
    df = df[df["Status"].isin(["ACTIVE", "INACTIVE"])]
    return dict(zip(df["Article_No"], df["Status"]))


# ── 3. Content File ────────────────────────────────────────────────────────────
def load_content(file) -> pd.DataFrame:
    """
    ONLY USED FOR: Article_No → EAN variant mapping.
    NOT used for active/inactive decisions.
    Returns columns: EAN, Article_No [, Size, Color, Description if found]
    """
    df = safe_read(file)

    df = norm_col(df, [
        "EAN", "ean", "Barcode", "barcode", "Child SKU", "ChildSKU",
        "Seller SKU", "SellerSKU", "Item EAN", "GTIN", "UPC",
    ], "EAN")

    df = norm_col(df, [
        "Color_No", "ColorNo", "Article No", "ArticleNo",
        "PIM Article#", "Article Number", "Parent SKU",
        "Color No", "Style Number", "StyleNo",
    ], "Article_No")

    # Capture optional size/color/description for richer variant output
    extra = []
    for label, keys in [
        ("Size",        ["size", "sz"]),
        ("Color",       ["colour", "color"]),
        ("Description", ["desc", "product name", "name"]),
    ]:
        col = first_col_with(df, keys, exclude={"EAN", "Article_No"})
        if col:
            df = df.rename(columns={col: label})
            extra.append(label)

    for col in ["EAN", "Article_No"]:
        if col not in df.columns:
            df[col] = np.nan

    df["EAN"]        = cs(df["EAN"])
    df["Article_No"] = cs(df["Article_No"])
    df = df[(df["EAN"].str.len() > 0) & (df["Article_No"].str.len() > 0)]

    keep = ["EAN", "Article_No"] + [e for e in extra if e in df.columns]
    return df[keep].drop_duplicates(subset=["EAN"]).reset_index(drop=True)


# ── 4. Inventory File ─────────────────────────────────────────────────────────
def load_inventory(file, region: str) -> tuple:
    """
    Returns (DataFrame, debug_dict).
    DataFrame columns: EAN, Inv_Lazada, Inv_Shopee, Inv_Zalora, Inv_TikTok
    - Tries channel-specific columns first (containing marketplace name)
    - Falls back to single total/qty/soh/available column
    - Applies PH Lazada buffer (-1, floor 0)
    """
    df = safe_read(file)
    all_cols = list(df.columns)

    df = norm_col(df, [
        "EAN", "ean", "Barcode", "barcode", "Item EAN", "GTIN",
        "SKU", "sku", "Item Code", "Seller SKU", "SellerSKU",
        "Material", "Material Number",
    ], "EAN")

    if "EAN" not in df.columns:
        st.warning(
            f"[{region}] Inventory: Cannot find EAN/Barcode column.\n"
            f"Columns in file: {all_cols[:25]}"
        )
        empty = pd.DataFrame(columns=["EAN","Inv_Lazada","Inv_Shopee","Inv_Zalora","Inv_TikTok"])
        return empty, {"Error": "EAN column not found", "All columns": str(all_cols[:30])}

    df["EAN"] = cs(df["EAN"])
    df = df[df["EAN"].str.len() > 0]

    excl = {"EAN"}

    # ── PRIMARY: Avail_Qty is the confirmed source-of-truth stock column ──────
    # Check exact match first, then case-insensitive, then channel-specific,
    # then generic fallback — in that order.
    avail_qty_exact = "Avail_Qty" if "Avail_Qty" in df.columns else None
    avail_qty_ci    = next(
        (c for c in df.columns if _cl(c) == "avail_qty"), None
    ) if not avail_qty_exact else avail_qty_exact

    # Primary stock column = Avail_Qty (any casing)
    primary_c = avail_qty_exact or avail_qty_ci

    # ── SECONDARY: channel-specific columns (used if present AND Avail_Qty missing) ──
    laz_c = first_col_with(df, ["lazada"],             excl) if not primary_c else None
    sho_c = first_col_with(df, ["shopee"],             excl) if not primary_c else None
    zal_c = first_col_with(df, ["zalora"],             excl) if not primary_c else None
    ttk_c = first_col_with(df, ["tiktok", "tik tok"],  excl) if not primary_c else None

    # ── TERTIARY: generic total/qty/soh fallback ──────────────────────────────
    used = {primary_c, laz_c, sho_c, zal_c, ttk_c} - {None}
    tot_c = None if primary_c else first_col_with(df, [
        "avail_qty", "available", "on hand", "onhand", "total", "qty",
        "quantity", "stock", "soh", "free", "unrestricted",
    ], excl | used)

    # Final resolved column for each channel
    def resolve(ch_col):
        """Return the best available stock column for a channel."""
        if primary_c and primary_c in df.columns:
            return primary_c          # Avail_Qty wins for every channel
        if ch_col and ch_col in df.columns:
            return ch_col             # channel-specific
        if tot_c and tot_c in df.columns:
            return tot_c              # generic fallback
        return None

    laz_src = resolve(laz_c)
    sho_src = resolve(sho_c)
    zal_src = resolve(zal_c)
    ttk_src = resolve(ttk_c)

    debug = {
        "PRIMARY stock col (Avail_Qty)" : primary_c or "NOT FOUND",
        "Lazada  → using col"           : laz_src or "(none — stock = 0)",
        "Shopee  → using col"           : sho_src or "(none — stock = 0)",
        "Zalora  → using col"           : zal_src or "(none — stock = 0)",
        "TikTok  → using col"           : ttk_src or "(none — stock = 0)",
        "Generic fallback col"          : tot_c or "(not needed / not found)",
        "EAN rows loaded"               : len(df),
        "All columns"                   : ", ".join(all_cols[:50]),
    }

    if not primary_c:
        st.warning(
            f"[{region}] Inventory: 'Avail_Qty' column NOT found. "
            f"Columns in file: {all_cols[:30]}. "
            f"Using fallback: '{tot_c or 'none (stock=0)'}'"
        )

    def get_stock(src_col):
        if src_col and src_col in df.columns:
            return to_num(df[src_col])
        return pd.Series(0, index=df.index)

    result = pd.DataFrame({
        "EAN"        : df["EAN"],
        "Inv_Lazada" : get_stock(laz_src),
        "Inv_Shopee" : get_stock(sho_src),
        "Inv_Zalora" : get_stock(zal_src),
        "Inv_TikTok" : get_stock(ttk_src),
    })

    # Apply channel buffer
    buf = CHANNEL_BUFFER.get(region, {})
    for mp, col in [("Lazada","Inv_Lazada"),("Shopee","Inv_Shopee"),
                    ("Zalora","Inv_Zalora"),("TikTok","Inv_TikTok")]:
        b = buf.get(mp, 0)
        if b > 0:
            result[col] = (result[col] - b).clip(lower=0)

    result = result.drop_duplicates(subset=["EAN"]).reset_index(drop=True)
    return result, debug


# ── 5. Marketplace Loaders ────────────────────────────────────────────────────
def _load_mp(df: pd.DataFrame, mp: str) -> pd.DataFrame:
    """
    Generic MP loader using confirmed EAN column names per marketplace.
    Returns: EAN, MP_Status, MP_Stock, Marketplace
    """
    ean_cols    = MP_EAN_COLS[mp]
    status_cols = ["Status", "status", "ItemStatus", "Item Status",
                   "Listing Status", "Product Status", "Active", "Publish"]
    stock_cols  = ["Available", "Stock", "Quantity", "available",
                   "quantity", "Qty", "FreeQty", "sellable_quantity",
                   "Current Stock", "Available Stock", "free qty"]

    df = norm_col(df, ean_cols,    "EAN")
    df = norm_col(df, status_cols, "MP_Status")
    df = norm_col(df, stock_cols,  "MP_Stock")

    for col in ["EAN", "MP_Status", "MP_Stock"]:
        if col not in df.columns:
            df[col] = np.nan

    df["EAN"]        = cs(df["EAN"])
    df["MP_Stock"]   = to_num(df["MP_Stock"])
    df["Marketplace"] = mp
    df = df[df["EAN"].str.len() > 0]
    return df[["EAN", "MP_Status", "MP_Stock", "Marketplace"]].drop_duplicates(subset=["EAN"])

def load_lazada(file) -> pd.DataFrame:
    return _load_mp(safe_read(file), "Lazada")

def load_shopee(file) -> pd.DataFrame:
    return _load_mp(safe_read(file), "Shopee")

def load_tiktok(file) -> pd.DataFrame:
    return _load_mp(safe_read(file), "TikTok")

def load_zalora(status_file, stock_file) -> pd.DataFrame:
    """
    Zalora comes as two separate files:
    SellerStatusTemplate → status
    SellerStockTemplate  → stock
    Merge on EAN (SellerSku).
    """
    ds   = safe_read(status_file)
    dstk = safe_read(stock_file)

    ds   = norm_col(ds,   MP_EAN_COLS["Zalora"], "EAN")
    ds   = norm_col(ds,   ["Status","status","ItemStatus","Item Status","Active"], "MP_Status")
    dstk = norm_col(dstk, MP_EAN_COLS["Zalora"], "EAN")
    dstk = norm_col(dstk, ["Stock","Quantity","Available","Qty","quantity"], "MP_Stock")

    for col in ["EAN", "MP_Status"]:
        if col not in ds.columns:
            ds[col] = np.nan
    if "EAN"      not in dstk.columns: dstk["EAN"]      = np.nan
    if "MP_Stock" not in dstk.columns: dstk["MP_Stock"]  = np.nan

    ds["EAN"]   = cs(ds["EAN"])
    dstk["EAN"] = cs(dstk["EAN"])

    merged = ds.merge(
        dstk[["EAN", "MP_Stock"]].drop_duplicates("EAN"),
        on="EAN", how="left"
    )
    merged["MP_Stock"]   = to_num(merged["MP_Stock"])
    merged["Marketplace"] = "Zalora"
    merged = merged[merged["EAN"].str.len() > 0]
    return merged[["EAN","MP_Status","MP_Stock","Marketplace"]].drop_duplicates(subset=["EAN"])


# ══════════════════════════════════════════════════════════════════════════════
# STATUS DECISION ENGINE
# ══════════════════════════════════════════════════════════════════════════════
def resolve_tracker(val) -> str:
    """Raw tracker cell → 'ACTIVE' | 'INACTIVE' | 'UNKNOWN'"""
    if pd.isna(val):
        return "UNKNOWN"
    v = str(val).strip().upper()
    if v == "YES":
        return "ACTIVE"
    if v in ("NO", "OFF"):
        return "INACTIVE"
    return "UNKNOWN"

def is_mp_active(val) -> bool:
    """Marketplace status cell → True if the listing is active."""
    if pd.isna(val):
        return False
    return str(val).strip().upper() in {
        "ACTIVE", "1", "TRUE", "YES", "ENABLED", "PUBLISHED",
        "LISTED", "NORMAL", "ACTIVATED", "FOR SALE",
    }

def decide_expected_status(
    zecom_row: pd.Series,
    mp: str,
    override_map: dict,
) -> tuple:
    """
    Returns (expected_status: str, reasons: list[str])
    expected_status = 'ACTIVE' | 'INACTIVE'

    PRIORITY ORDER:
      1. Special Override → overrides EVERYTHING
      2. ZeCom Tracker (NO/OFF → INACTIVE)
      3. Launch Date (future → INACTIVE)
      4. ZeCom Tracker YES → ACTIVE
      5. Blank tracker → INACTIVE (unknown = not confirmed active)
    """
    art     = str(zecom_row.get("Article_No", "")).strip()
    reasons = []

    # PRIORITY 1: Special Override
    if art in override_map:
        status = override_map[art]
        reasons.append(f"Special Article Override → {status}")
        return status, reasons

    # PRIORITY 2: ZeCom Tracker
    t_raw    = zecom_row.get(f"Tracker_{mp}", np.nan)
    t_status = resolve_tracker(t_raw)
    t_raw_str = str(t_raw).strip() if pd.notna(t_raw) else "blank"

    if t_status == "INACTIVE":
        reasons.append(f"ZeCom Tracker = {t_raw_str} (NO/OFF → INACTIVE)")
        return "INACTIVE", reasons

    # PRIORITY 3: Launch Date
    launch = zecom_row.get("Launch_Date", pd.NaT)
    today  = pd.Timestamp.today().normalize()
    if pd.notna(launch):
        try:
            lts = pd.Timestamp(launch).normalize()
            if lts > today:
                reasons.append(f"Future Launch Date: {lts.date()} (today: {today.date()})")
                return "INACTIVE", reasons
            else:
                reasons.append(f"Launch Date: {lts.date()} ✓ (past)")
        except Exception:
            pass

    # PRIORITY 4: YES → ACTIVE
    if t_status == "ACTIVE":
        reasons.append(f"ZeCom Tracker = YES → ACTIVE")
        return "ACTIVE", reasons

    # PRIORITY 5: blank/unknown → INACTIVE (not confirmed)
    reasons.append(f"ZeCom Tracker = {t_raw_str} (blank/unknown → treated as INACTIVE)")
    return "INACTIVE", reasons


# ══════════════════════════════════════════════════════════════════════════════
# STOCK DISCREPANCY CHECK
# ══════════════════════════════════════════════════════════════════════════════
INV_COL = {
    "Lazada": "Inv_Lazada",
    "Shopee": "Inv_Shopee",
    "Zalora": "Inv_Zalora",
    "TikTok": "Inv_TikTok",
}

def get_inv_stock(ean: str, mp: str, inv_idx: dict) -> int:
    row = inv_idx.get(ean)
    if row is None:
        return 0
    col = INV_COL.get(mp, "Inv_Lazada")
    v   = row.get(col, 0)
    r   = pd.to_numeric(v, errors="coerce")
    return int(r) if pd.notna(r) else 0

def stock_discrepancy_note(inv_stock: int, mp_stock: int) -> str:
    """
    Compare inventory file stock vs marketplace reported stock.
    Returns a discrepancy note or '-'.
    """
    if inv_stock > 0 and mp_stock == 0:
        return f"⚠ DISCREPANCY: Inventory has {inv_stock} units but MP shows 0 stock"
    if inv_stock == 0 and mp_stock > 0:
        return f"⚠ DISCREPANCY: Inventory is 0 but MP shows {mp_stock} units"
    return "-"


# ══════════════════════════════════════════════════════════════════════════════
# AUDIT ENGINE — ZeCom Article-Driven
# ══════════════════════════════════════════════════════════════════════════════
def run_audit(
    mp_dfs       : dict,            # {mp: DataFrame[EAN, MP_Status, MP_Stock]}
    inv_df       : pd.DataFrame,    # EAN-level inventory stocks
    zecom_df     : pd.DataFrame,    # Article_No-level tracker
    content_df   : pd.DataFrame,    # EAN → Article_No mapping
    override_map : dict,            # Article_No → ACTIVE/INACTIVE
    region       : str,
) -> dict:
    """
    DRIVEN BY ZECOM: iterate every Article_No in ZeCom, not by MP listings.

    Returns:
      {mp: {cat: DataFrame}}
    where cat ∈ ALL_CATS
    """
    today = pd.Timestamp.today().normalize()

    # ── Build lookup structures ───────────────────────────────────────────────

    # Article_No → sorted list of EANs (from content)
    art_to_eans: dict[str, list] = {}
    for _, row in content_df.iterrows():
        art = str(row["Article_No"]).strip()
        ean = str(row["EAN"]).strip()
        if art and ean:
            art_to_eans.setdefault(art, []).append(ean)

    # EAN → metadata dict (Size, Color, Description from content)
    extra_cols = [c for c in content_df.columns if c not in ("EAN", "Article_No")]
    ean_meta: dict[str, dict] = {}
    for _, row in content_df.iterrows():
        ean = str(row["EAN"]).strip()
        if ean:
            ean_meta[ean] = {c: row[c] for c in extra_cols if c in row.index}

    # Inventory index: EAN → Series
    inv_idx: dict[str, pd.Series] = {}
    if not inv_df.empty:
        for _, row in inv_df.iterrows():
            inv_idx[str(row["EAN"]).strip()] = row

    results = {}

    for mp in MARKETPLACES:
        mp_df = mp_dfs.get(mp, pd.DataFrame())

        # EAN index for this marketplace
        mp_idx: dict[str, pd.Series] = {}
        if not mp_df.empty:
            for _, row in mp_df.iterrows():
                e = str(row["EAN"]).strip()
                if e:
                    mp_idx[e] = row

        active_rows    = []
        inactive_rows  = []
        nl_art_rows    = []
        nl_ean_rows    = []

        # ── Iterate every Article_No in ZeCom ────────────────────────────────
        for _, z_row in zecom_df.iterrows():
            art = str(z_row.get("Article_No", "")).strip()
            if not art:
                continue

            exp_status, exp_reasons = decide_expected_status(z_row, mp, override_map)

            # All EANs for this article (from content)
            variants = art_to_eans.get(art, [])

            # EANs actually listed on this MP
            listed_eans   = [e for e in variants if e in mp_idx]
            unlisted_eans = [e for e in variants if e not in mp_idx]

            # ── CASE A: Article has no variants in Content ────────────────────
            if not variants:
                base_row = {
                    "Region"             : region,
                    "Marketplace"        : mp,
                    "Article No"         : art,
                    "EAN (Seller SKU)"   : "-",
                    "Expected Status"    : exp_status,
                    "MP Listed"          : "NO",
                    "MP Status"          : "-",
                    "MP Stock"           : 0,
                    "Inventory Stock"    : 0,
                    "Stock Discrepancy"  : "-",
                    "Reason"             : "; ".join(exp_reasons),
                    "Note"               : "Article has no EAN variants in Content File",
                }
                if exp_status == "INACTIVE":
                    inactive_rows.append(base_row)
                else:
                    nl_art_rows.append({**base_row,
                        "Note": "ACTIVE per ZeCom but no EAN variants in Content File"})
                continue

            # ── CASE B: INACTIVE articles ─────────────────────────────────────
            if exp_status == "INACTIVE":
                for ean in variants:
                    mp_row     = mp_idx.get(ean)
                    inv_stock  = get_inv_stock(ean, mp, inv_idx)
                    mp_stock   = int(mp_row["MP_Stock"]) if mp_row is not None else 0
                    meta       = ean_meta.get(ean, {})
                    inactive_rows.append({
                        "Region"            : region,
                        "Marketplace"       : mp,
                        "Article No"        : art,
                        "EAN (Seller SKU)"  : ean,
                        **meta,
                        "Expected Status"   : "INACTIVE",
                        "MP Listed"         : "YES" if mp_row is not None else "NO",
                        "MP Status"         : str(mp_row["MP_Status"]) if mp_row is not None else "Not Listed",
                        "MP Stock"          : mp_stock,
                        "Inventory Stock"   : inv_stock,
                        "Stock Discrepancy" : stock_discrepancy_note(inv_stock, mp_stock),
                        "Reason"            : "; ".join(exp_reasons),
                        "Action Required"   : "Delist from MP" if mp_row is not None else "-",
                    })
                continue

            # ── CASE C: ACTIVE articles ───────────────────────────────────────
            n_total  = len(variants)
            n_listed = len(listed_eans)

            # Sub-case C1: Zero variants on MP → Not Listed at Article Level
            if n_listed == 0:
                for ean in variants:
                    inv_stock = get_inv_stock(ean, mp, inv_idx)
                    meta      = ean_meta.get(ean, {})
                    nl_art_rows.append({
                        "Region"            : region,
                        "Marketplace"       : mp,
                        "Article No"        : art,
                        "EAN (Seller SKU)"  : ean,
                        **meta,
                        "Expected Status"   : "ACTIVE",
                        "MP Listed"         : "NO",
                        "MP Status"         : "Not Listed",
                        "MP Stock"          : 0,
                        "Inventory Stock"   : inv_stock,
                        "Stock Discrepancy" : "-",
                        "Reason"            : "; ".join(exp_reasons),
                        "Action Required"   : "List this article on MP",
                    })
                continue

            # Sub-case C2: Some variants listed, process each EAN
            for ean in variants:
                mp_row    = mp_idx.get(ean)
                inv_stock = get_inv_stock(ean, mp, inv_idx)
                meta      = ean_meta.get(ean, {})

                if mp_row is not None:
                    # EAN IS on marketplace
                    mp_active = is_mp_active(mp_row["MP_Status"])
                    mp_stock  = int(mp_row["MP_Stock"])
                    disc_note = stock_discrepancy_note(inv_stock, mp_stock)

                    rec = {
                        "Region"            : region,
                        "Marketplace"       : mp,
                        "Article No"        : art,
                        "EAN (Seller SKU)"  : ean,
                        **meta,
                        "Expected Status"   : "ACTIVE",
                        "MP Listed"         : "YES",
                        "MP Status"         : str(mp_row["MP_Status"]),
                        "MP Stock"          : mp_stock,
                        "Inventory Stock"   : inv_stock,
                        "Stock Discrepancy" : disc_note,
                        "Reason"            : "; ".join(exp_reasons),
                        "Action Required"   : "-",
                    }

                    if mp_active:
                        if disc_note != "-":
                            rec["Action Required"] = "Investigate stock discrepancy"
                        active_rows.append(rec)
                    else:
                        # Listed INACTIVE on MP despite ZeCom ACTIVE
                        rec["Reason"] = ("; ".join(exp_reasons)
                                         + "; Listed as INACTIVE on MP despite ZeCom = ACTIVE")
                        rec["Action Required"] = "Activate listing on MP"
                        # Treat as active (correct status would be active)
                        # but put in inactive tab because MP shows inactive
                        inactive_rows.append(rec)

                else:
                    # EAN NOT on marketplace → variant/size missing
                    nl_ean_rows.append({
                        "Region"            : region,
                        "Marketplace"       : mp,
                        "Article No"        : art,
                        "EAN (Seller SKU)"  : ean,
                        **meta,
                        "Expected Status"   : "ACTIVE",
                        "MP Listed"         : "NO",
                        "MP Status"         : "Not Listed",
                        "MP Stock"          : 0,
                        "Inventory Stock"   : inv_stock,
                        "Stock Discrepancy" : "-",
                        "Reason"            : ("; ".join(exp_reasons)
                                               + f"; {n_listed}/{n_total} variants listed on MP"
                                               + f" — this EAN/size is missing"),
                        "Action Required"   : "Add missing variant/size to MP listing",
                    })

        results[mp] = {
            CAT_ACTIVE  : pd.DataFrame(active_rows),
            CAT_INACTIVE: pd.DataFrame(inactive_rows),
            CAT_NL_ART  : pd.DataFrame(nl_art_rows),
            CAT_NL_EAN  : pd.DataFrame(nl_ean_rows),
        }

    return results


# ══════════════════════════════════════════════════════════════════════════════
# EXCEL EXPORT
# ══════════════════════════════════════════════════════════════════════════════
def build_excel(all_results: dict) -> bytes:
    """
    Sheet layout:
      📋 Summary          — region × marketplace × category counts
      {MP} {Region} Active
      {MP} {Region} Inactive
      {MP} {Region} Not Listed - Article Level
      {MP} {Region} Not Listed - EAN Level
    """
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb = writer.book

        def fmt(bold=False, bg=None, fc="#000000", border=1,
                align="left", wrap=False, italic=False, sz=9):
            d = {"font_name": "Arial", "font_size": sz, "font_color": fc,
                 "border": border, "align": align, "valign": "vcenter",
                 "bold": bold, "italic": italic, "text_wrap": wrap}
            if bg:
                d["bg_color"] = bg
            return wb.add_format(d)

        ttl  = fmt(bold=True, sz=13, fc="#0f3460", border=0)
        sub  = fmt(italic=True, sz=8, fc="#718096", border=0)
        norm = fmt()

        cat_hdr_fmts  = {c: fmt(bold=True, bg=CAT_STYLE[c]["hdr"],
                                 fc="#ffffff", align="center") for c in ALL_CATS}
        cat_data_fmts = {c: fmt(bg=CAT_STYLE[c]["bg"],
                                 fc=CAT_STYLE[c]["fc"]) for c in ALL_CATS}
        col_hdr_fmts  = {c: fmt(bold=True, bg=CAT_STYLE[c]["hdr"],
                                 fc="#ffffff", align="center", wrap=True) for c in ALL_CATS}

        # ── Summary Sheet ─────────────────────────────────────────────────────
        ws  = wb.add_worksheet("Summary")
        writer.sheets["Summary"] = ws
        ws.set_zoom(90)
        ws.set_column("A:A", 10)
        ws.set_column("B:B", 14)
        ws.set_column("C:F", 28)
        ws.write("A1", "PUMA Listing Audit — Summary", ttl)
        ws.write("A2", f"Generated: {datetime.now().strftime('%d %b %Y  %H:%M')}", sub)

        r = 4
        # Header
        ws.write(r, 0, "Region",     fmt(bold=True, bg="#0f3460", fc="#ffffff", align="center"))
        ws.write(r, 1, "Marketplace",fmt(bold=True, bg="#0f3460", fc="#ffffff", align="center"))
        for ci, cat in enumerate(ALL_CATS):
            ws.write(r, ci + 2, f"{CAT_STYLE[cat]['icon']} {cat}",
                     fmt(bold=True, bg=CAT_STYLE[cat]["hdr"], fc="#ffffff", align="center", wrap=True))
        ws.set_row(r, 35)
        r += 1

        for region, mp_res in all_results.items():
            for mp in MARKETPLACES:
                if mp not in mp_res:
                    continue
                ws.write(r, 0, region,     norm)
                ws.write(r, 1, mp,         norm)
                for ci, cat in enumerate(ALL_CATS):
                    cnt = len(mp_res[mp].get(cat, pd.DataFrame()))
                    ws.write(r, ci + 2, cnt, cat_data_fmts[cat])
                r += 1

        # ── Detail Sheets ─────────────────────────────────────────────────────
        for region, mp_res in all_results.items():
            for mp in MARKETPLACES:
                if mp not in mp_res:
                    continue
                for cat in ALL_CATS:
                    df = mp_res[mp].get(cat, pd.DataFrame())

                    # Sheet name: max 31 chars
                    raw_name = f"{mp[:6]} {region} {cat[:20]}"
                    sname    = raw_name[:31].strip()

                    ws2 = wb.add_worksheet(sname)
                    writer.sheets[sname] = ws2
                    ws2.set_zoom(85)

                    ws2.write(0, 0,
                              f"{CAT_STYLE[cat]['icon']} {mp} [{region}] — {cat}",
                              ttl)
                    ws2.write(1, 0, f"Total records: {len(df)}", sub)

                    if df.empty:
                        ws2.write(2, 0, "No records in this category.", sub)
                        continue

                    hf = col_hdr_fmts[cat]
                    df = df.reset_index(drop=True)

                    for ci, col in enumerate(df.columns):
                        ws2.write(2, ci, col, hf)
                        # Auto-width estimate
                        max_w = max(
                            len(str(col)),
                            df[col].astype(str).str.len().max() if len(df) > 0 else 10
                        )
                        ws2.set_column(ci, ci, min(max_w + 2, 45))

                    # Highlight columns
                    highlight_cols = {"Expected Status", "Reason", "Action Required",
                                      "MP Status", "Stock Discrepancy"}

                    for ri, (_, rec) in enumerate(df.iterrows()):
                        for ci, col in enumerate(df.columns):
                            val = rec[col]
                            safe_val = ("" if (isinstance(val, float) and np.isnan(val))
                                        else str(val))
                            if col in highlight_cols:
                                ws2.write(ri + 3, ci, safe_val, cat_data_fmts[cat])
                            else:
                                ws2.write(ri + 3, ci, safe_val, norm)

    return output.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════════════════════════
for key in ("audit_results", "inv_debug"):
    if key not in st.session_state:
        st.session_state[key] = {}

# ══════════════════════════════════════════════════════════════════════════════
# UI TABS
# ══════════════════════════════════════════════════════════════════════════════
tab_upload, tab_results, tab_debug, tab_help = st.tabs([
    "📁 Upload & Run", "📊 Results & Download", "🔧 Inventory Debug", "❓ Help",
])

# ════════════════════ HELP ════════════════════════════════════════════════════
with tab_help:
    st.markdown("""
## 📖 File Reference

| File | Expected Filename Pattern | Required? |
|---|---|---|
| ZeCom Tracker | any name (contains zecom/tracker) | ✅ Yes |
| Content Master | any name (EAN ↔ Article No mapping) | ✅ Yes |
| Special Override | any name (Article_No + Status columns) | Optional |
| Lazada | starts with `pricestock` | Per region |
| Shopee | contains `Shopee` + `Masterfile` | Per region |
| Zalora Status | starts with `SellerStatusTemplate` | Per region |
| Zalora Stock | starts with `SellerStockTemplate` | Per region |
| TikTok | starts with `TikTokSellerCenter` or has `batchedit` | Per region |
| Inventory PH | starts with `Inventory_` | Per region |
| Inventory MY | starts with `PUMA_MY_B2C_Channel_Inventory_` | Per region |
| Inventory SG | starts with `SG_PUMA SG B2C Inventory Rpt_New_` | Per region |

---
## ✅ Status Decision Logic (in priority order)

```
PRIORITY 1 — Special Article Override
   → If Article No is in override file: use that status (ACTIVE/INACTIVE)
   → Overrides ZeCom, inventory, everything

PRIORITY 2 — ZeCom Tracker
   → YES  = potentially ACTIVE
   → NO   = INACTIVE
   → OFF  = INACTIVE
   → blank/unknown = INACTIVE (not confirmed active)

PRIORITY 3 — Launch Date
   → Future date = INACTIVE (not launched yet)
   → Past date or blank = does not block ACTIVE

Result: ACTIVE only if ZeCom=YES AND launch past/blank AND no override to INACTIVE
```

---
## 📋 Output Categories

| Category | Meaning | Action |
|---|---|---|
| 🟢 **Active** | ZeCom ACTIVE, listed on MP, MP status = active | Monitor stock discrepancies |
| 🔴 **Inactive** | ZeCom INACTIVE or MP shows inactive despite ZeCom ACTIVE | Delist / investigate |
| 🟡 **Not Listed - Article Level** | ZeCom ACTIVE but ZERO variants exist on MP | List entire article |
| 🟣 **Not Listed - EAN Level** | ZeCom ACTIVE, article partially listed, but specific size/EAN missing | Add missing variant |

---
## 🔑 EAN Identifiers Per Marketplace

| Marketplace | EAN Column |
|---|---|
| Lazada | `SellerSKU` |
| Shopee | `SKU` |
| Zalora | `SellerSku` |
| TikTok | `Seller sku` |

## 📦 Inventory Buffer
- **PH Lazada only**: effective stock = inventory − 1 (minimum 0)
- All other channels / regions: use inventory directly
- Stock discrepancy flagged when inventory ≠ MP stock
    """)

# ════════════════════ UPLOAD & RUN ═══════════════════════════════════════════
with tab_upload:

    st.markdown("### Step 1 — Select Region(s) to Audit")
    selected_regions = st.multiselect(
        "Regions:", REGIONS, default=REGIONS,
        help="Choose one or more regions. Each region needs its own marketplace + inventory files."
    )

    st.markdown("### Step 2 — Upload Master Files *(apply to all regions)*")
    mc1, mc2, mc3 = st.columns(3)
    with mc1:
        zecom_file = st.file_uploader(
            "📋 ZeCom Tracker  **[required]**",
            type=["xlsx","xls","csv"], key="zecom",
            help="Columns: Article No | Lazada (YES/NO/OFF) | Shopee | Zalora | TikTok | Launch Date"
        )
    with mc2:
        content_file = st.file_uploader(
            "📦 Content Master  **[required]**",
            type=["xlsx","xls","csv"], key="content",
            help="Columns: EAN (child/variant) | Article No (parent/style) [+ Size, Color optional]"
        )
    with mc3:
        override_file = st.file_uploader(
            "⚡ Special Article Override  *(optional)*",
            type=["xlsx","xls","csv"], key="override",
            help="Columns: Article_No | Status (ACTIVE / INACTIVE) — overrides ZeCom + inventory for all regions"
        )

    st.markdown("### Step 3 — Upload Region Files")

    region_files: dict = {}
    for region in selected_regions:
        with st.expander(f"📂 {region} Files", expanded=True):
            region_files[region] = {}
            c1, c2 = st.columns(2)
            with c1:
                region_files[region]["lazada"] = st.file_uploader(
                    f"Lazada ({region}) — pricestock*.xlsx",
                    type=["xlsx","xls","csv"], key=f"laz_{region}",
                    help="EAN column: SellerSKU"
                )
                region_files[region]["shopee"] = st.file_uploader(
                    f"Shopee ({region}) — Shopee*Masterfile*.xlsx",
                    type=["xlsx","xls","csv"], key=f"sho_{region}",
                    help="EAN column: SKU"
                )
                region_files[region]["tiktok"] = st.file_uploader(
                    f"TikTok ({region}) — TikTokSellerCenter*.xlsx",
                    type=["xlsx","xls","csv"], key=f"ttk_{region}",
                    help="EAN column: Seller sku"
                )
            with c2:
                region_files[region]["zalora_status"] = st.file_uploader(
                    f"Zalora Status ({region}) — SellerStatusTemplate*.xlsx",
                    type=["xlsx","xls","csv"], key=f"zst_{region}",
                    help="EAN column: SellerSku | Status column"
                )
                region_files[region]["zalora_stock"] = st.file_uploader(
                    f"Zalora Stock ({region}) — SellerStockTemplate*.xlsx",
                    type=["xlsx","xls","csv"], key=f"zsk_{region}",
                    help="EAN column: SellerSku | Stock column"
                )
                region_files[region]["inventory"] = st.file_uploader(
                    f"Inventory ({region})",
                    type=["xlsx","xls","csv"], key=f"inv_{region}",
                    help=(
                        "PH → Inventory_*  |  "
                        "MY → PUMA_MY_B2C_Channel_Inventory_*  |  "
                        "SG → SG_PUMA SG B2C Inventory Rpt_New_*\n"
                        "Needs EAN column + stock columns (per channel or single total)"
                    )
                )

    st.markdown("---")
    run_col, _ = st.columns([2, 3])
    with run_col:
        run_btn = st.button(
            "🚀 Run Listing Audit", type="primary", use_container_width=True
        )

    if run_btn:
        errors = []
        if not zecom_file:   errors.append("ZeCom Tracker file is required.")
        if not content_file: errors.append("Content Master file is required.")
        if not selected_regions: errors.append("Select at least one region.")
        if errors:
            for e in errors:
                st.error(e)
        else:
            # ── Load master files ────────────────────────────────────────────
            prog = st.progress(0, text="Loading ZeCom Tracker…")

            with st.spinner("Loading ZeCom Tracker…"):
                zecom_df = load_zecom(zecom_file)
            st.success(f"✅ ZeCom Tracker: **{len(zecom_df):,}** articles")
            prog.progress(15, text="Loading Content Master…")

            with st.spinner("Loading Content Master…"):
                content_df = load_content(content_file)
            st.success(
                f"✅ Content Master: **{len(content_df):,}** EANs "
                f"across **{content_df['Article_No'].nunique():,}** articles"
            )
            prog.progress(25, text="Loading override…")

            override_map: dict = {}
            if override_file:
                with st.spinner("Loading Special Article Override…"):
                    override_map = load_override(override_file)
                st.info(
                    f"⚡ Special Override: **{len(override_map):,}** articles "
                    f"({sum(1 for v in override_map.values() if v=='ACTIVE')} ACTIVE, "
                    f"{sum(1 for v in override_map.values() if v=='INACTIVE')} INACTIVE)"
                )

            # ── Per-region audit ─────────────────────────────────────────────
            all_results: dict  = {}
            inv_debug_all: dict = {}
            step = 25
            step_size = int(70 / max(len(selected_regions), 1))

            for region in selected_regions:
                rf = region_files.get(region, {})
                prog.progress(step, text=f"[{region}] Loading marketplace files…")
                st.markdown(f"#### 📂 Region: {region}")

                # Load marketplace files
                mp_dfs: dict = {}
                if rf.get("lazada"):
                    with st.spinner(f"[{region}] Loading Lazada…"):
                        mp_dfs["Lazada"] = load_lazada(rf["lazada"])
                    st.write(f"  Lazada: **{len(mp_dfs['Lazada']):,}** EANs loaded")

                if rf.get("shopee"):
                    with st.spinner(f"[{region}] Loading Shopee…"):
                        mp_dfs["Shopee"] = load_shopee(rf["shopee"])
                    st.write(f"  Shopee: **{len(mp_dfs['Shopee']):,}** EANs loaded")

                if rf.get("zalora_status") and rf.get("zalora_stock"):
                    with st.spinner(f"[{region}] Loading Zalora…"):
                        mp_dfs["Zalora"] = load_zalora(
                            rf["zalora_status"], rf["zalora_stock"]
                        )
                    st.write(f"  Zalora: **{len(mp_dfs['Zalora']):,}** EANs loaded")
                elif rf.get("zalora_status"):
                    st.warning(f"[{region}] Zalora: Stock file missing — Zalora skipped")

                if rf.get("tiktok"):
                    with st.spinner(f"[{region}] Loading TikTok…"):
                        mp_dfs["TikTok"] = load_tiktok(rf["tiktok"])
                    st.write(f"  TikTok: **{len(mp_dfs['TikTok']):,}** EANs loaded")

                if not mp_dfs:
                    st.warning(f"[{region}] No marketplace files uploaded — region skipped.")
                    continue

                # Load inventory
                inv_df = pd.DataFrame(
                    columns=["EAN","Inv_Lazada","Inv_Shopee","Inv_Zalora","Inv_TikTok"]
                )
                if rf.get("inventory"):
                    with st.spinner(f"[{region}] Loading Inventory…"):
                        inv_df, inv_dbg = load_inventory(rf["inventory"], region)
                    inv_debug_all[region] = inv_dbg
                    st.write(
                        f"  Inventory: **{len(inv_df):,}** EANs "
                        f"— see 🔧 Inventory Debug tab to verify column mapping"
                    )
                else:
                    st.warning(
                        f"[{region}] No Inventory file — all stock values will be 0"
                    )

                prog.progress(step + step_size // 2,
                              text=f"[{region}] Running audit logic…")

                with st.spinner(f"[{region}] Running audit…"):
                    region_result = run_audit(
                        mp_dfs, inv_df, zecom_df, content_df, override_map, region
                    )

                all_results[region] = region_result

                # Region summary
                totals_r = {
                    cat: sum(
                        len(region_result[mp].get(cat, pd.DataFrame()))
                        for mp in MARKETPLACES if mp in region_result
                    )
                    for cat in ALL_CATS
                }
                st.success(
                    f"✅ [{region}] Done — "
                    + "  |  ".join(
                        f"{CAT_STYLE[c]['icon']} {c}: **{totals_r[c]:,}**"
                        for c in ALL_CATS
                    )
                )
                step += step_size

            prog.progress(100, text="Complete!")
            st.session_state.audit_results  = all_results
            st.session_state.inv_debug      = inv_debug_all

            if all_results:
                st.success("🎉 Audit complete! Go to **📊 Results & Download** tab.")

# ════════════════════ RESULTS ═════════════════════════════════════════════════
with tab_results:
    results = st.session_state.audit_results
    if not results:
        st.info("No results yet — run the audit in the **📁 Upload & Run** tab.")
    else:
        # ── Global KPIs ──────────────────────────────────────────────────────
        st.markdown("### 📊 Overall Totals")
        grand = {cat: 0 for cat in ALL_CATS}
        for mp_res in results.values():
            for mp_cats in mp_res.values():
                for cat in ALL_CATS:
                    grand[cat] += len(mp_cats.get(cat, pd.DataFrame()))

        kcols = st.columns(4)
        css_map = {
            CAT_ACTIVE: "cg", CAT_INACTIVE: "cr",
            CAT_NL_ART: "co", CAT_NL_EAN:  "cp",
        }
        for i, cat in enumerate(ALL_CATS):
            kcols[i].markdown(
                f"<div class='metric-box'>"
                f"<div class='metric-value {css_map[cat]}'>{grand[cat]:,}</div>"
                f"<div class='metric-label'>{CAT_STYLE[cat]['icon']} {cat}</div>"
                f"</div>",
                unsafe_allow_html=True,
            )

        # ── Summary table ─────────────────────────────────────────────────────
        st.markdown("### 📋 Breakdown by Region × Marketplace")
        rows = []
        for region, mp_res in results.items():
            for mp in MARKETPLACES:
                if mp not in mp_res:
                    continue
                row = {"Region": region, "Marketplace": mp}
                for cat in ALL_CATS:
                    row[f"{CAT_STYLE[cat]['icon']} {cat}"] = len(
                        mp_res[mp].get(cat, pd.DataFrame())
                    )
                rows.append(row)
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

        # ── Drilldown ─────────────────────────────────────────────────────────
        st.markdown("### 🔍 Drilldown")
        d1, d2, d3 = st.columns(3)
        dr_region = d1.selectbox("Region",      list(results.keys()))
        dr_mp     = d2.selectbox("Marketplace", [
            mp for mp in MARKETPLACES if mp in results.get(dr_region, {})
        ])
        dr_cat    = d3.selectbox("Category",    ALL_CATS)

        df_view = results[dr_region].get(dr_mp, {}).get(dr_cat, pd.DataFrame())
        icon    = CAT_STYLE[dr_cat]["icon"]
        st.markdown(
            f"**{icon} {dr_mp} [{dr_region}] — {dr_cat} : {len(df_view):,} records**"
        )
        if not df_view.empty:
            # Search filter
            search = st.text_input(
                "🔎 Search (Article No / EAN)", "",
                help="Filter rows containing this text in Article No or EAN columns"
            )
            if search.strip():
                mask = pd.Series(False, index=df_view.index)
                for col in ["Article No", "EAN (Seller SKU)"]:
                    if col in df_view.columns:
                        mask = mask | df_view[col].astype(str).str.contains(
                            search.strip(), case=False, na=False
                        )
                df_view = df_view[mask]
                st.caption(f"Showing {len(df_view):,} filtered records")

            st.dataframe(df_view, use_container_width=True, height=440)
        else:
            st.success("✅ No records in this category.")

        # ── Download ──────────────────────────────────────────────────────────
        st.markdown("---")
        st.markdown("### 💾 Download Full Audit Report")
        with st.spinner("Building Excel report…"):
            excel_bytes = build_excel(results)
        fname = f"PUMA_Listing_Audit_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        st.download_button(
            label="📥 Download Audit Report (.xlsx)",
            data=excel_bytes,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )
        st.caption(
            "Sheets: Summary + one sheet per Marketplace × Region × Category  "
            "(Active / Inactive / Not Listed Article / Not Listed EAN)"
        )

# ════════════════════ INVENTORY DEBUG ════════════════════════════════════════
with tab_debug:
    st.markdown("### 🔧 Inventory Column Mapping Debug")
    st.markdown(
        "Verify that the app correctly identified stock columns in your inventory file. "
        "If mapping is wrong, check the column names in your file."
    )
    inv_debug = st.session_state.inv_debug
    if not inv_debug:
        st.info("Run the audit first. Debug info will appear here after inventory files are loaded.")
    else:
        for region, dbg in inv_debug.items():
            st.markdown(f"#### Region: {region}")
            dbg_rows = [{"Field": k, "Detected Value / Column": str(v)} for k, v in dbg.items()]
            st.dataframe(pd.DataFrame(dbg_rows), use_container_width=True, hide_index=True)
            st.markdown("---")

        st.markdown("""
**How to read this table:**
- `Lazada stock col = lazada_qty` → stock for Lazada is taken from column named `lazada_qty` ✅
- `Lazada stock col = (not found — using 'qty_total')` → no Lazada-specific column; using total stock ⚠️
- `Lazada stock col = (not found — using 'None')` → no stock column found; stock will be 0 ❌

**To fix:** Make sure your inventory file has columns named to contain:
`lazada`, `shopee`, `zalora`, `tiktok` — OR — a single column named
`available`, `qty`, `stock`, `soh`, `on hand`, `quantity`
        """)
