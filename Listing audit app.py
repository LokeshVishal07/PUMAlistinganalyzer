"""
PUMA Listing Audit Analyzer  v5.0
===================================
CONFIRMED RULES (locked):

ACTIVE — ALL of these must be true:
  1. ZeCom Tracker = YES
  2. Launch Date is past (or blank)
  3. Avail_Qty (inventory) > 0  for that EAN on that channel

INACTIVE — ANY of these:
  1. ZeCom Tracker = NO or OFF
  2. Launch Date is in the future
  3. Avail_Qty (inventory) = 0

SPECIAL ARTICLE OVERRIDE:
  - Uploaded file: Article_No | Status (ACTIVE/INACTIVE) | Marketplace (col C)
  - If Status = ACTIVE  → active ONLY if Avail_Qty > 0, otherwise INACTIVE
  - If Status = INACTIVE → always INACTIVE regardless of stock
  - Marketplace column = which MP the override applies to (blank = all)

OUTPUT CATEGORIES per Marketplace per Region:
  🟢 Active           – ZeCom YES + past launch + stock > 0 + listed on MP as active
  🔴 Inactive         – ZeCom NO/OFF or future launch or stock=0 (listed or not)
  🟡 Not Listed - Article Level  – ZeCom ACTIVE but ZERO variants on MP at all
  🟣 Not Listed - EAN Level      – ZeCom ACTIVE, article partly listed, specific EAN missing

MARKETPLACE EAN IDENTIFIERS (confirmed):
  Lazada  → SellerSKU    + Product ID
  Shopee  → SKU          + Item ID / Product ID
  Zalora  → SellerSku
  TikTok  → Seller sku

INVENTORY: Avail_Qty column = source of truth for ALL channels
  PH Lazada: effective = Avail_Qty - 1 (floor 0)
  All others: Avail_Qty directly
"""

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import warnings
warnings.filterwarnings("ignore")

st.set_page_config(page_title="PUMA Listing Audit Analyzer", layout="wide", page_icon="📊")

st.markdown("""
<style>
.main-header{background:linear-gradient(135deg,#1a1a2e 0%,#16213e 50%,#0f3460 100%);
  padding:2rem;border-radius:12px;text-align:center;margin-bottom:1.5rem;color:white;}
.main-header h1{font-size:2rem;font-weight:700;margin:0;}
.main-header p{color:#a0aec0;margin-top:.4rem;font-size:.9rem;}
.metric-box{background:white;border-radius:8px;padding:.9rem;
  border:1px solid #e2e8f0;text-align:center;margin-bottom:.5rem;}
.metric-value{font-size:1.7rem;font-weight:800;}
.metric-label{font-size:.75rem;color:#718096;margin-top:.15rem;}
.cg{color:#276221;}.cr{color:#9c0006;}.co{color:#7b5800;}.cp{color:#4a235a;}
.dbg{background:#f0fff4;border:1px solid #9ae6b4;border-radius:6px;
  padding:.5rem .8rem;font-size:.78rem;margin:.3rem 0;font-family:monospace;}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class='main-header'>
  <h1>📊 PUMA Listing Audit Analyzer</h1>
  <p>ZeCom-Driven · Stock-Gated · Article &amp; Variant Level · MY / SG / PH</p>
</div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════
REGIONS      = ["MY", "SG", "PH"]
MARKETPLACES = ["Lazada", "Shopee", "Zalora", "TikTok"]

CHANNEL_BUFFER = {
    "PH": {"Lazada": 1, "Shopee": 0, "Zalora": 0, "TikTok": 0},
    "MY": {"Lazada": 0, "Shopee": 0, "Zalora": 0, "TikTok": 0},
    "SG": {"Lazada": 0, "Shopee": 0, "Zalora": 0, "TikTok": 0},
}

# Confirmed EAN column names per marketplace (priority order)
MP_EAN_COLS = {
    "Lazada": ["SellerSKU","Seller SKU","seller_sku","SellerSku","Seller Sku","SKU"],
    "Shopee": ["SKU","sku","Variation SKU","variation_sku","Model SKU",
               "Seller SKU","SellerSKU","Item SKU"],
    "Zalora": ["SellerSku","Seller Sku","SellerSKU","Seller SKU","seller_sku"],
    "TikTok": ["Seller sku","Seller SKU","SellerSKU","seller_sku","SKU"],
}

# Marketplace product/item ID column names
MP_ID_COLS = {
    "Lazada": ["ItemId","Item Id","item_id","Product ID","ProductId","product_id","ITEM ID"],
    "Shopee": ["Item ID","ItemId","item_id","Product ID","ProductId",
               "Item Id","itemid","parent_sku","Parent Item ID"],
    "Zalora": ["ConfigSku","Config SKU","config_sku","ProductId","Product ID"],
    "TikTok": ["Product ID","ProductId","product_id","Item ID","ItemId"],
}

# Known INACTIVE status strings across all marketplaces
MP_INACTIVE_STATUSES = {
    "INACTIVE","INACTIVATED","DEACTIVATED",
    "DELETED","SELLER_DELETED","BANNED",
    "UNLISTED","UNLIST",
    "SUSPENDED","BLOCKED",
    "VIOLATION","DELISTED",
    "REJECTED","FAILED",
    "PROHIBITED","TAKEN DOWN",
    "0","FALSE","NO","OFF",
    "NOT LISTED","NOT_LISTED",
    "SOLDOUT","SOLD OUT",
}

CAT_ACTIVE   = "Active"
CAT_INACTIVE = "Inactive"
CAT_NL_ART   = "Not Listed - Article Level"
CAT_NL_EAN   = "Not Listed - EAN Level"
ALL_CATS     = [CAT_ACTIVE, CAT_INACTIVE, CAT_NL_ART, CAT_NL_EAN]

CAT_STYLE = {
    CAT_ACTIVE  : {"hdr":"#1e6f39","bg":"#c6efce","fc":"#276221","icon":"🟢"},
    CAT_INACTIVE: {"hdr":"#9c0006","bg":"#ffc7ce","fc":"#9c0006","icon":"🔴"},
    CAT_NL_ART  : {"hdr":"#7b5800","bg":"#ffeb9c","fc":"#7b5800","icon":"🟡"},
    CAT_NL_EAN  : {"hdr":"#4a235a","bg":"#e8d5f5","fc":"#4a235a","icon":"🟣"},
}

# ══════════════════════════════════════════════════════════════════════════════
# COLUMN HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def _sc(c) -> str:
    try:
        s = str(c).strip()
        return "" if s in ("nan","None","NaT","") else s
    except:
        return ""

def _cl(c) -> str:
    return _sc(c).lower()

def sanitize_cols(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [_sc(c) or f"_col_{i}" for i,c in enumerate(df.columns)]
    return df

def find_header_row(file, max_scan: int = 30) -> int:
    """Scan first max_scan rows, return the one with most non-empty string cells."""
    try:
        prev = pd.read_excel(file, header=None, nrows=max_scan)
        file.seek(0)
        best_r, best_s = 0, -1
        for i, row in prev.iterrows():
            sc = int(row.dropna().apply(
                lambda v: isinstance(v,str) and len(v.strip()) > 0
            ).sum())
            if sc > best_s:
                best_s, best_r = sc, int(i)
        return best_r
    except:
        try: file.seek(0)
        except: pass
        return 0

def safe_read(file, sheet=0) -> pd.DataFrame:
    """Read Excel/CSV with auto header detection. Always returns string-column DataFrame."""
    try:
        hdr = find_header_row(file)
        df  = pd.read_excel(file, header=hdr, sheet_name=sheet)
    except:
        try:
            file.seek(0)
            df = pd.read_excel(file, header=0)
        except:
            try:
                file.seek(0)
                df = pd.read_csv(file, header=0)
            except Exception as e:
                st.warning(f"Cannot read file: {e}")
                return pd.DataFrame()
    df = sanitize_cols(df)
    return df.dropna(how="all").reset_index(drop=True)

def norm_col(df: pd.DataFrame, candidates: list, new_name: str) -> pd.DataFrame:
    cands = [str(x).strip().lower() for x in candidates]
    for col in df.columns:
        if _cl(col) in cands:
            if col != new_name:
                df = df.rename(columns={col: new_name})
            break
    return df

def first_col_with(df: pd.DataFrame, keywords: list, exclude: set = None) -> str | None:
    excl = exclude or set()
    for kw in keywords:
        for c in df.columns:
            if kw in _cl(c) and c not in excl:
                return c
    return None

def cs(s: pd.Series) -> pd.Series:
    return s.astype(str).str.strip().str.upper().replace(
        {"NAN":"","NONE":"","NAT":"","NA":"","<NA>":""}
    )

def to_num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0).clip(lower=0).astype(int)

def is_mp_active(val) -> bool:
    """Blacklist approach — anything not explicitly inactive = active."""
    if pd.isna(val): return False
    v = str(val).strip().upper()
    if v in ("","NAN","NONE","N/A","-"): return False
    return v not in MP_INACTIVE_STATUSES

def resolve_tracker(val) -> str:
    if pd.isna(val): return "UNKNOWN"
    v = str(val).strip().upper()
    if v == "YES": return "ACTIVE"
    if v in ("NO","OFF"): return "INACTIVE"
    return "UNKNOWN"

# ══════════════════════════════════════════════════════════════════════════════
# FILE LOADERS
# ══════════════════════════════════════════════════════════════════════════════

def load_zecom(file) -> pd.DataFrame:
    df = safe_read(file)
    df = norm_col(df, [
        "PIM Article#","PIM Article","Article No","ArticleNo",
        "Color_No","ColorNo","Color No","Article Number",
        "Style Number","StyleNo","Style No","Parent SKU",
    ], "Article_No")
    if "Article_No" not in df.columns:
        st.error("❌ ZeCom: Cannot find Article No column.")
        return pd.DataFrame(columns=["Article_No"])
    reserved = {"Article_No"}
    for mp in MARKETPLACES:
        already = {f"Tracker_{m}" for m in MARKETPLACES}
        cols = [c for c in df.columns
                if mp.lower() in _cl(c) and c not in already and c not in reserved]
        if cols:
            df = df.rename(columns={cols[0]: f"Tracker_{mp}"})
    for mp in MARKETPLACES:
        if f"Tracker_{mp}" not in df.columns:
            df[f"Tracker_{mp}"] = np.nan
    skip = reserved | {f"Tracker_{m}" for m in MARKETPLACES}
    date_cols = [c for c in df.columns
                 if any(k in _cl(c) for k in ["launch","go live","golive","live date"])
                 and c not in skip]
    if date_cols:
        df = df.rename(columns={date_cols[0]: "Launch_Date"})
        df["Launch_Date"] = pd.to_datetime(df["Launch_Date"], errors="coerce")
    else:
        df["Launch_Date"] = pd.NaT
    df["Article_No"] = cs(df["Article_No"])
    df = df[df["Article_No"].str.len() > 0].drop_duplicates("Article_No").reset_index(drop=True)
    return df


def load_override(file) -> pd.DataFrame:
    """
    Special Article Override.
    Expected columns: Article_No | Status (ACTIVE/INACTIVE) | Marketplace (optional, col C)
    If Marketplace blank/missing → applies to ALL marketplaces.
    Stock check still applies: ACTIVE override needs Avail_Qty > 0 to be active.
    """
    df = safe_read(file)
    df = norm_col(df, [
        "PIM Article#","PIM Article","Article No","ArticleNo",
        "Color_No","ColorNo","Article Number","Style Number",
    ], "Article_No")
    df = norm_col(df, ["Status","status","Override Status","Listing Status","Flag"], "Status")

    # Marketplace column (3rd column / col C)
    mp_col_candidates = ["Marketplace","marketplace","Channel","channel","Platform"]
    mp_col = None
    for c in df.columns:
        if _cl(c) in [x.lower() for x in mp_col_candidates]:
            mp_col = c
            break
    # Fallback: use 3rd column if it exists and has MP-like values
    if not mp_col and len(df.columns) >= 3:
        third_col = df.columns[2]
        sample = df[third_col].dropna().astype(str).str.upper().unique()
        if any(mp.upper() in " ".join(sample) for mp in MARKETPLACES):
            mp_col = third_col

    if "Article_No" not in df.columns or "Status" not in df.columns:
        st.warning("⚠️ Special Override: needs Article_No and Status columns.")
        return pd.DataFrame(columns=["Article_No","Status","Marketplace"])

    df["Article_No"] = cs(df["Article_No"])
    df["Status"]     = cs(df["Status"])
    if mp_col:
        df["Marketplace"] = df[mp_col].astype(str).str.strip().str.upper()
        df["Marketplace"] = df["Marketplace"].replace({"NAN":"ALL","NONE":"ALL","":"ALL"})
    else:
        df["Marketplace"] = "ALL"

    df = df[df["Article_No"].str.len() > 0]
    df = df[df["Status"].isin(["ACTIVE","INACTIVE"])]
    return df[["Article_No","Status","Marketplace"]].reset_index(drop=True)


def load_content(file) -> pd.DataFrame:
    """ONLY for Article_No → EAN mapping. NOT used for active/inactive decisions."""
    df = safe_read(file)
    df = norm_col(df, [
        "EAN","ean","Barcode","barcode","Child SKU","ChildSKU",
        "Seller SKU","SellerSKU","Item EAN","GTIN","UPC",
    ], "EAN")
    df = norm_col(df, [
        "Color_No","ColorNo","Article No","ArticleNo",
        "PIM Article#","Article Number","Parent SKU","Color No",
        "Style Number","StyleNo",
    ], "Article_No")
    extra = []
    for label, keys in [("Size",["size","sz"]),
                         ("Color",["colour","color"]),
                         ("Description",["desc","product name","name"])]:
        col = first_col_with(df, keys, exclude={"EAN","Article_No"})
        if col:
            df = df.rename(columns={col: label})
            extra.append(label)
    for col in ["EAN","Article_No"]:
        if col not in df.columns:
            df[col] = np.nan
    df["EAN"]        = cs(df["EAN"])
    df["Article_No"] = cs(df["Article_No"])
    df = df[(df["EAN"].str.len() > 0) & (df["Article_No"].str.len() > 0)]
    keep = ["EAN","Article_No"] + [e for e in extra if e in df.columns]
    return df[keep].drop_duplicates("EAN").reset_index(drop=True)


def load_inventory(file, region: str) -> tuple:
    """
    Source of truth for stock.
    Primary column: Avail_Qty (exact or case-insensitive).
    Returns (DataFrame[EAN, Inv_Stock], debug_dict).
    Inv_Stock = Avail_Qty with PH Lazada buffer applied per channel at audit time.
    """
    df = safe_read(file)
    all_cols = list(df.columns)

    df = norm_col(df, [
        "EAN","ean","Barcode","barcode","Item EAN","GTIN",
        "SKU","sku","Item Code","Seller SKU","SellerSKU",
        "Material","Material Number",
    ], "EAN")

    if "EAN" not in df.columns:
        st.warning(f"[{region}] Inventory: Cannot find EAN column. Columns: {all_cols[:20]}")
        return pd.DataFrame(columns=["EAN","Inv_Stock"]), {"Error": "EAN not found"}

    df["EAN"] = cs(df["EAN"])
    df = df[df["EAN"].str.len() > 0]

    # PRIMARY: Avail_Qty (exact → case-insensitive → fallback)
    avail_col = None
    if "Avail_Qty" in df.columns:
        avail_col = "Avail_Qty"
    else:
        for c in df.columns:
            if _cl(c) == "avail_qty":
                avail_col = c
                break
    # Fallback to other known stock column names
    if not avail_col:
        avail_col = first_col_with(df, [
            "avail_qty","available","on hand","onhand","total",
            "qty","quantity","stock","soh","free","unrestricted",
        ], exclude={"EAN"})

    debug = {
        "Stock column used": avail_col or "NOT FOUND (stock = 0)",
        "EAN rows loaded"  : len(df),
        "All columns"      : ", ".join(all_cols[:50]),
    }

    if not avail_col:
        st.warning(
            f"[{region}] Inventory: 'Avail_Qty' not found. "
            f"Stock will be 0. Columns: {all_cols[:20]}"
        )
        df["Inv_Stock"] = 0
    else:
        df["Inv_Stock"] = to_num(df[avail_col])

    result = df[["EAN","Inv_Stock"]].drop_duplicates("EAN").reset_index(drop=True)
    return result, debug


def _load_mp_generic(df: pd.DataFrame, mp: str, extra_id_cols: list) -> pd.DataFrame:
    """
    Load a marketplace DataFrame.
    Returns: EAN, MP_Status, MP_Stock, MP_ID, Marketplace
    Also preserves raw MP_Status_Raw for debugging.
    """
    # EAN
    df = norm_col(df, MP_EAN_COLS[mp], "EAN")

    # Status — broad list to catch all known MP status column names
    df = norm_col(df, [
        "Status","status","ItemStatus","Item Status","Listing Status",
        "Product Status","Active","Publish","Activation","State","Condition",
        "listing_status","item_status","product_status",
    ], "MP_Status")

    # Stock from MP's own file
    df = norm_col(df, [
        "Available","Stock","Quantity","available","quantity","Qty",
        "FreeQty","sellable_quantity","Current Stock","Available Stock",
        "free qty","stock_quantity","sellable_stock",
    ], "MP_Stock")

    # Product/Item ID
    id_found = None
    for cand in extra_id_cols + MP_ID_COLS.get(mp,[]):
        for col in df.columns:
            if _cl(col) == cand.lower().strip():
                id_found = col
                break
        if id_found:
            break
    if id_found:
        df = df.rename(columns={id_found: "MP_ID"})
    else:
        df["MP_ID"] = ""

    for col in ["EAN","MP_Status","MP_Stock"]:
        if col not in df.columns:
            df[col] = np.nan

    # Keep raw status for debug
    df["MP_Status_Raw"] = df["MP_Status"].astype(str).str.strip()

    df["EAN"]         = cs(df["EAN"])
    df["MP_Stock"]    = to_num(df["MP_Stock"])
    df["MP_ID"]       = df["MP_ID"].astype(str).str.strip().replace({"nan":"","None":""})
    df["Marketplace"] = mp
    df = df[df["EAN"].str.len() > 0]

    return df[["EAN","MP_Status","MP_Status_Raw","MP_Stock","MP_ID","Marketplace"]].drop_duplicates("EAN")


def load_lazada(file) -> pd.DataFrame:
    df = safe_read(file)
    return _load_mp_generic(df, "Lazada", ["ItemId","Item Id","item_id"])

def load_shopee(file) -> pd.DataFrame:
    """
    Shopee Masterfile can have product-level + variation-level rows.
    The SKU (variation) is what maps to EAN.
    We try all sheets and pick the one with the most SKU data.
    """
    # Try reading all sheets, pick the best one
    best_df = pd.DataFrame()
    try:
        xf = pd.ExcelFile(file)
        sheets = xf.sheet_names
    except:
        sheets = [0]

    for sh in sheets:
        try:
            file.seek(0)
            hdr = find_header_row(file)
            file.seek(0)
            tmp = pd.read_excel(file, header=hdr, sheet_name=sh)
            tmp = sanitize_cols(tmp)
            tmp = tmp.dropna(how="all").reset_index(drop=True)

            # Score: how many rows have a plausible EAN/SKU column
            ean_col_found = any(
                _cl(c) in [x.lower() for x in MP_EAN_COLS["Shopee"]]
                for c in tmp.columns
            )
            if ean_col_found and len(tmp) > len(best_df):
                best_df = tmp
        except:
            continue

    if best_df.empty:
        file.seek(0)
        best_df = safe_read(file)

    return _load_mp_generic(best_df, "Shopee", ["Item ID","ItemId","itemid","item_id"])

def load_tiktok(file) -> pd.DataFrame:
    df = safe_read(file)
    return _load_mp_generic(df, "TikTok", ["Product ID","ProductId","product_id"])

def load_zalora(status_file, stock_file) -> pd.DataFrame:
    ds   = safe_read(status_file)
    dstk = safe_read(stock_file)
    ds   = norm_col(ds,   MP_EAN_COLS["Zalora"], "EAN")
    ds   = norm_col(ds,   ["Status","status","ItemStatus","Item Status","Active"], "MP_Status")
    dstk = norm_col(dstk, MP_EAN_COLS["Zalora"], "EAN")
    dstk = norm_col(dstk, ["Stock","Quantity","Available","Qty","quantity"], "MP_Stock")
    for col in ["EAN","MP_Status"]:
        if col not in ds.columns: ds[col] = np.nan
    if "EAN"      not in dstk.columns: dstk["EAN"]      = np.nan
    if "MP_Stock" not in dstk.columns: dstk["MP_Stock"]  = np.nan
    ds["EAN"]         = cs(ds["EAN"])
    dstk["EAN"]       = cs(dstk["EAN"])
    ds["MP_Status_Raw"] = ds["MP_Status"].astype(str).str.strip()
    ds["MP_ID"]         = ""
    merged = ds.merge(dstk[["EAN","MP_Stock"]].drop_duplicates("EAN"), on="EAN", how="left")
    merged["MP_Stock"]    = to_num(merged["MP_Stock"])
    merged["Marketplace"] = "Zalora"
    merged = merged[merged["EAN"].str.len() > 0]
    return merged[["EAN","MP_Status","MP_Status_Raw","MP_Stock","MP_ID","Marketplace"]].drop_duplicates("EAN")


# ══════════════════════════════════════════════════════════════════════════════
# STATUS DECISION ENGINE
# ══════════════════════════════════════════════════════════════════════════════
def decide_status(
    zecom_row   : pd.Series,
    mp          : str,
    inv_stock   : int,          # effective stock (after buffer)
    override_df : pd.DataFrame, # full override DataFrame
    art_no      : str,
) -> tuple:
    """
    Returns (expected_status: 'ACTIVE'|'INACTIVE', reasons: list[str])

    PRIORITY:
      1. Special Override (checks stock too)
      2. ZeCom Tracker = NO/OFF → INACTIVE
      3. Launch Date future → INACTIVE
      4. ZeCom Tracker = YES + stock > 0 → ACTIVE
      5. ZeCom Tracker = YES + stock = 0 → INACTIVE
      6. Blank tracker → INACTIVE
    """
    reasons = []

    # ── PRIORITY 1: Special Override ─────────────────────────────────────────
    if not override_df.empty:
        # Match article + (marketplace = this MP or ALL)
        mp_upper = mp.upper()
        mask = (
            (override_df["Article_No"] == art_no) &
            (override_df["Marketplace"].isin([mp_upper, "ALL"]))
        )
        match = override_df[mask]
        if not match.empty:
            ov_status = match.iloc[0]["Status"]
            if ov_status == "ACTIVE":
                if inv_stock > 0:
                    reasons.append(f"Special Override = ACTIVE + stock {inv_stock} > 0 ✓")
                    return "ACTIVE", reasons
                else:
                    reasons.append(f"Special Override = ACTIVE but stock = 0 → INACTIVE")
                    return "INACTIVE", reasons
            else:
                reasons.append(f"Special Override = INACTIVE")
                return "INACTIVE", reasons

    # ── PRIORITY 2: ZeCom Tracker ─────────────────────────────────────────────
    t_raw    = zecom_row.get(f"Tracker_{mp}", np.nan)
    t_status = resolve_tracker(t_raw)
    t_str    = str(t_raw).strip() if pd.notna(t_raw) else "blank"

    if t_status == "INACTIVE":
        reasons.append(f"ZeCom Tracker = {t_str} (NO/OFF → INACTIVE)")
        return "INACTIVE", reasons

    # ── PRIORITY 3: Launch Date ───────────────────────────────────────────────
    launch = zecom_row.get("Launch_Date", pd.NaT)
    today  = pd.Timestamp.today().normalize()
    if pd.notna(launch):
        try:
            lts = pd.Timestamp(launch).normalize()
            if lts > today:
                reasons.append(f"Future Launch Date: {lts.date()}")
                return "INACTIVE", reasons
            else:
                reasons.append(f"Launch Date: {lts.date()} ✓ (past)")
        except:
            pass

    # ── PRIORITY 4 & 5: Tracker YES — gate on stock ───────────────────────────
    if t_status == "ACTIVE":
        if inv_stock > 0:
            reasons.append(f"ZeCom = YES + Avail_Qty = {inv_stock} > 0 ✓")
            return "ACTIVE", reasons
        else:
            reasons.append(f"ZeCom = YES but Avail_Qty = 0 → INACTIVE")
            return "INACTIVE", reasons

    # ── PRIORITY 6: Blank/Unknown ─────────────────────────────────────────────
    reasons.append(f"ZeCom Tracker = {t_str} (blank/unknown → INACTIVE)")
    return "INACTIVE", reasons


# ══════════════════════════════════════════════════════════════════════════════
# INVENTORY LOOKUP
# ══════════════════════════════════════════════════════════════════════════════
def get_eff_stock(ean: str, mp: str, region: str, inv_idx: dict) -> int:
    """Return effective stock for this EAN+channel+region (buffer applied)."""
    row = inv_idx.get(ean)
    raw = int(row["Inv_Stock"]) if row is not None else 0
    buf = CHANNEL_BUFFER.get(region, {}).get(mp, 0)
    return max(0, raw - buf)

def stock_disc(inv_stock: int, mp_stock: int) -> str:
    if inv_stock > 0 and mp_stock == 0:
        return f"⚠ Inventory={inv_stock} but MP shows 0"
    if inv_stock == 0 and mp_stock > 0:
        return f"⚠ Inventory=0 but MP shows {mp_stock}"
    return "-"


# ══════════════════════════════════════════════════════════════════════════════
# AUDIT ENGINE
# ══════════════════════════════════════════════════════════════════════════════
def run_audit(
    mp_dfs       : dict,
    inv_df       : pd.DataFrame,
    zecom_df     : pd.DataFrame,
    content_df   : pd.DataFrame,
    override_df  : pd.DataFrame,
    region       : str,
) -> dict:

    # ── Lookup maps ──────────────────────────────────────────────────────────
    art_to_eans: dict[str,list] = {}
    for _, row in content_df.iterrows():
        art = str(row["Article_No"]).strip()
        ean = str(row["EAN"]).strip()
        if art and ean:
            art_to_eans.setdefault(art,[]).append(ean)

    extra_cols = [c for c in content_df.columns if c not in ("EAN","Article_No")]
    ean_meta: dict[str,dict] = {}
    for _, row in content_df.iterrows():
        ean = str(row["EAN"]).strip()
        if ean:
            ean_meta[ean] = {c: row[c] for c in extra_cols if c in row.index}

    inv_idx: dict[str,pd.Series] = {}
    if not inv_df.empty:
        for _, row in inv_df.iterrows():
            inv_idx[str(row["EAN"]).strip()] = row

    results = {}

    for mp in MARKETPLACES:
        mp_df = mp_dfs.get(mp, pd.DataFrame())

        mp_idx: dict[str,pd.Series] = {}
        if not mp_df.empty:
            for _, row in mp_df.iterrows():
                e = str(row["EAN"]).strip()
                if e:
                    mp_idx[e] = row

        active_rows   = []
        inactive_rows = []
        nl_art_rows   = []
        nl_ean_rows   = []

        for _, z_row in zecom_df.iterrows():
            art = str(z_row.get("Article_No","")).strip()
            if not art:
                continue

            variants     = art_to_eans.get(art,[])
            listed_eans  = [e for e in variants if e in mp_idx]
            n_total      = len(variants)
            n_listed     = len(listed_eans)

            # ── No variants in content ────────────────────────────────────────
            if not variants:
                # Determine status using a dummy stock of 0 (no EAN to check)
                exp, reasons = decide_status(z_row, mp, 0, override_df, art)
                base = {
                    "Region"           : region,
                    "Marketplace"      : mp,
                    "Article No"       : art,
                    "EAN (Seller SKU)" : "-",
                    "MP ID"            : "-",
                    "Expected Status"  : exp,
                    "MP Listed"        : "NO",
                    "MP Status"        : "-",
                    "Avail_Qty (Inv)"  : 0,
                    "MP Stock"         : 0,
                    "Stock Discrepancy": "-",
                    "Reason"           : "; ".join(reasons),
                    "Action Required"  : "-",
                    "Note"             : "No EAN variants in Content File",
                }
                if exp == "INACTIVE":
                    inactive_rows.append(base)
                else:
                    nl_art_rows.append({**base,
                        "Note":"ACTIVE per ZeCom but NO EAN variants in Content File"})
                continue

            # ── Process each variant ──────────────────────────────────────────
            for ean in variants:
                mp_row    = mp_idx.get(ean)
                eff_stock = get_eff_stock(ean, mp, region, inv_idx)
                raw_stock = int(inv_idx[ean]["Inv_Stock"]) if ean in inv_idx else 0
                meta      = ean_meta.get(ean, {})
                mp_stock  = int(mp_row["MP_Stock"]) if mp_row is not None else 0
                mp_id     = str(mp_row["MP_ID"])    if mp_row is not None else "-"
                mp_status_raw = str(mp_row["MP_Status_Raw"]) if mp_row is not None else "Not Listed"

                exp, reasons = decide_status(z_row, mp, eff_stock, override_df, art)

                disc = stock_disc(eff_stock, mp_stock)

                base_rec = {
                    "Region"           : region,
                    "Marketplace"      : mp,
                    "Article No"       : art,
                    "EAN (Seller SKU)" : ean,
                    **meta,
                    "MP ID"            : mp_id,
                    "Expected Status"  : exp,
                    "MP Listed"        : "YES" if mp_row is not None else "NO",
                    "MP Status"        : mp_status_raw,
                    "Avail_Qty (Inv)"  : eff_stock,
                    "MP Stock"         : mp_stock,
                    "Stock Discrepancy": disc,
                    "Reason"           : "; ".join(reasons),
                    "Action Required"  : "-",
                }

                if mp_row is None:
                    # EAN not on MP at all
                    if exp == "INACTIVE":
                        base_rec["Action Required"] = "-"
                        inactive_rows.append(base_rec)
                    else:
                        # Expected ACTIVE but not listed
                        if n_listed == 0:
                            base_rec["Action Required"] = "List entire article on MP"
                            nl_art_rows.append(base_rec)
                        else:
                            base_rec["Reason"] += (
                                f"; {n_listed}/{n_total} variants listed — this EAN/size missing"
                            )
                            base_rec["Action Required"] = "Add missing variant/size to listing"
                            nl_ean_rows.append(base_rec)
                else:
                    # EAN is on MP
                    mp_active = is_mp_active(mp_row["MP_Status"])

                    if exp == "ACTIVE" and mp_active:
                        if disc != "-":
                            base_rec["Action Required"] = "Investigate stock discrepancy"
                        active_rows.append(base_rec)

                    elif exp == "ACTIVE" and not mp_active:
                        base_rec["Reason"] += "; MP status is INACTIVE despite ZeCom = ACTIVE"
                        base_rec["Action Required"] = "Activate listing on MP"
                        # Goes to inactive (MP currently inactive)
                        inactive_rows.append(base_rec)

                    elif exp == "INACTIVE":
                        base_rec["Action Required"] = (
                            "Delist / deactivate on MP" if mp_active else "-"
                        )
                        inactive_rows.append(base_rec)

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
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb = writer.book

        def fmt(**kw):
            d = {"font_name":"Arial","font_size":9,"border":1,
                 "valign":"vcenter","align":"left"}
            d.update(kw)
            return wb.add_format(d)

        ttl  = fmt(bold=True, font_size=13, font_color="#0f3460", border=0)
        sub  = fmt(italic=True, font_size=8, font_color="#718096", border=0)
        norm = fmt()
        date_fmt = fmt()

        def chdr(bg, fc="#ffffff"):
            return fmt(bold=True, bg_color=bg, font_color=fc,
                       align="center", text_wrap=True)
        def cdata(bg, fc="#000000"):
            return fmt(bg_color=bg, font_color=fc)

        # ── Summary sheet ────────────────────────────────────────────────────
        ws = wb.add_worksheet("Summary")
        writer.sheets["Summary"] = ws
        ws.set_zoom(90)
        ws.freeze_panes(5, 2)
        ws.set_column("A:A", 8)
        ws.set_column("B:B", 14)
        ws.set_column("C:F", 30)
        ws.write("A1", "PUMA Listing Audit — Summary", ttl)
        ws.write("A2", f"Generated: {datetime.now().strftime('%d %b %Y  %H:%M')}", sub)
        ws.write("A3", f"Logic: ZeCom YES + Avail_Qty > 0 + Past Launch = ACTIVE", sub)

        r = 4
        ws.write(r, 0, "Region",      chdr("#0f3460"))
        ws.write(r, 1, "Marketplace", chdr("#0f3460"))
        for ci, cat in enumerate(ALL_CATS):
            ws.write(r, ci+2,
                     f"{CAT_STYLE[cat]['icon']} {cat}",
                     chdr(CAT_STYLE[cat]["hdr"]))
        ws.set_row(r, 35)
        r += 1

        for region, mp_res in all_results.items():
            for mp in MARKETPLACES:
                if mp not in mp_res: continue
                ws.write(r, 0, region, norm)
                ws.write(r, 1, mp,     norm)
                for ci, cat in enumerate(ALL_CATS):
                    cnt = len(mp_res[mp].get(cat, pd.DataFrame()))
                    ws.write(r, ci+2, cnt, cdata(CAT_STYLE[cat]["bg"],
                                                   CAT_STYLE[cat]["fc"]))
                r += 1

        # ── Detail sheets ────────────────────────────────────────────────────
        for region, mp_res in all_results.items():
            for mp in MARKETPLACES:
                if mp not in mp_res: continue
                for cat in ALL_CATS:
                    df = mp_res[mp].get(cat, pd.DataFrame())
                    sname = f"{mp[:5]} {region} {cat[:17]}"[:31]

                    ws2 = wb.add_worksheet(sname)
                    writer.sheets[sname] = ws2
                    ws2.set_zoom(85)

                    ws2.write(0, 0,
                              f"{CAT_STYLE[cat]['icon']} {mp} [{region}] — {cat}", ttl)
                    ws2.write(1, 0, f"Total records: {len(df)}", sub)

                    if df.empty:
                        ws2.write(2, 0, "No records in this category.", sub)
                        continue

                    hf  = chdr(CAT_STYLE[cat]["hdr"])
                    rfd = cdata(CAT_STYLE[cat]["bg"], CAT_STYLE[cat]["fc"])

                    highlight = {"Expected Status","Reason","Action Required",
                                 "MP Status","Stock Discrepancy"}

                    for ci, col in enumerate(df.columns):
                        ws2.write(2, ci, col, hf)
                        try:
                            max_w = max(
                                len(str(col)),
                                int(df[col].astype(str).str.len().max())
                            )
                        except:
                            max_w = len(str(col))
                        ws2.set_column(ci, ci, min(max_w + 3, 50))

                    ws2.freeze_panes(3, 0)

                    for ri, (_, rec) in enumerate(df.iterrows()):
                        for ci, col in enumerate(df.columns):
                            val = rec[col]
                            sv  = "" if (isinstance(val,float) and np.isnan(val)) else str(val)
                            f   = rfd if col in highlight else norm
                            ws2.write(ri+3, ci, sv, f)

    return output.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════════════════════════
for key in ("audit_results","inv_debug","mp_debug"):
    if key not in st.session_state:
        st.session_state[key] = {}

# ══════════════════════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════════════════════
tab_upload, tab_results, tab_debug, tab_help = st.tabs([
    "📁 Upload & Run","📊 Results & Download","🔧 Debug","❓ Help",
])

# ════ HELP ════════════════════════════════════════════════════════════════════
with tab_help:
    st.markdown("""
## ✅ Active / Inactive Decision Rules (v5.0)

| Condition | Result |
|---|---|
| ZeCom = YES **AND** past launch **AND** Avail_Qty > 0 | **ACTIVE** |
| ZeCom = NO or OFF | **INACTIVE** |
| Launch Date is in the future | **INACTIVE** |
| ZeCom = YES but Avail_Qty = 0 | **INACTIVE** |
| Tracker blank / unknown | **INACTIVE** |

### Special Article Override (col C = Marketplace)
| Override Status | Stock | Result |
|---|---|---|
| ACTIVE | Avail_Qty > 0 | **ACTIVE** |
| ACTIVE | Avail_Qty = 0 | **INACTIVE** (no stock) |
| INACTIVE | any | **INACTIVE** |

## 📋 Output Categories
| Icon | Category | Meaning |
|---|---|---|
| 🟢 | Active | ZeCom YES + past launch + stock + listed active on MP |
| 🔴 | Inactive | Any inactive condition, listed or not |
| 🟡 | Not Listed – Article Level | Should be active, ZERO variants on MP |
| 🟣 | Not Listed – EAN Level | Should be active, some variants on MP, specific EAN missing |

## 🔑 EAN Column Names (confirmed)
| MP | EAN Col | ID Col |
|---|---|---|
| Lazada | `SellerSKU` | `ItemId` / `Product ID` |
| Shopee | `SKU` | `Item ID` |
| Zalora | `SellerSku` | — |
| TikTok | `Seller sku` | `Product ID` |

## 📦 Inventory
- Column: **`Avail_Qty`** (primary) — used for ALL channels
- PH Lazada: effective = Avail_Qty − 1 (min 0)
- All other channels / regions: Avail_Qty directly
    """)

# ════ UPLOAD ══════════════════════════════════════════════════════════════════
with tab_upload:
    st.markdown("### Step 1 — Select Regions")
    selected_regions = st.multiselect("Regions:", REGIONS, default=REGIONS)

    st.markdown("### Step 2 — Master Files *(all regions)*")
    mc1, mc2, mc3 = st.columns(3)
    with mc1:
        zecom_file = st.file_uploader(
            "📋 ZeCom Tracker **[required]**",
            type=["xlsx","xls","csv"], key="zecom",
            help="Article No | Lazada YES/NO/OFF | Shopee | Zalora | TikTok | Launch Date"
        )
    with mc2:
        content_file = st.file_uploader(
            "📦 Content Master **[required]**",
            type=["xlsx","xls","csv"], key="content",
            help="EAN (variant) ↔ Article No (parent) mapping only"
        )
    with mc3:
        override_file = st.file_uploader(
            "⚡ Special Article Override *(optional)*",
            type=["xlsx","xls","csv"], key="override",
            help="Columns: Article_No | Status (ACTIVE/INACTIVE) | Marketplace (col C, blank=ALL). Stock still checked for ACTIVE."
        )

    st.markdown("### Step 3 — Region Files")
    region_files: dict = {}
    for region in selected_regions:
        with st.expander(f"📂 {region}", expanded=True):
            region_files[region] = {}
            c1, c2 = st.columns(2)
            with c1:
                region_files[region]["lazada"] = st.file_uploader(
                    f"Lazada ({region}) — pricestock*.xlsx",
                    type=["xlsx","xls","csv"], key=f"laz_{region}",
                    help="EAN: SellerSKU | ID: ItemId")
                region_files[region]["shopee"] = st.file_uploader(
                    f"Shopee ({region}) — Shopee*Masterfile*.xlsx",
                    type=["xlsx","xls","csv"], key=f"sho_{region}",
                    help="EAN: SKU | ID: Item ID")
                region_files[region]["tiktok"] = st.file_uploader(
                    f"TikTok ({region}) — TikTokSellerCenter*.xlsx",
                    type=["xlsx","xls","csv"], key=f"ttk_{region}",
                    help="EAN: Seller sku | ID: Product ID")
            with c2:
                region_files[region]["zalora_status"] = st.file_uploader(
                    f"Zalora Status ({region}) — SellerStatusTemplate*.xlsx",
                    type=["xlsx","xls","csv"], key=f"zst_{region}",
                    help="EAN: SellerSku")
                region_files[region]["zalora_stock"] = st.file_uploader(
                    f"Zalora Stock ({region}) — SellerStockTemplate*.xlsx",
                    type=["xlsx","xls","csv"], key=f"zsk_{region}")
                region_files[region]["inventory"] = st.file_uploader(
                    f"Inventory ({region})",
                    type=["xlsx","xls","csv"], key=f"inv_{region}",
                    help="Must have: EAN/Barcode column + Avail_Qty column")

    st.markdown("---")
    run_btn = st.button("🚀 Run Listing Audit", type="primary", use_container_width=True)

    if run_btn:
        errs = []
        if not zecom_file:   errs.append("ZeCom Tracker required.")
        if not content_file: errs.append("Content Master required.")
        if not selected_regions: errs.append("Select at least one region.")
        for e in errs: st.error(e)

        if not errs:
            prog = st.progress(0, text="Loading ZeCom…")

            with st.spinner("Loading ZeCom Tracker…"):
                zecom_df = load_zecom(zecom_file)
            st.success(f"✅ ZeCom: **{len(zecom_df):,}** articles")
            prog.progress(10)

            with st.spinner("Loading Content Master…"):
                content_df = load_content(content_file)
            st.success(
                f"✅ Content: **{len(content_df):,}** EANs "
                f"/ **{content_df['Article_No'].nunique():,}** articles"
            )
            prog.progress(20)

            override_df = pd.DataFrame(columns=["Article_No","Status","Marketplace"])
            if override_file:
                with st.spinner("Loading Special Override…"):
                    override_df = load_override(override_file)
                active_ov   = (override_df["Status"]=="ACTIVE").sum()
                inactive_ov = (override_df["Status"]=="INACTIVE").sum()
                st.info(
                    f"⚡ Override: **{len(override_df):,}** articles "
                    f"({active_ov} ACTIVE, {inactive_ov} INACTIVE) — stock still checked for ACTIVE"
                )
            prog.progress(25)

            all_results: dict  = {}
            inv_debug_all: dict = {}
            mp_debug_all: dict  = {}
            step_size = max(1, int(70 / len(selected_regions)))
            step = 25

            for region in selected_regions:
                rf = region_files.get(region, {})
                st.markdown(f"#### 📂 {region}")
                mp_dfs: dict = {}
                mp_dbg: dict = {}

                def load_and_report(name, loader_fn, *args):
                    with st.spinner(f"[{region}] Loading {name}…"):
                        result = loader_fn(*args)
                    n = len(result)
                    ean_col_ok = "EAN" in result.columns and n > 0
                    status_vals = (
                        result["MP_Status_Raw"].value_counts().head(8).to_dict()
                        if "MP_Status_Raw" in result.columns and n > 0 else {}
                    )
                    mp_dbg[name] = {
                        "rows": n,
                        "EAN col found": ean_col_ok,
                        "status values": status_vals,
                        "columns": list(result.columns),
                    }
                    st.write(
                        f"  {name}: **{n:,}** EANs | "
                        f"Status values: `{list(status_vals.keys())[:5]}`"
                    )
                    return result

                if rf.get("lazada"):
                    mp_dfs["Lazada"] = load_and_report("Lazada", load_lazada, rf["lazada"])
                if rf.get("shopee"):
                    mp_dfs["Shopee"] = load_and_report("Shopee", load_shopee, rf["shopee"])
                if rf.get("zalora_status") and rf.get("zalora_stock"):
                    mp_dfs["Zalora"] = load_and_report(
                        "Zalora", load_zalora, rf["zalora_status"], rf["zalora_stock"]
                    )
                elif rf.get("zalora_status"):
                    st.warning(f"[{region}] Zalora Stock file missing — skipped")
                if rf.get("tiktok"):
                    mp_dfs["TikTok"] = load_and_report("TikTok", load_tiktok, rf["tiktok"])

                if not mp_dfs:
                    st.warning(f"[{region}] No MP files — skipping.")
                    continue

                inv_df = pd.DataFrame(columns=["EAN","Inv_Stock"])
                if rf.get("inventory"):
                    with st.spinner(f"[{region}] Loading Inventory…"):
                        inv_df, inv_dbg = load_inventory(rf["inventory"], region)
                    inv_debug_all[region] = inv_dbg
                    non_zero = int((inv_df["Inv_Stock"] > 0).sum())
                    st.write(
                        f"  Inventory: **{len(inv_df):,}** EANs | "
                        f"In-stock: **{non_zero:,}** | "
                        f"Column used: `{inv_dbg.get('Stock column used','?')}`"
                    )
                else:
                    st.warning(f"[{region}] No Inventory file — all stock = 0")

                mp_debug_all[region] = mp_dbg
                prog.progress(step, text=f"[{region}] Running audit…")

                with st.spinner(f"[{region}] Running audit…"):
                    region_result = run_audit(
                        mp_dfs, inv_df, zecom_df, content_df, override_df, region
                    )
                all_results[region] = region_result

                summary_line = "  |  ".join(
                    f"{CAT_STYLE[c]['icon']} {c}: **"
                    f"{sum(len(region_result[mp].get(c,pd.DataFrame())) for mp in MARKETPLACES if mp in region_result):,}**"
                    for c in ALL_CATS
                )
                st.success(f"✅ [{region}] Done — {summary_line}")
                step = min(step + step_size, 95)

            prog.progress(100, text="Complete!")
            st.session_state.audit_results = all_results
            st.session_state.inv_debug     = inv_debug_all
            st.session_state.mp_debug      = mp_debug_all
            if all_results:
                st.success("🎉 Done! Go to **📊 Results & Download** tab.")

# ════ RESULTS ═════════════════════════════════════════════════════════════════
with tab_results:
    results = st.session_state.audit_results
    if not results:
        st.info("Run the audit first.")
    else:
        # KPIs
        st.markdown("### 📊 Overall")
        grand = {cat:0 for cat in ALL_CATS}
        for mp_res in results.values():
            for cats in mp_res.values():
                for cat in ALL_CATS:
                    grand[cat] += len(cats.get(cat, pd.DataFrame()))
        css_map = {CAT_ACTIVE:"cg",CAT_INACTIVE:"cr",CAT_NL_ART:"co",CAT_NL_EAN:"cp"}
        kcols = st.columns(4)
        for i, cat in enumerate(ALL_CATS):
            kcols[i].markdown(
                f"<div class='metric-box'>"
                f"<div class='metric-value {css_map[cat]}'>{grand[cat]:,}</div>"
                f"<div class='metric-label'>{CAT_STYLE[cat]['icon']} {cat}</div>"
                f"</div>", unsafe_allow_html=True)

        # Summary table
        st.markdown("### 📋 Region × Marketplace Breakdown")
        rows = []
        for region, mp_res in results.items():
            for mp in MARKETPLACES:
                if mp not in mp_res: continue
                row = {"Region":region,"Marketplace":mp}
                for cat in ALL_CATS:
                    row[f"{CAT_STYLE[cat]['icon']} {cat}"] = len(
                        mp_res[mp].get(cat, pd.DataFrame()))
                rows.append(row)
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

        # Drilldown
        st.markdown("### 🔍 Drilldown")
        d1,d2,d3 = st.columns(3)
        dr_region = d1.selectbox("Region",      list(results.keys()))
        dr_mp     = d2.selectbox("Marketplace", [
            mp for mp in MARKETPLACES if mp in results.get(dr_region,{})])
        dr_cat    = d3.selectbox("Category",    ALL_CATS)

        df_view = results[dr_region].get(dr_mp,{}).get(dr_cat, pd.DataFrame())
        st.markdown(
            f"**{CAT_STYLE[dr_cat]['icon']} {dr_mp} [{dr_region}] "
            f"— {dr_cat}: {len(df_view):,} records**"
        )
        if not df_view.empty:
            search = st.text_input("🔎 Filter by Article No / EAN", "")
            if search.strip():
                mask = pd.Series(False, index=df_view.index)
                for col in ["Article No","EAN (Seller SKU)"]:
                    if col in df_view.columns:
                        mask |= df_view[col].astype(str).str.contains(
                            search.strip(), case=False, na=False)
                df_view = df_view[mask]
                st.caption(f"{len(df_view):,} filtered records")
            st.dataframe(df_view, use_container_width=True, height=450)
        else:
            st.success("✅ No records in this category.")

        # Download
        st.markdown("---")
        st.markdown("### 💾 Download")
        with st.spinner("Building Excel…"):
            xlsx = build_excel(results)
        fname = f"PUMA_Audit_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        st.download_button(
            "📥 Download Full Audit Report (.xlsx)",
            data=xlsx, file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, type="primary"
        )
        st.caption(
            "Sheets: Summary + one tab per MP × Region × Category | "
            "Columns include MP ID (Lazada ItemId / Shopee Item ID)"
        )

# ════ DEBUG ═══════════════════════════════════════════════════════════════════
with tab_debug:
    st.markdown("### 🔧 Inventory Debug")
    inv_debug = st.session_state.inv_debug
    if inv_debug:
        for region, dbg in inv_debug.items():
            st.markdown(f"**{region}**")
            st.dataframe(
                pd.DataFrame([{"Field":k,"Value":str(v)} for k,v in dbg.items()]),
                use_container_width=True, hide_index=True)
    else:
        st.info("Run audit first.")

    st.markdown("---")
    st.markdown("### 🔧 Marketplace File Debug")
    st.markdown("""
Use this to diagnose why records may not appear in Active/Inactive.
Check that:
1. EAN column was found ✓
2. Status values match expected (see actual values below)
3. If status values are unusual, they will be treated as ACTIVE (blacklist logic)
    """)
    mp_debug = st.session_state.mp_debug
    if mp_debug:
        for region, mp_dbg in mp_debug.items():
            st.markdown(f"**{region}**")
            for mp, info in mp_dbg.items():
                with st.expander(f"{mp} — {info['rows']:,} rows | EAN found: {info['EAN col found']}"):
                    st.write(f"**Columns:** {info['columns']}")
                    st.write(f"**Status values (top 8):** {info['status values']}")
    else:
        st.info("Run audit first.")
