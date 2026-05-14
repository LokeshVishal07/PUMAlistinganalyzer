"""
PUMA Marketplace Listing Audit Analyzer  v7.0
==============================================
Architecture: ZeCom-driven, vectorized, 100k+ EAN capable.

MASTER FILE  (all regions): Content File only
PER-REGION   : ZeCom Tracker, Special Request, Inventory,
               Lazada, Shopee, Zalora (status+stock), TikTok

STATUS RULES
  ACTIVE   = Tracker=YES AND Launch<=today AND Stock>0 AND Special≠Inactive
  INACTIVE = Tracker=NO/OFF OR Stock=0 OR Launch>today OR Special=Inactive

LISTING ELIGIBILITY (can appear on MP)
  Tracker=YES AND Stock>0 AND Launch<=today+30d AND Special≠Inactive

LISTING ANALYSIS (article level, content EANs vs MP EANs)
  All content EANs on MP            → Already Listed
  Some content EANs on MP           → Add Variant
  Zero content EANs on MP           → Full New Listing

STOCK BUFFER: PH × Lazada only → expected = inventory - 1 (min 0)

OUTPUT: 5-sheet Excel
  1. Listing Analysis   2. Status Validation   3. Stock Validation
  4. Missing Variants   5. Summary Dashboard
"""

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="PUMA Listing Audit Analyzer",
    layout="wide", page_icon="📊"
)

st.markdown("""
<style>
.main-header{background:linear-gradient(135deg,#1a1a2e,#0f3460);
  padding:1.8rem;border-radius:12px;text-align:center;color:white;margin-bottom:1.5rem;}
.main-header h1{font-size:1.9rem;font-weight:700;margin:0;}
.main-header p{color:#a0aec0;margin:.3rem 0 0;font-size:.88rem;}
.metric-box{background:white;border-radius:8px;padding:.85rem;
  border:1px solid #e2e8f0;text-align:center;margin-bottom:.5rem;}
.mv{font-size:1.6rem;font-weight:800;}
.ml{font-size:.72rem;color:#718096;margin-top:.1rem;}
.cg{color:#276221;}.cr{color:#9c0006;}.co{color:#7b5800;}.cp{color:#4a235a;}.cb{color:#1a5276;}
div[data-testid="stExpander"] summary{font-weight:600;}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class='main-header'>
  <h1>📊 PUMA Marketplace Listing Audit Analyzer</h1>
  <p>ZeCom-Driven · Vectorized · PH / MY / SG · Lazada / Shopee / Zalora / TikTok</p>
</div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════
REGIONS      = ["PH", "MY", "SG"]
MARKETPLACES = ["Lazada", "Shopee", "Zalora", "TikTok"]
TODAY        = pd.Timestamp.today().normalize()
FUTURE_WINDOW = TODAY + pd.Timedelta(days=30)

# PH Lazada: expected_stock = inventory - 1 (min 0)
STOCK_BUFFER = {("PH","Lazada"): 1}

# Confirmed stock column per region inventory file
# Inventory stock column names — confirmed per region
INVENTORY_STOCK_COL = {
    "PH": "Avail_Qty",        # file prefix: Inventory_
    "MY": "QtyAvailable",     # file prefix: PUMA_MY_B2C_Channel_Inventory_
    "SG": "QTY",              # file prefix: SG_PUMA SG B2C Inventory
}

# Inventory filename prefix → region mapping (for validation/warning)
INVENTORY_FILE_PREFIX = {
    "inventory_"                    : "PH",
    "puma_my_b2c_channel_inventory_": "MY",
    "sg_puma sg b2c inventory"      : "SG",
}

def detect_inventory_region(filename: str) -> str:
    """Detect region from inventory filename. Returns region string or empty."""
    fn = filename.lower().strip()
    for prefix, region in INVENTORY_FILE_PREFIX.items():
        if fn.startswith(prefix):
            return region
    return ""

# Confirmed MP column names
MP_CONFIG = {
    "Lazada": {"ean":"SellerSKU",  "status":"status",     "stock":"Quantity",  "id":"Product ID"},
    "Shopee": {"ean":"SKU",        "status":"Status",      "stock":"Stock",     "id":"Product ID"},
    "Zalora": {"ean":"SellerSku",  "status":"Status",      "stock":"Quantity",  "id":""},
    "TikTok": {"ean":"Seller sku", "status":"Status",      "stock":"Quantity",  "id":"Product ID"},
}

# Known inactive status strings (case-insensitive blacklist)
INACTIVE_STATUSES = {
    "inactive","inactivated","deactivated","deleted","seller_deleted",
    "seller deleted","banned","banned by admin","unlisted","unlist",
    "suspended","blocked","violation","delisted","rejected","failed",
    "prohibited","taken down","not listed","not_listed","no","off","0","false",
}

FILE_TYPES = ["xlsx","xls","xlsm","xlsb","csv","tsv","ods"]

# ══════════════════════════════════════════════════════════════════════════════
# LOW-LEVEL UTILITIES
# ══════════════════════════════════════════════════════════════════════════════

def _ean(v) -> str:
    """Normalize EAN: strip .0 suffix from float-read ints."""
    s = str(v).strip().split(".")[0].strip()
    return s if s not in ("nan","None","NaT","") else ""

def _s(v) -> str:
    s = str(v).strip()
    return s if s not in ("nan","None","NaT","") else ""

def _i(v) -> int:
    """Convert any value to non-negative int safely."""
    try:
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return 0
        return max(0, int(float(str(v).strip().replace(",", ""))))
    except:
        return 0

def _is_active_status(v) -> bool:
    return str(v).strip().lower() not in INACTIVE_STATUSES

def _resolve_sheet(file, wanted: str, engines: list) -> str:
    """Return actual sheet name from file that matches wanted (case-insensitive)."""
    for eng in engines:
        try:
            file.seek(0)
            sheets = pd.ExcelFile(file, engine=eng).sheet_names
            if wanted in sheets:
                return wanted
            match = next((s for s in sheets
                          if s.strip().lower() == wanted.strip().lower()), None)
            return match if match else sheets[0]
        except Exception:
            continue
    return wanted

def read_file(file, sheet_name=0, header=0) -> pd.DataFrame:
    """
    Universal reader: xlsx/xls/xlsm/xlsb/ods/csv/tsv.
    Resolves sheet names case-insensitively. Always returns string-column DF.
    """
    name = getattr(file, "name", "") or ""
    ext  = name.rsplit(".", 1)[-1].lower() if "." in name else ""

    if ext in ("csv","tsv"):
        sep = "\t" if ext == "tsv" else ","
        for enc in ("utf-8","utf-8-sig","latin-1"):
            try:
                file.seek(0)
                df = pd.read_csv(file, sep=sep, header=header,
                                 dtype=str, encoding=enc, low_memory=False)
                df.columns = [str(c).strip() for c in df.columns]
                df = df.astype(object)
                return df.dropna(how="all").reset_index(drop=True)
            except UnicodeDecodeError:
                continue
        return pd.DataFrame()

    engines = {"xlsx":["openpyxl"],"xlsm":["openpyxl"],
               "xls":["xlrd","openpyxl"],"xlsb":["pyxlsb","openpyxl"],
               "ods":["odf","openpyxl"]}.get(ext, ["openpyxl","xlrd"])

    if isinstance(sheet_name, str):
        sheet_name = _resolve_sheet(file, sheet_name, engines)

    for eng in engines:
        try:
            file.seek(0)
            df = pd.read_excel(file, sheet_name=sheet_name,
                               header=header, engine=eng)
            df.columns = [str(c).strip() for c in df.columns]
            df = df.astype(object)
            return df.dropna(how="all").reset_index(drop=True)
        except Exception:
            continue
    try:
        file.seek(0)
        df = pd.read_excel(file, sheet_name=sheet_name, header=header)
        df.columns = [str(c).strip() for c in df.columns]
        df = df.astype(object)
        return df.dropna(how="all").reset_index(drop=True)
    except Exception as e:
        st.error(f"Cannot read '{name}': {e}")
        return pd.DataFrame()

def norm_col(df: pd.DataFrame, candidates: list, new_name: str) -> pd.DataFrame:
    """Rename first matching column (case-insensitive)."""
    cands = {c.strip().lower() for c in candidates}
    for col in df.columns:
        if col.strip().lower() in cands:
            return df.rename(columns={col: new_name}) if col != new_name else df
    return df

# ══════════════════════════════════════════════════════════════════════════════
# FILE LOADERS
# ══════════════════════════════════════════════════════════════════════════════

# ── Content File ──────────────────────────────────────────────────────────────
def load_content(file) -> pd.DataFrame:
    """
    Master mapping: Article_No ↔ EAN (+ Size if available).
    Sheet: 'content'. Cols: Color_No → Article_No, EAN.
    """
    df = read_file(file, sheet_name="content")
    if df.empty:
        file.seek(0); df = read_file(file)

    df = norm_col(df, ["Color_No","ColorNo","Color No","Article No",
                        "ArticleNo","PIM Article#","Style Number"], "Article_No")
    df = norm_col(df, ["EAN","ean","Barcode","GTIN","UPC","Child SKU"], "EAN")

    if "Article_No" not in df.columns or "EAN" not in df.columns:
        st.error("Content: Missing Article_No or EAN column.")
        return pd.DataFrame(columns=["Article_No","EAN","Size"])

    df["Article_No"] = df["Article_No"].apply(_s)
    df["EAN"]        = df["EAN"].apply(_ean)

    size_col = next((c for c in df.columns
                     if "size" in c.lower() and c not in ("Article_No","EAN")), None)
    df["Size"] = df[size_col].apply(_s) if size_col else ""

    df = df[(df["Article_No"] != "") & (df["EAN"] != "")]
    return df[["Article_No","EAN","Size"]].drop_duplicates("EAN").reset_index(drop=True)


# ── ZeCom Tracker ─────────────────────────────────────────────────────────────
def load_zecom(file) -> pd.DataFrame:
    """
    Sheet: 'PH'/'MY'/'SG' (try all, pick first that works with data).
    Header row 2. Cols: PIM Article# | LAZADA | SHOPEE | ZALORA | TIK TOK | Launch Dates
    """
    sheet_tried = None
    df = pd.DataFrame()

    # Try region-named sheets, then generic fallback
    for sh in ["PH","MY","SG","Sheet1","Sheet","Data",0]:
        try:
            file.seek(0)
            tmp = read_file(file, sheet_name=sh, header=2)
            if len(tmp) > 5:
                df = tmp; sheet_tried = sh; break
        except Exception:
            continue

    if df.empty:
        file.seek(0); df = read_file(file, header=2)

    df = norm_col(df, ["PIM Article#","PIM Article","Article No","ArticleNo",
                        "Color_No","ColorNo","Style Number","Article Number"], "Article_No")

    # Tracker columns
    for mp, keys in [("Lazada",["LAZADA","Lazada","lazada"]),
                      ("Shopee",["SHOPEE","Shopee","shopee"]),
                      ("Zalora",["ZALORA","Zalora","zalora"]),
                      ("TikTok",["TIK TOK","TIKTOK","TikTok","tiktok","Tik Tok"])]:
        df = norm_col(df, keys, f"Tracker_{mp}")

    # Launch date
    df = norm_col(df, ["Launch Dates","Launch Date","LaunchDate","Go Live","live date"], "Launch_Date")

    for col in [f"Tracker_{m}" for m in MARKETPLACES] + ["Launch_Date","Article_No"]:
        if col not in df.columns:
            df[col] = np.nan

    df["Article_No"]  = df["Article_No"].apply(_s)
    df["Launch_Date"] = pd.to_datetime(df["Launch_Date"], errors="coerce")
    df = df[df["Article_No"].str.match(r'^\S+.*\S+$', na=False) &
            (df["Article_No"].str.len() > 2)]
    df = df.drop_duplicates("Article_No").reset_index(drop=True)

    cols = ["Article_No","Tracker_Lazada","Tracker_Shopee",
            "Tracker_Zalora","Tracker_TikTok","Launch_Date"]
    return df[cols]


# ── Special Request File ──────────────────────────────────────────────────────
def load_special_request(file) -> pd.DataFrame:
    """
    Per-region. Cols: Article_No | Marketplace | Status (Active/Inactive).
    Status=Inactive → always INACTIVE (overrides everything).
    Status=Active   → ACTIVE only if stock > 0.
    """
    df = read_file(file)
    df = norm_col(df, ["Article No","ArticleNo","Color_No","PIM Article#",
                        "Color No","Style Number"], "Article_No")
    df = norm_col(df, ["Marketplace","marketplace","Channel","Platform"], "Marketplace")
    df = norm_col(df, ["Status","status","Override","Flag"], "Status")

    for col in ["Article_No","Marketplace","Status"]:
        if col not in df.columns:
            df[col] = ""

    df["Article_No"] = df["Article_No"].apply(_s)
    df["Marketplace"]= df["Marketplace"].apply(_s).str.upper()
    df["Status"]     = df["Status"].apply(_s).str.upper()
    df = df[df["Article_No"] != ""]
    df = df[df["Status"].isin(["ACTIVE","INACTIVE"])]
    return df[["Article_No","Marketplace","Status"]].reset_index(drop=True)


# ── Inventory File ────────────────────────────────────────────────────────────
def load_inventory(file, region: str) -> tuple:
    """
    Returns (DataFrame[EAN, Inv_Stock], debug_dict).

    Stock column per region (confirmed):
      PH  → Avail_Qty        (file: Inventory_*)
      MY  → QtyAvailable     (file: PUMA_MY_B2C_Channel_Inventory_*)
      SG  → QTY              (file: SG_PUMA SG B2C Inventory*)

    Also validates filename matches expected region prefix and warns if not.
    Fallback order if primary column not found:
      Avail_Qty → QtyAvailable → QTY → stock per ean → available → qty → stock
    """
    filename = getattr(file, "name", "") or ""

    # ── Filename validation ───────────────────────────────────────────────
    detected_region = detect_inventory_region(filename)
    if detected_region and detected_region != region:
        st.warning(
            f"⚠️ [{region}] Inventory filename '{filename}' looks like a "
            f"**{detected_region}** file (expected prefix for {region}). "
            f"Please verify you uploaded the correct file."
        )
    elif not detected_region and filename:
        st.warning(
            f"[{region}] Inventory filename '{filename}' does not match any known prefix. "
            "Expected — PH: Inventory_*  |  MY: PUMA_MY_B2C_Channel_Inventory_*  |  SG: SG_PUMA SG B2C Inventory*"
        )

    df = read_file(file)
    all_cols = list(df.columns)

    df = norm_col(df, ["EAN","ean","Barcode","barcode","SKU","Material"], "EAN")
    if "EAN" not in df.columns:
        st.warning(f"[{region}] Inventory: EAN column not found. Cols: {all_cols[:15]}")
        return pd.DataFrame(columns=["EAN","Inv_Stock"]), {"error":"EAN not found"}

    df["EAN"] = df["EAN"].apply(_ean)
    df = df[df["EAN"].str.match(r'^\d{5,}$', na=False)]

    # ── Primary stock column (region-specific, confirmed names) ──────────
    primary   = INVENTORY_STOCK_COL.get(region, "Avail_Qty")
    stock_col = None

    # Exact match first
    if primary in df.columns:
        stock_col = primary
    else:
        # Case-insensitive match
        stock_col = next(
            (c for c in df.columns if c.strip().lower() == primary.lower()),
            None
        )

    # ── Fallback chain if primary not found ──────────────────────────────
    if not stock_col:
        # Try the other regions' known columns before generic fallbacks
        region_cols_priority = ["avail_qty", "qtyavailable", "qty",
                                 "stock per ean", "available",
                                 "on hand", "soh", "quantity", "stock"]
        stock_col = next(
            (c for c in df.columns
             if c.strip().lower() in region_cols_priority),
            None
        )
        if stock_col:
            st.warning(
                f"[{region}] Inventory: Expected column '{primary}' not found. "
                f"Using fallback column '{stock_col}'. "
                f"Please verify this is the correct stock column."
            )
        else:
            st.error(
                f"[{region}] Inventory: Cannot find any stock column. "
                f"Expected: '{primary}'. Columns in file: {all_cols}"
            )

    # Convert stock column to numeric FIRST — avoids TypeError on string dtype comparison
    if stock_col and stock_col in df.columns:
        df["Inv_Stock"] = pd.to_numeric(df[stock_col], errors="coerce").fillna(0).clip(lower=0).astype(int)
    else:
        df["Inv_Stock"] = 0

    debug = {
        "Region"               : region,
        "Filename"             : filename,
        "Detected region"      : detected_region or "unknown",
        "Expected stock col"   : primary,
        "Actual stock col used": stock_col or "NOT FOUND (stock=0)",
        "EAN rows"             : len(df),
        "Non-zero stock EANs"  : int((df["Inv_Stock"] > 0).sum()),
        "All columns"          : ", ".join(all_cols),
    }

    result = df[["EAN","Inv_Stock"]].drop_duplicates("EAN").reset_index(drop=True)
    return result, debug


# ── Marketplace Loaders ───────────────────────────────────────────────────────
def _load_mp(file, mp: str, extra_sheet=None) -> pd.DataFrame:
    cfg = MP_CONFIG[mp]
    df  = read_file(file, sheet_name=extra_sheet) if extra_sheet else read_file(file)

    df = norm_col(df, [cfg["ean"]] + ["SellerSKU","SKU","Seller sku","SellerSku"], "EAN")
    df = norm_col(df, [cfg["status"],"Status","status","ItemStatus","Listing Status"], "MP_Status")
    df = norm_col(df, [cfg["stock"],"Stock","Quantity","Available","Qty"], "MP_Stock")

    if cfg["id"]:
        df = norm_col(df, [cfg["id"],"ItemId","item_id","Product ID","ProductId"], "MP_ID")
    if "MP_ID" not in df.columns:
        df["MP_ID"] = ""

    for col in ["EAN","MP_Status","MP_Stock"]:
        if col not in df.columns: df[col] = np.nan

    df["EAN"]       = df["EAN"].apply(_ean)
    df["MP_Stock"]  = df["MP_Stock"].apply(_i)
    df["MP_Status"] = df["MP_Status"].apply(_s)
    df["MP_ID"]     = df["MP_ID"].apply(lambda v: _s(str(v).split(".")[0]))
    df["Marketplace"] = mp

    # Filter to EAN-like rows only (removes header/instruction rows)
    df = df[df["EAN"].str.match(r'^\d{8,}$', na=False)]
    return df[["EAN","MP_Status","MP_Stock","MP_ID","Marketplace"]].drop_duplicates("EAN")


def load_lazada(file)  -> pd.DataFrame: return _load_mp(file, "Lazada", "template")
def load_shopee(file)  -> pd.DataFrame: return _load_mp(file, "Shopee")
def load_tiktok(file)  -> pd.DataFrame: return _load_mp(file, "TikTok")

def load_zalora(sf, stf) -> pd.DataFrame:
    ds   = read_file(sf,  sheet_name="ProductStatuses")
    dstk = read_file(stf, sheet_name="Sheet")
    ds   = norm_col(ds,   ["SellerSku","SellerSKU","Seller SKU"], "EAN")
    ds   = norm_col(ds,   ["Status","status"], "MP_Status")
    dstk = norm_col(dstk, ["SellerSku","SellerSKU","Seller SKU"], "EAN")
    dstk = norm_col(dstk, ["Quantity","Stock","Available","Qty"], "MP_Stock")
    for col in ["EAN","MP_Status"]:
        if col not in ds.columns: ds[col] = np.nan
    if "EAN"      not in dstk.columns: dstk["EAN"]      = np.nan
    if "MP_Stock" not in dstk.columns: dstk["MP_Stock"]  = 0
    ds["EAN"]    = ds["EAN"].apply(_ean)
    dstk["EAN"]  = dstk["EAN"].apply(_ean)
    ds["MP_Status"] = ds["MP_Status"].apply(_s)
    merged = ds.merge(dstk[["EAN","MP_Stock"]].drop_duplicates("EAN"), on="EAN", how="left")
    merged["MP_Stock"]    = merged["MP_Stock"].apply(_i)
    merged["MP_ID"]       = ""
    merged["Marketplace"] = "Zalora"
    merged = merged[merged["EAN"].str.match(r'^\d{8,}$', na=False)]
    return merged[["EAN","MP_Status","MP_Stock","MP_ID","Marketplace"]].drop_duplicates("EAN")


# ══════════════════════════════════════════════════════════════════════════════
# VECTORIZED AUDIT ENGINE
# ══════════════════════════════════════════════════════════════════════════════

def run_audit(
    mp_dfs      : dict,           # {mp: DataFrame}
    inv_df      : pd.DataFrame,   # EAN, Inv_Stock
    zecom_df    : pd.DataFrame,   # Article_No, Tracker_*, Launch_Date
    content_df  : pd.DataFrame,   # Article_No, EAN, Size
    special_df  : pd.DataFrame,   # Article_No, Marketplace, Status
    region      : str,
) -> dict:
    """
    Fully vectorized audit. Returns dict with 4 DataFrames:
      listing_analysis, status_validation, stock_validation, missing_variants
    """

    # ── Build article → EAN sets from content ──────────────────────────────
    art_eans = (content_df.groupby("Article_No")["EAN"]
                .apply(set).reset_index()
                .rename(columns={"EAN":"Content_EANs"}))

    ean_size = content_df.set_index("EAN")["Size"].to_dict()
    ean_art  = content_df.set_index("EAN")["Article_No"].to_dict()

    # ── Inventory index ─────────────────────────────────────────────────────
    inv = inv_df.set_index("EAN")["Inv_Stock"].to_dict() if not inv_df.empty else {}

    # ── Special request index: (Article_No, MP_UPPER) → status ─────────────
    sp_idx: dict = {}
    if not special_df.empty:
        for _, row in special_df.iterrows():
            key = (row["Article_No"], row["Marketplace"])
            sp_idx[key] = row["Status"]
            if row["Marketplace"] == "ALL":
                for mp in MARKETPLACES:
                    sp_idx[(row["Article_No"], mp.upper())] = row["Status"]

    results_listing    = []
    results_status     = []
    results_stock      = []
    results_missing    = []

    for mp in MARKETPLACES:
        mp_df = mp_dfs.get(mp, pd.DataFrame())
        mp_idx: dict = {}   # EAN → row
        if not mp_df.empty:
            mp_idx = mp_df.set_index("EAN").to_dict("index")

        buf = STOCK_BUFFER.get((region, mp), 0)

        # Iterate ZeCom articles (this is the driver)
        for _, z in zecom_df.iterrows():
            art     = z["Article_No"]
            tracker = _s(z.get(f"Tracker_{mp}", ""))
            launch  = z["Launch_Date"]
            t_upper = tracker.upper()

            # ── Tracker decode ──────────────────────────────────────────────
            tracker_active = (t_upper == "YES")
            tracker_inactive = (t_upper in ("NO","OFF"))

            # ── Launch date decode ─────────────────────────────────────────
            has_launch  = pd.notna(launch)
            past_launch = (not has_launch) or (launch <= TODAY)
            future_30   = has_launch and (TODAY < launch <= FUTURE_WINDOW)
            far_future  = has_launch and (launch > FUTURE_WINDOW)

            # ── Special request ─────────────────────────────────────────────
            sp_key = (art, mp.upper())
            sp_val = sp_idx.get(sp_key, "")   # "ACTIVE" | "INACTIVE" | ""

            # ── Content EANs for this article ──────────────────────────────
            row_art = art_eans[art_eans["Article_No"] == art]
            content_eans: set = row_art.iloc[0]["Content_EANs"] if len(row_art) else set()

            # ── Expected status at article level ───────────────────────────
            def get_expected(ean: str) -> tuple:
                """Returns (expected_status, reason)"""
                stock = _i(inv.get(ean, 0))   # _i() ensures int, never string
                eff   = max(0, stock - buf)

                # Special request overrides everything
                if sp_val == "INACTIVE":
                    return "INACTIVE", "Special Request = INACTIVE"
                if sp_val == "ACTIVE":
                    if eff > 0:
                        return "ACTIVE", "Special Request = ACTIVE + stock > 0"
                    return "INACTIVE", "Special Request = ACTIVE but stock = 0"

                if tracker_inactive:
                    return "INACTIVE", f"Tracker = {tracker}"
                if far_future:
                    return "INACTIVE", f"Launch Date {launch.date()} > today+30d"
                if not tracker_active:
                    return "INACTIVE", f"Tracker = '{tracker}' (blank/unknown)"
                if eff <= 0:
                    return "INACTIVE", f"Tracker=YES but stock=0"
                if future_30:
                    return "INACTIVE", f"Tracker=YES, stock={eff}, but launch {launch.date()} is future (listable)"
                return "ACTIVE", f"Tracker=YES, stock={eff}, launch OK"

            # ── Listing Analysis (article level) ───────────────────────────
            listed_eans   = {e for e in content_eans if e in mp_idx}
            missing_eans  = content_eans - listed_eans
            n_content     = len(content_eans)
            n_listed      = len(listed_eans)
            n_missing     = len(missing_eans)

            if n_content == 0:
                listing_action = "No Content EANs"
            elif n_listed == 0:
                listing_action = "Full New Listing"
            elif n_missing == 0:
                listing_action = "Already Listed"
            else:
                listing_action = "Add Variant"

            results_listing.append({
                "Region"         : region,
                "Marketplace"    : mp,
                "Article No"     : art,
                "Tracker Status" : tracker if tracker else "blank",
                "Launch Date"    : launch.date() if pd.notna(launch) else "—",
                "Total EANs"     : n_content,
                "Listed EANs"    : n_listed,
                "Missing EANs"   : n_missing,
                "Listing Action" : listing_action,
            })

            # Missing variants detail
            for ean in sorted(missing_eans):
                results_missing.append({
                    "Region"      : region,
                    "Marketplace" : mp,
                    "Article No"  : art,
                    "Missing EAN" : ean,
                    "Size"        : ean_size.get(ean,""),
                })

            # ── Status & Stock Validation (EAN level, listed EANs only) ───
            for ean in listed_eans:
                mp_row   = mp_idx[ean]
                mp_st    = mp_row["MP_Status"]
                mp_stk   = mp_row["MP_Stock"]
                mp_id    = mp_row.get("MP_ID","")
                inv_stk  = _i(inv.get(ean, 0))   # ensure int
                eff_stk  = max(0, inv_stk - buf)
                exp_stk  = eff_stk

                exp_status, reason = get_expected(ean)
                act_status = "ACTIVE" if _is_active_status(mp_st) else "INACTIVE"

                # Status validation
                status_result = "✓ PASS" if exp_status == act_status else "✗ FAIL"
                results_status.append({
                    "Region"          : region,
                    "Marketplace"     : mp,
                    "Article No"      : art,
                    "EAN"             : ean,
                    "Size"            : ean_size.get(ean,""),
                    "MP ID"           : mp_id,
                    "Expected Status" : exp_status,
                    "Actual Status"   : act_status,
                    "MP Status Raw"   : mp_st,
                    "Result"          : status_result,
                    "Reason"          : reason,
                })

                # Stock validation
                stock_result = "✓ PASS" if exp_stk == mp_stk else "✗ FAIL"
                results_stock.append({
                    "Region"          : region,
                    "Marketplace"     : mp,
                    "Article No"      : art,
                    "EAN"             : ean,
                    "Size"            : ean_size.get(ean,""),
                    "MP ID"           : mp_id,
                    "Inventory Stock" : inv_stk,
                    "Buffer Applied"  : buf,
                    "Expected Stock"  : exp_stk,
                    "Marketplace Stock": mp_stk,
                    "Result"          : stock_result,
                    "Discrepancy"     : mp_stk - exp_stk,
                })

    return {
        "listing_analysis" : pd.DataFrame(results_listing),
        "status_validation": pd.DataFrame(results_status),
        "stock_validation" : pd.DataFrame(results_stock),
        "missing_variants" : pd.DataFrame(results_missing),
    }


# ══════════════════════════════════════════════════════════════════════════════
# EXCEL EXPORT
# ══════════════════════════════════════════════════════════════════════════════

LISTING_COLORS = {
    "Full New Listing": ("#1a5276","#d6eaf8"),
    "Add Variant"     : ("#7b5800","#ffeb9c"),
    "Already Listed"  : ("#1e6f39","#c6efce"),
    "No Content EANs" : ("#636363","#f2f2f2"),
}

def build_excel(all_results: dict, regions_run: list) -> bytes:
    out = BytesIO()

    # Combine all regions
    la = pd.concat([r["listing_analysis"]  for r in all_results.values()], ignore_index=True)
    sv = pd.concat([r["status_validation"] for r in all_results.values()], ignore_index=True)
    stk= pd.concat([r["stock_validation"]  for r in all_results.values()], ignore_index=True)
    mv = pd.concat([r["missing_variants"]  for r in all_results.values()], ignore_index=True)

    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        wb = writer.book

        def F(**kw):
            d = {"font_name":"Arial","font_size":9,"border":1,
                 "valign":"vcenter","align":"left"}
            d.update(kw)
            return wb.add_format(d)

        norm    = F()
        bold    = F(bold=True)
        ttl     = F(bold=True, font_size=13, font_color="#0f3460", border=0)
        sub     = F(italic=True, font_size=8, font_color="#718096", border=0)
        pass_f  = F(bg_color="#c6efce", font_color="#276221")
        fail_f  = F(bg_color="#ffc7ce", font_color="#9c0006")
        hdr_f   = F(bold=True, bg_color="#0f3460", font_color="#ffffff",
                    align="center", text_wrap=True)
        num_f   = F(align="center")

        def write_sheet(ws, df, title, col_widths=None):
            ws.write(0, 0, title, ttl)
            ws.write(1, 0, f"Records: {len(df):,}", sub)
            if df.empty:
                ws.write(2, 0, "No data.", sub)
                return
            for ci, col in enumerate(df.columns):
                ws.write(2, ci, col, hdr_f)
                w = col_widths.get(col, 18) if col_widths else 18
                ws.set_column(ci, ci, w)
            ws.freeze_panes(3, 0)
            return df

        # ── Sheet 1: Listing Analysis ──────────────────────────────────────
        ws1 = wb.add_worksheet("1. Listing Analysis")
        writer.sheets["1. Listing Analysis"] = ws1
        ws1.write(0, 0, "📋 Listing Analysis", ttl)
        ws1.write(1, 0, f"Records: {len(la):,}", sub)

        la_cols = list(la.columns)
        widths_la = {"Region":8,"Marketplace":12,"Article No":16,
                     "Tracker Status":14,"Launch Date":14,"Total EANs":11,
                     "Listed EANs":11,"Missing EANs":12,"Listing Action":20}
        for ci, col in enumerate(la_cols):
            ws1.write(2, ci, col, hdr_f)
            ws1.set_column(ci, ci, widths_la.get(col, 15))
        ws1.freeze_panes(3, 0)

        action_col = la_cols.index("Listing Action") if "Listing Action" in la_cols else -1
        for ri, (_, row) in enumerate(la.iterrows()):
            action = str(row.get("Listing Action",""))
            colors = LISTING_COLORS.get(action, ("#000000","#ffffff"))
            af = F(bg_color=colors[1], font_color=colors[0])
            for ci, col in enumerate(la_cols):
                v = row[col]; sv2 = "" if (isinstance(v,float) and np.isnan(v)) else str(v)
                f2 = af if ci == action_col else norm
                ws1.write(ri+3, ci, sv2, f2)

        # ── Sheet 2: Status Validation ─────────────────────────────────────
        ws2 = wb.add_worksheet("2. Status Validation")
        writer.sheets["2. Status Validation"] = ws2
        ws2.write(0, 0, "✅ Status Validation", ttl)
        ws2.write(1, 0, f"Records: {len(sv):,}  |  "
                        f"PASS: {(sv['Result']=='✓ PASS').sum():,}  |  "
                        f"FAIL: {(sv['Result']=='✗ FAIL').sum():,}", sub)

        sv_cols = list(sv.columns)
        widths_sv = {"Region":8,"Marketplace":12,"Article No":16,"EAN":16,
                     "Size":8,"MP ID":14,"Expected Status":16,"Actual Status":14,
                     "MP Status Raw":14,"Result":10,"Reason":42}
        for ci, col in enumerate(sv_cols):
            ws2.write(2, ci, col, hdr_f)
            ws2.set_column(ci, ci, widths_sv.get(col, 15))
        ws2.freeze_panes(3, 0)

        res_col = sv_cols.index("Result") if "Result" in sv_cols else -1
        for ri, (_, row) in enumerate(sv.iterrows()):
            res = str(row.get("Result",""))
            rf  = pass_f if res == "✓ PASS" else fail_f
            for ci, col in enumerate(sv_cols):
                v = row[col]; sv2 = "" if (isinstance(v,float) and np.isnan(v)) else str(v)
                f2 = rf if ci == res_col else norm
                ws2.write(ri+3, ci, sv2, f2)

        # ── Sheet 3: Stock Validation ──────────────────────────────────────
        ws3 = wb.add_worksheet("3. Stock Validation")
        writer.sheets["3. Stock Validation"] = ws3
        ws3.write(0, 0, "📦 Stock Validation", ttl)
        ws3.write(1, 0, f"Records: {len(stk):,}  |  "
                        f"PASS: {(stk['Result']=='✓ PASS').sum():,}  |  "
                        f"FAIL: {(stk['Result']=='✗ FAIL').sum():,}", sub)

        stk_cols = list(stk.columns)
        widths_stk = {"Region":8,"Marketplace":12,"Article No":16,"EAN":16,
                      "Size":8,"MP ID":14,"Inventory Stock":16,"Buffer Applied":14,
                      "Expected Stock":14,"Marketplace Stock":16,"Result":10,"Discrepancy":12}
        for ci, col in enumerate(stk_cols):
            ws3.write(2, ci, col, hdr_f)
            ws3.set_column(ci, ci, widths_stk.get(col, 15))
        ws3.freeze_panes(3, 0)

        res_col3 = stk_cols.index("Result") if "Result" in stk_cols else -1
        for ri, (_, row) in enumerate(stk.iterrows()):
            res = str(row.get("Result",""))
            rf  = pass_f if res == "✓ PASS" else fail_f
            for ci, col in enumerate(stk_cols):
                v = row[col]; sv2 = "" if (isinstance(v,float) and np.isnan(v)) else str(v)
                f2 = rf if ci == res_col3 else norm
                ws3.write(ri+3, ci, sv2, f2)

        # ── Sheet 4: Missing Variants ──────────────────────────────────────
        ws4 = wb.add_worksheet("4. Missing Variants")
        writer.sheets["4. Missing Variants"] = ws4
        ws4.write(0, 0, "🟣 Missing Variants (Not Listed EANs)", ttl)
        ws4.write(1, 0, f"Records: {len(mv):,}", sub)
        mv_cols = list(mv.columns)
        widths_mv = {"Region":8,"Marketplace":12,"Article No":16,
                     "Missing EAN":18,"Size":8}
        for ci, col in enumerate(mv_cols):
            ws4.write(2, ci, col, hdr_f)
            ws4.set_column(ci, ci, widths_mv.get(col, 15))
        ws4.freeze_panes(3, 0)
        mv_f = F(bg_color="#e8d5f5", font_color="#4a235a")
        for ri, (_, row) in enumerate(mv.iterrows()):
            for ci, col in enumerate(mv_cols):
                v = row[col]; sv2 = "" if (isinstance(v,float) and np.isnan(v)) else str(v)
                ws4.write(ri+3, ci, sv2, mv_f)

        # ── Sheet 5: Summary Dashboard ─────────────────────────────────────
        ws5 = wb.add_worksheet("5. Summary Dashboard")
        writer.sheets["5. Summary Dashboard"] = ws5
        ws5.set_column("A:A", 22); ws5.set_column("B:H", 16)
        ws5.write(0, 0, "📊 Summary Dashboard", ttl)
        ws5.write(1, 0, f"Generated: {datetime.now().strftime('%d %b %Y %H:%M')}", sub)

        sum_hdrs = ["Region","Marketplace","Total Articles","Total EANs (Content)",
                    "Active EANs","Inactive EANs","Full New Listing",
                    "Add Variant","Already Listed",
                    "Status PASS","Status FAIL","Stock PASS","Stock FAIL"]
        for ci, h in enumerate(sum_hdrs):
            ws5.write(3, ci, h, hdr_f)
            ws5.set_column(ci, ci, 20)
        ws5.set_row(3, 28)

        r = 4
        for region in regions_run:
            for mp in MARKETPLACES:
                la_rm  = la[(la["Region"]==region)&(la["Marketplace"]==mp)]
                sv_rm  = sv[(sv["Region"]==region)&(sv["Marketplace"]==mp)]
                stk_rm = stk[(stk["Region"]==region)&(stk["Marketplace"]==mp)]

                n_arts  = la_rm["Article No"].nunique()
                n_eans  = la_rm["Total EANs"].sum()
                n_act   = (sv_rm["Expected Status"]=="ACTIVE").sum()
                n_ina   = (sv_rm["Expected Status"]=="INACTIVE").sum()
                n_fnl   = (la_rm["Listing Action"]=="Full New Listing").sum()
                n_av    = (la_rm["Listing Action"]=="Add Variant").sum()
                n_al    = (la_rm["Listing Action"]=="Already Listed").sum()
                n_sp    = (sv_rm["Result"]=="✓ PASS").sum()
                n_sf    = (sv_rm["Result"]=="✗ FAIL").sum()
                n_tkp   = (stk_rm["Result"]=="✓ PASS").sum()
                n_tkf   = (stk_rm["Result"]=="✗ FAIL").sum()

                row_vals = [region,mp,n_arts,int(n_eans),
                            int(n_act),int(n_ina),int(n_fnl),
                            int(n_av),int(n_al),
                            int(n_sp),int(n_sf),int(n_tkp),int(n_tkf)]

                for ci, v in enumerate(row_vals):
                    fmt = num_f if isinstance(v,int) else norm
                    ws5.write(r, ci, v, fmt)
                r += 1

    return out.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════════════════════════
for k in ("audit_results","inv_debug","run_log"):
    if k not in st.session_state:
        st.session_state[k] = {} if k != "run_log" else []

# ══════════════════════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════════════════════
tab_upload, tab_results, tab_debug, tab_help = st.tabs([
    "📁 Upload & Run", "📊 Results & Download", "🔧 Debug", "❓ Help"
])

# ─── HELP ────────────────────────────────────────────────────────────────────
with tab_help:
    st.markdown("""
## ✅ Status Logic

| Condition | Expected Status |
|---|---|
| Tracker=YES + Past/Today Launch + Stock>0 + No Override | 🟢 ACTIVE |
| Tracker=NO or OFF | 🔴 INACTIVE |
| Inventory Stock = 0 | 🔴 INACTIVE |
| Launch Date > today + 30 days | 🔴 INACTIVE |
| Launch Date within 30 days (future) | 🔴 INACTIVE (listable, not active yet) |
| Special Request = INACTIVE | 🔴 INACTIVE (overrides everything) |
| Special Request = ACTIVE + stock > 0 | 🟢 ACTIVE |
| Special Request = ACTIVE + stock = 0 | 🔴 INACTIVE |

## 📋 Listing Analysis Logic

| Situation | Action |
|---|---|
| Zero content EANs on MP | Full New Listing |
| Some content EANs on MP, some missing | Add Variant |
| All content EANs on MP | Already Listed |

## 📁 File Guide

| File | Scope | Key Columns |
|---|---|---|
| Content Master | All regions | Color_No (Article), EAN, Size |
| ZeCom Tracker | Per region | PIM Article#, LAZADA, SHOPEE, ZALORA, TIK TOK, Launch Dates |
| Special Request | Per region | Article No, Marketplace, Status (Active/Inactive) |
| Inventory PH (`Inventory_*`) | Per region | EAN, **Avail_Qty** |
| Inventory MY (`PUMA_MY_B2C_Channel_Inventory_*`) | Per region | EAN, **QtyAvailable** |
| Inventory SG (`SG_PUMA SG B2C Inventory*`) | Per region | EAN, **QTY** |
| Lazada | Per region | Sheet=template, SellerSKU, status, Quantity, Product ID |
| Shopee | Per region | SKU, Status, Stock, Product ID |
| Zalora | Per region (2 files) | SellerSku, Status / Quantity |
| TikTok | Per region | Seller sku, Status, Quantity, Product ID |

## 📦 Stock Buffer
- **PH × Lazada only**: Expected Stock = Inventory − 1 (min 0)
- All other region × marketplace: use inventory directly
    """)

# ─── UPLOAD & RUN ─────────────────────────────────────────────────────────────
with tab_upload:

    st.markdown("### Step 1 — Select Regions")
    selected_regions = st.multiselect(
        "Regions to audit:", REGIONS, default=["PH"],
        help="Each region requires its own ZeCom, Special Request, Inventory, and MP files."
    )

    st.markdown("### Step 2 — Master File *(shared across all regions)*")
    content_file = st.file_uploader(
        "📦 Content Master File **[required]**",
        type=FILE_TYPES, key="content",
        help="Sheet: content | Cols: Color_No (Article No), EAN, Size No."
    )

    st.markdown("### Step 3 — Region Files")
    region_files: dict = {}

    for region in selected_regions:
        with st.expander(f"📂 {region} Files", expanded=True):
            region_files[region] = {}
            st.markdown(f"**{region} — ZeCom & Override**")
            c1, c2 = st.columns(2)
            with c1:
                region_files[region]["zecom"] = st.file_uploader(
                    f"📋 ZeCom Tracker ({region}) **[required]**",
                    type=FILE_TYPES, key=f"zec_{region}",
                    help="PH_MP_eCOM_Tracking_File | Sheet: PH/MY/SG | Header row 2"
                )
            with c2:
                region_files[region]["special"] = st.file_uploader(
                    f"⚡ Special Request ({region}) *(optional)*",
                    type=FILE_TYPES, key=f"sp_{region}",
                    help="Cols: Article No | Marketplace | Status (Active/Inactive)"
                )

            st.markdown(f"**{region} — Inventory**")
            region_files[region]["inventory"] = st.file_uploader(
                f"📦 Inventory ({region})",
                type=FILE_TYPES, key=f"inv_{region}",
                help=(
                    "PH → Inventory_* → col: Avail_Qty  |  "
                    "MY → PUMA_MY_B2C_Channel_Inventory_* → col: QtyAvailable  |  "
                    "SG → SG_PUMA SG B2C Inventory* → col: QTY"
                )
            )

            st.markdown(f"**{region} — Marketplace Files**")
            mc1, mc2 = st.columns(2)
            with mc1:
                region_files[region]["lazada"] = st.file_uploader(
                    f"Lazada ({region}) — pricestock*.xlsx",
                    type=FILE_TYPES, key=f"laz_{region}",
                    help="Sheet: template | EAN: SellerSKU | Status: status | Stock: Quantity"
                )
                region_files[region]["shopee"] = st.file_uploader(
                    f"Shopee ({region}) — Shopee*.xlsx",
                    type=FILE_TYPES, key=f"sho_{region}",
                    help="EAN: SKU | Status: Status | Stock: Stock"
                )
            with mc2:
                region_files[region]["zalora_status"] = st.file_uploader(
                    f"Zalora Status ({region}) — SellerStatusTemplate*.xlsx",
                    type=FILE_TYPES, key=f"zst_{region}",
                    help="Sheet: ProductStatuses | EAN: SellerSku | Status: Status"
                )
                region_files[region]["zalora_stock"] = st.file_uploader(
                    f"Zalora Stock ({region}) — SellerStockTemplate*.xlsx",
                    type=FILE_TYPES, key=f"zsk_{region}",
                    help="Sheet: Sheet | EAN: SellerSku | Stock: Quantity"
                )
                region_files[region]["tiktok"] = st.file_uploader(
                    f"TikTok ({region}) — Tiktoksellercenter_batchedit*.xlsx",
                    type=FILE_TYPES, key=f"ttk_{region}",
                    help="EAN: Seller sku | Status: Status | Stock: Quantity"
                )

    st.markdown("---")
    run_btn = st.button("🚀 Run Audit", type="primary", width='stretch')

    if run_btn:
        errors = []
        if not content_file:           errors.append("Content Master file is required.")
        if not selected_regions:       errors.append("Select at least one region.")
        for rg in selected_regions:
            if not region_files.get(rg,{}).get("zecom"):
                errors.append(f"[{rg}] ZeCom Tracker is required.")

        for e in errors: st.error(e)

        if not errors:
            prog = st.progress(0, text="Loading Content Master…")
            run_log = []

            # Load Content (master, once)
            with st.spinner("Loading Content Master…"):
                content_df = load_content(content_file)
            n_eans = len(content_df)
            n_arts = content_df["Article_No"].nunique()
            st.success(f"✅ Content: **{n_eans:,}** EANs / **{n_arts:,}** articles")
            run_log.append(f"Content: {n_eans:,} EANs, {n_arts:,} articles")
            prog.progress(10)

            all_results: dict  = {}
            inv_debug_all: dict = {}
            step = 10
            step_sz = max(1, int(85 / max(len(selected_regions),1)))

            for region in selected_regions:
                rf = region_files.get(region, {})
                st.markdown(f"#### 📂 {region}")

                # ZeCom
                with st.spinner(f"[{region}] Loading ZeCom…"):
                    zecom_df = load_zecom(rf["zecom"])
                st.write(f"  ZeCom: **{len(zecom_df):,}** articles | "
                         f"Lazada YES: {(zecom_df['Tracker_Lazada'].str.upper()=='YES').sum():,}")
                run_log.append(f"[{region}] ZeCom: {len(zecom_df):,} articles")

                # Special Request
                special_df = pd.DataFrame(columns=["Article_No","Marketplace","Status"])
                if rf.get("special"):
                    with st.spinner(f"[{region}] Loading Special Request…"):
                        special_df = load_special_request(rf["special"])
                    st.write(f"  Special Request: **{len(special_df):,}** overrides "
                             f"({(special_df['Status']=='INACTIVE').sum():,} INACTIVE, "
                             f"{(special_df['Status']=='ACTIVE').sum():,} ACTIVE)")

                # Inventory
                inv_df = pd.DataFrame(columns=["EAN","Inv_Stock"])
                if rf.get("inventory"):
                    with st.spinner(f"[{region}] Loading Inventory…"):
                        inv_df, inv_dbg = load_inventory(rf["inventory"], region)
                    inv_debug_all[region] = inv_dbg
                    st.write(f"  Inventory: **{len(inv_df):,}** EANs | "
                             f"Column: `{inv_dbg.get('Actual stock col used','?')}` | "
                             f"In-stock: **{inv_dbg.get('Non-zero stock EANs',0):,}**")
                else:
                    st.warning(f"[{region}] No Inventory file — all stock = 0")

                # Marketplace files
                mp_dfs: dict = {}

                if rf.get("lazada"):
                    with st.spinner(f"[{region}] Lazada…"):
                        tmp = load_lazada(rf["lazada"])
                    st.write(f"  Lazada: **{len(tmp):,}** EANs | "
                             f"Status: `{dict(tmp['MP_Status'].value_counts().head(3))}`")
                    mp_dfs["Lazada"] = tmp

                if rf.get("shopee"):
                    with st.spinner(f"[{region}] Shopee…"):
                        tmp = load_shopee(rf["shopee"])
                    st.write(f"  Shopee: **{len(tmp):,}** EANs | "
                             f"Status: `{dict(tmp['MP_Status'].value_counts().head(3))}`")
                    mp_dfs["Shopee"] = tmp

                if rf.get("zalora_status") and rf.get("zalora_stock"):
                    with st.spinner(f"[{region}] Zalora…"):
                        tmp = load_zalora(rf["zalora_status"], rf["zalora_stock"])
                    st.write(f"  Zalora: **{len(tmp):,}** EANs | "
                             f"Status: `{dict(tmp['MP_Status'].value_counts().head(3))}`")
                    mp_dfs["Zalora"] = tmp
                elif rf.get("zalora_status"):
                    st.warning(f"[{region}] Zalora Stock file missing — Zalora skipped")

                if rf.get("tiktok"):
                    with st.spinner(f"[{region}] TikTok…"):
                        tmp = load_tiktok(rf["tiktok"])
                    st.write(f"  TikTok: **{len(tmp):,}** EANs")
                    mp_dfs["TikTok"] = tmp

                if not mp_dfs:
                    st.warning(f"[{region}] No marketplace files — skipping region.")
                    step += step_sz
                    prog.progress(min(step,95))
                    continue

                prog.progress(min(step + step_sz//2, 95), text=f"[{region}] Running audit…")

                with st.spinner(f"[{region}] Running audit engine…"):
                    region_result = run_audit(
                        mp_dfs, inv_df, zecom_df, content_df, special_df, region
                    )

                all_results[region] = region_result

                la = region_result["listing_analysis"]
                sv = region_result["status_validation"]
                st.success(
                    f"✅ [{region}] Done — "
                    f"Full New: **{(la['Listing Action']=='Full New Listing').sum():,}** | "
                    f"Add Variant: **{(la['Listing Action']=='Add Variant').sum():,}** | "
                    f"Already Listed: **{(la['Listing Action']=='Already Listed').sum():,}** | "
                    f"Status FAIL: **{(sv['Result']=='✗ FAIL').sum():,}**"
                )
                step += step_sz
                prog.progress(min(step, 95))

            prog.progress(100, text="Complete!")
            st.session_state.audit_results  = all_results
            st.session_state.inv_debug      = inv_debug_all
            st.session_state.run_log        = run_log
            st.session_state.regions_run    = selected_regions

            if all_results:
                st.success("🎉 Audit complete! Go to **📊 Results & Download** tab.")

# ─── RESULTS ─────────────────────────────────────────────────────────────────
with tab_results:
    results = st.session_state.audit_results
    if not results:
        st.info("Run the audit first from the Upload & Run tab.")
    else:
        regions_run = st.session_state.get("regions_run", list(results.keys()))

        # Combine all
        la  = pd.concat([r["listing_analysis"]  for r in results.values()], ignore_index=True)
        sv  = pd.concat([r["status_validation"] for r in results.values()], ignore_index=True)
        stk = pd.concat([r["stock_validation"]  for r in results.values()], ignore_index=True)
        mv  = pd.concat([r["missing_variants"]  for r in results.values()], ignore_index=True)

        # KPIs
        st.markdown("### 📊 Overall Summary")
        k = st.columns(8)
        kpis = [
            ("Total Articles", la["Article No"].nunique(), "cb"),
            ("Total Listed EANs", len(sv), "cb"),
            ("Full New Listing", (la["Listing Action"]=="Full New Listing").sum(), "co"),
            ("Add Variant", (la["Listing Action"]=="Add Variant").sum(), "cp"),
            ("Already Listed", (la["Listing Action"]=="Already Listed").sum(), "cg"),
            ("Active EANs", (sv["Expected Status"]=="ACTIVE").sum(), "cg"),
            ("Status Failures", (sv["Result"]=="✗ FAIL").sum(), "cr"),
            ("Stock Failures", (stk["Result"]=="✗ FAIL").sum(), "cr"),
        ]
        for i,(label,val,css) in enumerate(kpis):
            k[i].markdown(
                f"<div class='metric-box'>"
                f"<div class='mv {css}'>{int(val):,}</div>"
                f"<div class='ml'>{label}</div></div>",
                unsafe_allow_html=True)

        # Breakdown table
        st.markdown("### 📋 Region × Marketplace Breakdown")
        brows = []
        for region in regions_run:
            for mp in MARKETPLACES:
                la_rm  = la[(la["Region"]==region)&(la["Marketplace"]==mp)]
                sv_rm  = sv[(sv["Region"]==region)&(sv["Marketplace"]==mp)]
                stk_rm = stk[(stk["Region"]==region)&(stk["Marketplace"]==mp)]
                brows.append({
                    "Region":region,"Marketplace":mp,
                    "Articles":la_rm["Article No"].nunique(),
                    "Full New Listing":(la_rm["Listing Action"]=="Full New Listing").sum(),
                    "Add Variant":(la_rm["Listing Action"]=="Add Variant").sum(),
                    "Already Listed":(la_rm["Listing Action"]=="Already Listed").sum(),
                    "Active":(sv_rm["Expected Status"]=="ACTIVE").sum(),
                    "Inactive":(sv_rm["Expected Status"]=="INACTIVE").sum(),
                    "Status PASS":(sv_rm["Result"]=="✓ PASS").sum(),
                    "Status FAIL":(sv_rm["Result"]=="✗ FAIL").sum(),
                    "Stock FAIL":(stk_rm["Result"]=="✗ FAIL").sum(),
                })
        st.dataframe(pd.DataFrame(brows), width='stretch', hide_index=True)

        # Drilldown
        st.markdown("### 🔍 Drilldown")
        t1, t2, t3, t4 = st.tabs([
            "📋 Listing Analysis","✅ Status Validation",
            "📦 Stock Validation","🟣 Missing Variants"
        ])

        def drilldown_filters(key_prefix, df):
            c1,c2,c3 = st.columns(3)
            regions_opts = df["Region"].unique().tolist()
            mp_opts      = df["Marketplace"].unique().tolist()
            dr = c1.multiselect("Region",      regions_opts, default=regions_opts, key=f"{key_prefix}_r")
            dm = c2.multiselect("Marketplace", mp_opts,      default=mp_opts,      key=f"{key_prefix}_m")
            search = c3.text_input("🔎 Search Article/EAN", "", key=f"{key_prefix}_s")
            out = df[df["Region"].isin(dr) & df["Marketplace"].isin(dm)]
            if search.strip():
                mask = pd.Series(False, index=out.index)
                for col in ["Article No","EAN","Missing EAN"]:
                    if col in out.columns:
                        mask |= out[col].astype(str).str.contains(search.strip(), case=False, na=False)
                out = out[mask]
            return out

        with t1:
            fla = drilldown_filters("la", la)
            st.caption(f"{len(fla):,} records")
            st.dataframe(fla, width='stretch', height=400)

        with t2:
            fsv = drilldown_filters("sv", sv)
            # Quick filter
            res_filter = st.multiselect("Result Filter", ["✓ PASS","✗ FAIL"],
                                        default=["✓ PASS","✗ FAIL"], key="sv_res")
            fsv = fsv[fsv["Result"].isin(res_filter)]
            st.caption(f"{len(fsv):,} records")
            st.dataframe(fsv, width='stretch', height=400)

        with t3:
            fstk = drilldown_filters("stk", stk)
            res_filter2 = st.multiselect("Result Filter", ["✓ PASS","✗ FAIL"],
                                         default=["✓ PASS","✗ FAIL"], key="stk_res")
            fstk = fstk[fstk["Result"].isin(res_filter2)]
            st.caption(f"{len(fstk):,} records")
            st.dataframe(fstk, width='stretch', height=400)

        with t4:
            fmv = drilldown_filters("mv", mv)
            st.caption(f"{len(fmv):,} missing variants")
            st.dataframe(fmv, width='stretch', height=400)

        # Download
        st.markdown("---")
        st.markdown("### 💾 Download Full Report")
        with st.spinner("Building Excel report…"):
            xlsx = build_excel(results, regions_run)
        fname = f"PUMA_Listing_Audit_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        st.download_button(
            "📥 Download Audit Report (.xlsx)",
            data=xlsx, file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width='stretch', type="primary"
        )
        st.caption(
            "5-sheet Excel: "
            "1. Listing Analysis  |  2. Status Validation  |  "
            "3. Stock Validation  |  4. Missing Variants  |  5. Summary Dashboard"
        )

# ─── DEBUG ───────────────────────────────────────────────────────────────────
with tab_debug:
    st.markdown("### 🔧 Inventory Column Debug")
    dbg = st.session_state.inv_debug
    if dbg:
        for region, info in dbg.items():
            st.markdown(f"**{region}**")
            st.dataframe(
                pd.DataFrame([{"Field":k,"Value":str(v)} for k,v in info.items()]), width='stretch', hide_index=True)
    else:
        st.info("Run audit first.")

    st.markdown("### 📋 Run Log")
    log = st.session_state.run_log
    if log:
        for line in log:
            st.text(line)
    else:
        st.info("Run audit first.")
