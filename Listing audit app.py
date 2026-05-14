"""
PUMA Listing Audit Analyzer  v6.0
===================================
Built from actual file inspection. All column names are exact.

FILE STRUCTURES (confirmed from real files):
─────────────────────────────────────────────
ZeCom Tracker  (PH_MP_eCOM_Tracking_File)
  Sheet     : "PH"  |  Header row: 2
  Article   : "PIM Article#"   (format: 531103_03)
  Lazada    : "LAZADA"         values: YES / NO / OFF
  Shopee    : "SHOPEE"         values: YES / NO / OFF
  Zalora    : "ZALORA"         values: YES / NO / OFF
  TikTok    : "TIK TOK"        values: YES / NO / OFF
  Launch    : "Launch Dates"   (datetime)

Content File  (Content_file)
  Sheet     : "content"  |  Header row: 0
  EAN       : "EAN"            (int64 → str)
  Article   : "Color_No"       (format: 352634_03)
  Size      : "Size No."

Special Override  (6_5_Tracker_PH_Remarks)
  Sheet     : "Sheet1"  |  Header row: 0
  Article   : "Color No"
  EAN       : "EAN"
  Lazada ID : "Lazada Status"   ← these are Product IDs, NOT status words
  Shopee ID : "Shopee Status"   ← Product IDs
  Zalora ID : "Zalora Status"   ← ShopSku / Config SKU
  NOTE: This file is an EAN-level override list. EANs in this file
        should be treated as ACTIVE on that marketplace (if stock > 0).
        EANs NOT in this file follow normal ZeCom rules.

Inventory  (Inventory_YYYYMMDD)
  Sheet     : first sheet  |  Header row: 0
  EAN       : "EAN"            (int64)
  Stock     : "STOCK Per EAN"  ← confirmed stock column
  PH Lazada buffer: effective = STOCK Per EAN - 1 (min 0)
  All other channels: STOCK Per EAN directly

Lazada  (pricestock*)
  Sheet     : "template"  |  Header row: 0
  EAN       : "SellerSKU"      (rows 0-2 are instructions, filter by numeric EAN)
  Status    : "status"         values: active / inactive
  Stock     : "Quantity"
  ID        : "Product ID"

Shopee  (Shopee*Masterfile*)
  Sheet     : "sheet 1"  |  Header row: 0
  EAN       : "SKU"            (int64)
  Status    : "Status"         values: Active / Inactive
  Stock     : "Stock"
  ID        : "Product ID"
  NOTE: File contains ALL listings (active + inactive + stock-0).
        Absence = Not Listed.

Zalora Status  (SellerStatusTemplate*)
  Sheet     : "ProductStatuses"  |  Header row: 0
  EAN       : "SellerSku"
  Status    : "Status"           values: active / inactive

Zalora Stock  (SellerStockTemplate*)
  Sheet     : "Sheet"  |  Header row: 0
  EAN       : "SellerSku"
  Stock     : "Quantity"

ACTIVE RULE (ALL must be true):
  1. ZeCom LAZADA/SHOPEE/ZALORA/TIK TOK = YES
  2. Launch Dates ≤ today (or blank)
  3. STOCK Per EAN > 0  (after PH Lazada buffer)

INACTIVE RULE (ANY of these):
  1. ZeCom = NO or OFF
  2. Launch Dates > today
  3. STOCK Per EAN = 0

SPECIAL OVERRIDE (6_5_Tracker file):
  EAN present in override file → ACTIVE on that marketplace IF stock > 0
  EAN NOT in override file     → follow ZeCom rules
  Stock still gated: if stock = 0 → INACTIVE even with override
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
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class='main-header'>
  <h1>📊 PUMA Listing Audit Analyzer</h1>
  <p>ZeCom-Driven · Stock-Gated · MY / SG / PH · Lazada / Shopee / Zalora / TikTok</p>
</div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════
REGIONS      = ["MY", "SG", "PH"]
MARKETPLACES = ["Lazada", "Shopee", "Zalora", "TikTok"]

# PH Lazada: effective stock = STOCK Per EAN - 1 (min 0). All others: direct.
CHANNEL_BUFFER = {"PH": {"Lazada":1,"Shopee":0,"Zalora":0,"TikTok":0},
                  "MY": {"Lazada":0,"Shopee":0,"Zalora":0,"TikTok":0},
                  "SG": {"Lazada":0,"Shopee":0,"Zalora":0,"TikTok":0}}

# Confirmed inactive status values (case-insensitive)
INACTIVE_SET = {
    "inactive","inactivated","deactivated",
    "deleted","seller_deleted","seller deleted","banned","banned by admin",
    "unlisted","unlist","suspended","blocked",
    "violation","delisted","rejected","failed",
    "prohibited","taken down","not listed","not_listed",
    "no","off","0","false",
}

CAT_ACTIVE   = "Active"
CAT_INACTIVE = "Inactive"
CAT_NL_ART   = "Not Listed - Article Level"
CAT_NL_EAN   = "Not Listed - EAN Level"
ALL_CATS     = [CAT_ACTIVE, CAT_INACTIVE, CAT_NL_ART, CAT_NL_EAN]
CAT_STYLE    = {
    CAT_ACTIVE  :{"hdr":"#1e6f39","bg":"#c6efce","fc":"#276221","icon":"🟢"},
    CAT_INACTIVE:{"hdr":"#9c0006","bg":"#ffc7ce","fc":"#9c0006","icon":"🔴"},
    CAT_NL_ART  :{"hdr":"#7b5800","bg":"#ffeb9c","fc":"#7b5800","icon":"🟡"},
    CAT_NL_EAN  :{"hdr":"#4a235a","bg":"#e8d5f5","fc":"#4a235a","icon":"🟣"},
}

# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def clean_ean(v) -> str:
    """Normalize EAN to plain string of digits, strip .0 from int-read floats."""
    s = str(v).strip().split(".")[0].strip()
    return s if s not in ("nan","None","") else ""

def clean_str(v) -> str:
    s = str(v).strip()
    return s if s not in ("nan","None","") else ""

def is_mp_active(val) -> bool:
    if pd.isna(val): return False
    return str(val).strip().lower() not in INACTIVE_SET

def resolve_tracker(val) -> str:
    if pd.isna(val): return "UNKNOWN"
    v = str(val).strip().upper()
    if v == "YES":        return "ACTIVE"
    if v in ("NO","OFF"): return "INACTIVE"
    return "UNKNOWN"

def to_int(v) -> int:
    try: return max(0, int(float(v)))
    except: return 0

def read_file(file, sheet_name=0, header=0) -> pd.DataFrame:
    """
    Universal file reader — handles xlsx, xls, xlsm, xlsb, csv, tsv, ods.
    When sheet_name is a string, checks actual sheet names in the file first
    (case-insensitive, strip spaces) so a wrong-case sheet name never errors.
    Falls back to first sheet if the named sheet is not found.
    """
    name = getattr(file, "name", "") or ""
    ext  = name.rsplit(".", 1)[-1].lower() if "." in name else ""

    # ── CSV / TSV ──────────────────────────────────────────────────────────
    if ext in ("csv", "tsv"):
        sep = "\t" if ext == "tsv" else ","
        for enc in ("utf-8", "latin-1", "utf-8-sig"):
            try:
                file.seek(0)
                df = pd.read_csv(file, sep=sep, header=header,
                                 dtype=str, encoding=enc, low_memory=False)
                df.columns = [str(c).strip() for c in df.columns]
                return df.dropna(how="all").reset_index(drop=True)
            except UnicodeDecodeError:
                continue
            except Exception as e:
                st.error(f"Cannot read CSV '{name}': {e}")
                return pd.DataFrame()
        return pd.DataFrame()

    # ── Excel / ODS — resolve sheet name first ────────────────────────────
    engines = {
        "xlsx": ["openpyxl"], "xlsm": ["openpyxl"],
        "xls":  ["xlrd", "openpyxl"],
        "xlsb": ["pyxlsb", "openpyxl"],
        "ods":  ["odf", "openpyxl"],
    }.get(ext, ["openpyxl", "xlrd"])

    # Step 1: resolve the real sheet name using ExcelFile
    resolved_sheet = sheet_name   # default (int index or already correct string)
    if isinstance(sheet_name, str):
        for eng in engines:
            try:
                file.seek(0)
                xf = pd.ExcelFile(file, engine=eng)
                actual_sheets = xf.sheet_names   # list of real sheet names
                # Exact match first
                if sheet_name in actual_sheets:
                    resolved_sheet = sheet_name
                    break
                # Case-insensitive + strip match
                match = next(
                    (s for s in actual_sheets
                     if s.strip().lower() == sheet_name.strip().lower()),
                    None
                )
                if match:
                    resolved_sheet = match
                    break
                # No match — fall back to first sheet
                resolved_sheet = 0
                break
            except Exception:
                continue

    # Step 2: read with resolved sheet name
    for eng in engines:
        try:
            file.seek(0)
            df = pd.read_excel(file, sheet_name=resolved_sheet,
                               header=header, engine=eng)
            df.columns = [str(c).strip() for c in df.columns]
            return df.dropna(how="all").reset_index(drop=True)
        except Exception:
            continue

    # Last resort
    try:
        file.seek(0)
        df = pd.read_excel(file, sheet_name=resolved_sheet, header=header)
        df.columns = [str(c).strip() for c in df.columns]
        return df.dropna(how="all").reset_index(drop=True)
    except Exception as e:
        st.error(f"Cannot read file '{name}': {e}")
        return pd.DataFrame()

# ══════════════════════════════════════════════════════════════════════════════
# FILE LOADERS  — exact column names from real files
# ══════════════════════════════════════════════════════════════════════════════

def load_zecom(file) -> pd.DataFrame:
    """
    Sheet: "PH" | Header row: 2
    Keeps all rows where PIM Article# looks like xxx_xx
    Returns: Article_No, Tracker_Lazada, Tracker_Shopee, Tracker_Zalora,
             Tracker_TikTok, Launch_Date
    """
    try:
        df = read_file(file, sheet_name="PH", header=2)
    except Exception:
        file.seek(0)
        df = read_file(file, header=2)

    df.columns = [str(c).strip() for c in df.columns]

    # Rename confirmed columns
    rename = {
        "PIM Article#" : "Article_No",
        "Launch Dates"  : "Launch_Date",
        "LAZADA"        : "Tracker_Lazada",
        "SHOPEE"        : "Tracker_Shopee",
        "ZALORA"        : "Tracker_Zalora",
        "TIK TOK"       : "Tracker_TikTok",
    }
    df = df.rename(columns=rename)

    # Fallback for Article_No
    if "Article_No" not in df.columns:
        for c in ["Article#","Syle#","PIM Style"]:
            if c in df.columns:
                df = df.rename(columns={c: "Article_No"})
                break

    # Guarantee tracker cols
    for col in ["Tracker_Lazada","Tracker_Shopee","Tracker_Zalora","Tracker_TikTok"]:
        if col not in df.columns:
            df[col] = np.nan
    if "Launch_Date" not in df.columns:
        df["Launch_Date"] = pd.NaT

    df["Article_No"]  = df["Article_No"].apply(clean_str)
    df["Launch_Date"] = pd.to_datetime(df["Launch_Date"], errors="coerce")

    # Keep only valid article rows
    df = df[df["Article_No"].str.match(r'^\S+_\S+$', na=False)]
    df = df.drop_duplicates("Article_No").reset_index(drop=True)

    needed = ["Article_No","Tracker_Lazada","Tracker_Shopee",
              "Tracker_Zalora","Tracker_TikTok","Launch_Date"]
    return df[needed]


def load_content(file) -> pd.DataFrame:
    """
    Sheet: "content" | Header: 0
    EAN: "EAN" (int) | Article: "Color_No" | Size: "Size No."
    """
    try:
        df = read_file(file, sheet_name="content")
    except Exception:
        file.seek(0)
        df = read_file(file)

    df.columns = [str(c).strip() for c in df.columns]
    df["EAN"]        = df["EAN"].apply(clean_ean)
    df["Article_No"] = df["Color_No"].apply(clean_str) if "Color_No" in df.columns else ""

    # Size for variant description
    if "Size No." in df.columns:
        df["Size"] = df["Size No."].apply(clean_str)
    elif "Print Size Code (UK)" in df.columns:
        df["Size"] = df["Print Size Code (UK)"].apply(clean_str)
    else:
        df["Size"] = ""

    df = df[(df["EAN"] != "") & (df["Article_No"] != "")]
    return df[["EAN","Article_No","Size"]].drop_duplicates("EAN").reset_index(drop=True)


def load_override(file) -> pd.DataFrame:
    """
    Sheet: "Sheet1" | Header: 0
    Cols: Color No | EAN | Lazada Status (ID) | Shopee Status (ID) | Zalora Status (ID)

    These are EAN-level overrides. An EAN present here should be treated as
    ACTIVE on that marketplace (subject to stock check).
    Lazada Status / Shopee Status / Zalora Status contain the marketplace
    PRODUCT IDs (not text status words) — their presence confirms the EAN
    is linked/listed on that platform.
    """
    try:
        df = read_file(file, sheet_name="Sheet1")
    except Exception:
        file.seek(0)
        df = read_file(file)

    df.columns = [str(c).strip() for c in df.columns]

    # Normalize
    df["EAN"]        = df["EAN"].apply(clean_ean) if "EAN" in df.columns else ""
    df["Article_No"] = df["Color No"].apply(clean_str) if "Color No" in df.columns else ""

    # Which marketplaces the EAN is linked to
    df["Has_Lazada"] = df["Lazada Status"].apply(
        lambda v: clean_str(v) not in ("","nan","None") if "Lazada Status" in df.columns else False
    ) if "Lazada Status" in df.columns else False
    df["Has_Shopee"] = df["Shopee Status"].apply(
        lambda v: clean_str(v) not in ("","nan","None")
    ) if "Shopee Status" in df.columns else False
    df["Has_Zalora"] = df["Zalora Status"].apply(
        lambda v: clean_str(v) not in ("","nan","None")
    ) if "Zalora Status" in df.columns else False

    # Lazada Product ID
    df["Lazada_ID"] = df["Lazada Status"].apply(
        lambda v: clean_str(str(v).split(".")[0])
    ) if "Lazada Status" in df.columns else ""
    df["Shopee_ID"] = df["Shopee Status"].apply(
        lambda v: clean_str(str(v).split(".")[0])
    ) if "Shopee Status" in df.columns else ""
    df["Zalora_ID"] = df["Zalora Status"].apply(clean_str) \
        if "Zalora Status" in df.columns else ""

    df = df[df["EAN"] != ""].reset_index(drop=True)
    return df[["EAN","Article_No","Has_Lazada","Has_Shopee","Has_Zalora",
               "Lazada_ID","Shopee_ID","Zalora_ID"]]


def load_inventory(file, region: str) -> tuple:
    """
    Header: 0
    EAN: "EAN" (int64) | Stock: "STOCK Per EAN"
    PH Lazada buffer applied at audit time per channel.
    """
    df = read_file(file)
    df.columns = [str(c).strip() for c in df.columns]

    if "EAN" not in df.columns:
        st.warning(f"[{region}] Inventory: 'EAN' column not found. Cols: {list(df.columns)}")
        return pd.DataFrame(columns=["EAN","Inv_Stock"]), {}

    df["EAN"] = df["EAN"].apply(clean_ean)
    df = df[df["EAN"] != ""]

    # Stock column — "STOCK Per EAN" confirmed
    if "STOCK Per EAN" in df.columns:
        stock_col = "STOCK Per EAN"
    else:
        # Fallback scan
        candidates = ["avail_qty","stock per ean","available","qty","stock","soh","on hand"]
        stock_col  = next(
            (c for c in df.columns if any(k in c.lower() for k in candidates)),
            None
        )

    debug = {
        "Stock column used": stock_col or "NOT FOUND (stock=0)",
        "EAN rows"         : len(df),
        "Non-zero stock"   : int((df[stock_col] > 0).sum()) if stock_col else 0,
    }

    df["Inv_Stock"] = df[stock_col].apply(to_int) if stock_col else 0
    result = df[["EAN","Inv_Stock"]].drop_duplicates("EAN").reset_index(drop=True)
    return result, debug


def load_lazada(file) -> pd.DataFrame:
    """
    Sheet: "template" | Header: 0
    Real data starts after 3 instruction rows — filter by numeric SellerSKU.
    EAN: "SellerSKU" | Status: "status" | Stock: "Quantity" | ID: "Product ID"
    """
    df = read_file(file, sheet_name="template")
    df.columns = [str(c).strip() for c in df.columns]

    df["EAN"] = df["SellerSKU"].apply(clean_ean)
    # Filter: only rows where EAN is all-digits (skip instruction rows)
    df = df[df["EAN"].str.match(r'^\d{8,}$', na=False)].copy()

    df["MP_Status"]     = df["status"].apply(clean_str)    if "status"     in df.columns else "active"
    df["MP_Stock"]      = df["Quantity"].apply(to_int)     if "Quantity"   in df.columns else 0
    df["MP_ID"]         = df["Product ID"].apply(lambda v: clean_str(str(v).split(".")[0])) \
                          if "Product ID" in df.columns else ""
    df["Marketplace"]   = "Lazada"

    return df[["EAN","MP_Status","MP_Stock","MP_ID","Marketplace"]].drop_duplicates("EAN")


def load_shopee(file) -> pd.DataFrame:
    """
    Sheet: tries "sheet 1", "Sheet1", "Sheet 1", first sheet — whichever exists.
    EAN: "SKU" (int) | Status: "Status" (Active/Inactive) | Stock: "Stock" | ID: "Product ID"
    File contains ALL listings (active + inactive + stock-0).
    Absence = Not Listed on Shopee.
    """
    # Try known sheet name variants, then fall back to first sheet
    sheet_candidates = ["sheet 1", "Sheet1", "Sheet 1", "SHEET1", "Masterfile", 0]
    df = pd.DataFrame()
    for sh in sheet_candidates:
        try:
            file.seek(0)
            df = read_file(file, sheet_name=sh)
            if not df.empty:
                break
        except Exception:
            continue

    df.columns = [str(c).strip() for c in df.columns]

    df["EAN"]       = df["SKU"].apply(clean_ean)    if "SKU"        in df.columns else ""
    df["MP_Status"] = df["Status"].apply(clean_str) if "Status"     in df.columns else "Active"
    df["MP_Stock"]  = df["Stock"].apply(to_int)     if "Stock"      in df.columns else 0
    df["MP_ID"]     = df["Product ID"].apply(lambda v: clean_str(str(v).split(".")[0])) \
                      if "Product ID" in df.columns else ""
    df["Marketplace"] = "Shopee"

    df = df[df["EAN"].str.match(r'^\d{8,}$', na=False)]
    return df[["EAN","MP_Status","MP_Stock","MP_ID","Marketplace"]].drop_duplicates("EAN")


def load_zalora(status_file, stock_file) -> pd.DataFrame:
    """
    Status: sheet "ProductStatuses" | EAN: "SellerSku" | Status: "Status" (active/inactive)
    Stock : sheet "Sheet"           | EAN: "SellerSku" | Stock : "Quantity"
    """
    try:
        ds = read_file(status_file, sheet_name="ProductStatuses")
    except Exception:
        status_file.seek(0)
        ds = read_file(status_file)
    ds.columns = [str(c).strip() for c in ds.columns]

    try:
        dstk = read_file(stock_file, sheet_name="Sheet")
    except Exception:
        stock_file.seek(0)
        dstk = read_file(stock_file)
    dstk.columns = [str(c).strip() for c in dstk.columns]

    ds["EAN"]   = ds["SellerSku"].apply(clean_ean)  if "SellerSku" in ds.columns   else ""
    ds["MP_Status"] = ds["Status"].apply(clean_str) if "Status"    in ds.columns   else ""
    dstk["EAN"]     = dstk["SellerSku"].apply(clean_ean) if "SellerSku" in dstk.columns else ""
    dstk["MP_Stock"]= dstk["Quantity"].apply(to_int)     if "Quantity"  in dstk.columns else 0

    merged = ds.merge(dstk[["EAN","MP_Stock"]].drop_duplicates("EAN"), on="EAN", how="left")
    merged["MP_Stock"]    = merged["MP_Stock"].apply(to_int)
    merged["MP_ID"]       = ""
    merged["Marketplace"] = "Zalora"
    merged = merged[merged["EAN"].str.match(r'^\d{8,}$', na=False)]
    return merged[["EAN","MP_Status","MP_Stock","MP_ID","Marketplace"]].drop_duplicates("EAN")


def load_tiktok(file) -> pd.DataFrame:
    """Generic TikTok loader — column names vary."""
    df = read_file(file)
    df.columns = [str(c).strip() for c in df.columns]

    # EAN
    ean_col = next((c for c in df.columns
                    if c.lower() in ["seller sku","seller_sku","sellersku","sku"]), None)
    if not ean_col:
        st.warning("TikTok: Cannot find Seller SKU column.")
        return pd.DataFrame(columns=["EAN","MP_Status","MP_Stock","MP_ID","Marketplace"])

    df["EAN"] = df[ean_col].apply(clean_ean)
    df = df[df["EAN"].str.match(r'^\d{8,}$', na=False)].copy()

    status_col = next((c for c in df.columns
                       if c.lower() in ["status","product status","item status"]), None)
    stock_col  = next((c for c in df.columns
                       if c.lower() in ["stock","quantity","available","qty"]), None)
    id_col     = next((c for c in df.columns
                       if c.lower() in ["product id","productid","item id"]), None)

    df["MP_Status"]   = df[status_col].apply(clean_str) if status_col else "active"
    df["MP_Stock"]    = df[stock_col].apply(to_int)     if stock_col  else 0
    df["MP_ID"]       = df[id_col].apply(lambda v: clean_str(str(v).split(".")[0])) \
                        if id_col else ""
    df["Marketplace"] = "TikTok"
    return df[["EAN","MP_Status","MP_Stock","MP_ID","Marketplace"]].drop_duplicates("EAN")


# ══════════════════════════════════════════════════════════════════════════════
# STATUS DECISION ENGINE
# ══════════════════════════════════════════════════════════════════════════════
MP_TRACKER_COL = {
    "Lazada": "Tracker_Lazada",
    "Shopee": "Tracker_Shopee",
    "Zalora": "Tracker_Zalora",
    "TikTok": "Tracker_TikTok",
}

def decide_status(
    zecom_row   : pd.Series,
    mp          : str,
    eff_stock   : int,
    ean         : str,
    override_idx: dict,   # {EAN → override_row}
) -> tuple:
    """
    Returns (expected: 'ACTIVE'|'INACTIVE', reason: str)

    PRIORITY:
      1. Special Override (EAN in override file for this MP) + stock gate
      2. ZeCom NO/OFF → INACTIVE
      3. Future Launch Date → INACTIVE
      4. ZeCom YES + stock > 0 → ACTIVE
      5. ZeCom YES + stock = 0 → INACTIVE
      6. Blank tracker → INACTIVE
    """
    # PRIORITY 1: Special Override
    if ean in override_idx:
        ov = override_idx[ean]
        mp_col = f"Has_{mp}"
        if ov.get(mp_col, False):
            if eff_stock > 0:
                return "ACTIVE", f"Special Override: EAN linked on {mp} + stock={eff_stock} ✓"
            else:
                return "INACTIVE", f"Special Override: EAN linked on {mp} but stock=0"

    # PRIORITY 2: ZeCom tracker
    t_col    = MP_TRACKER_COL.get(mp, "")
    t_raw    = zecom_row.get(t_col, np.nan)
    t_status = resolve_tracker(t_raw)
    t_str    = str(t_raw).strip() if pd.notna(t_raw) else "blank"

    if t_status == "INACTIVE":
        return "INACTIVE", f"ZeCom {mp} = {t_str}"

    # PRIORITY 3: Launch Date
    launch = zecom_row.get("Launch_Date", pd.NaT)
    today  = pd.Timestamp.today().normalize()
    if pd.notna(launch):
        lts = pd.Timestamp(launch).normalize()
        if lts > today:
            return "INACTIVE", f"Future Launch Date: {lts.date()}"

    # PRIORITY 4 & 5: Tracker YES → gate on stock
    if t_status == "ACTIVE":
        if eff_stock > 0:
            return "ACTIVE", f"ZeCom={t_str} + stock={eff_stock} ✓"
        else:
            return "INACTIVE", f"ZeCom={t_str} but stock=0"

    # PRIORITY 6: Blank / unknown
    return "INACTIVE", f"ZeCom {mp} = {t_str} (blank → INACTIVE)"


# ══════════════════════════════════════════════════════════════════════════════
# AUDIT ENGINE
# ══════════════════════════════════════════════════════════════════════════════
def run_audit(mp_dfs, inv_df, zecom_df, content_df, override_df, region) -> dict:

    # ── Build lookups ─────────────────────────────────────────────────────────
    art_to_eans: dict[str,list] = {}
    ean_to_size: dict[str,str]  = {}
    ean_to_art : dict[str,str]  = {}
    for _, row in content_df.iterrows():
        art = row["Article_No"]; ean = row["EAN"]; sz = row.get("Size","")
        if art and ean:
            art_to_eans.setdefault(art,[]).append(ean)
            ean_to_size[ean] = clean_str(sz)
            ean_to_art[ean]  = art

    # Inventory index: EAN → effective stock per channel
    inv_idx: dict[str,int] = {}
    if not inv_df.empty:
        buf = CHANNEL_BUFFER.get(region, {})
        for _, row in inv_df.iterrows():
            inv_idx[row["EAN"]] = to_int(row["Inv_Stock"])

    def eff_stock(ean: str, mp: str) -> int:
        raw = inv_idx.get(ean, 0)
        return max(0, raw - CHANNEL_BUFFER.get(region,{}).get(mp,0))

    # Override index: EAN → row dict
    override_idx: dict[str,dict] = {}
    if not override_df.empty:
        for _, row in override_df.iterrows():
            override_idx[row["EAN"]] = row.to_dict()

    results = {}

    for mp in MARKETPLACES:
        mp_df = mp_dfs.get(mp, pd.DataFrame())

        # MP index: EAN → mp row
        mp_idx: dict[str,pd.Series] = {}
        if not mp_df.empty:
            for _, row in mp_df.iterrows():
                if row["EAN"]:
                    mp_idx[row["EAN"]] = row

        active_rows   = []
        inactive_rows = []
        nl_art_rows   = []
        nl_ean_rows   = []

        for _, z_row in zecom_df.iterrows():
            art = z_row["Article_No"]
            if not art: continue

            variants    = art_to_eans.get(art, [])
            n_total     = len(variants)
            n_listed    = sum(1 for e in variants if e in mp_idx)

            # No variants in content
            if not variants:
                stock_0  = eff_stock("", mp)
                exp, reason = decide_status(z_row, mp, 0, "", override_idx)
                rec = {
                    "Region":region,"Marketplace":mp,"Article No":art,
                    "EAN":"—","Size":"—","MP ID":"—",
                    "Expected Status":exp,"MP Listed":"NO","MP Status":"—",
                    "Inventory Stock":0,"MP Stock":0,"Stock Discrepancy":"—",
                    "Reason":reason,"Action Required":"—",
                    "Note":"No EAN variants in Content File",
                }
                if exp == "INACTIVE": inactive_rows.append(rec)
                else: nl_art_rows.append(rec)
                continue

            for ean in variants:
                raw_inv   = inv_idx.get(ean, 0)
                eff       = eff_stock(ean, mp)
                sz        = ean_to_size.get(ean,"")
                mp_row    = mp_idx.get(ean)
                mp_stk    = to_int(mp_row["MP_Stock"]) if mp_row is not None else 0
                mp_id     = clean_str(mp_row["MP_ID"]) if mp_row is not None else "—"
                mp_st_raw = clean_str(mp_row["MP_Status"]) if mp_row is not None else "Not Listed"

                exp, reason = decide_status(z_row, mp, eff, ean, override_idx)

                # Stock discrepancy
                if eff > 0 and mp_stk == 0 and mp_row is not None:
                    disc = f"⚠ Inv={eff} but MP shows 0"
                elif eff == 0 and mp_stk > 0 and mp_row is not None:
                    disc = f"⚠ Inv=0 but MP shows {mp_stk}"
                else:
                    disc = "—"

                rec = {
                    "Region"          : region,
                    "Marketplace"     : mp,
                    "Article No"      : art,
                    "EAN"             : ean,
                    "Size"            : sz,
                    "MP ID"           : mp_id,
                    "Expected Status" : exp,
                    "MP Listed"       : "YES" if mp_row is not None else "NO",
                    "MP Status"       : mp_st_raw,
                    "Inventory Stock" : eff,
                    "MP Stock"        : mp_stk,
                    "Stock Discrepancy": disc,
                    "Reason"          : reason,
                    "Action Required" : "—",
                }

                if mp_row is None:
                    # EAN not on MP
                    if exp == "INACTIVE":
                        rec["Action Required"] = "—"
                        inactive_rows.append(rec)
                    else:
                        if n_listed == 0:
                            rec["Note"] = "Entire article missing from MP"
                            rec["Action Required"] = "List article on MP"
                            nl_art_rows.append(rec)
                        else:
                            rec["Reason"] += f" | {n_listed}/{n_total} variants listed — this size missing"
                            rec["Action Required"] = "Add missing variant/size"
                            nl_ean_rows.append(rec)
                else:
                    # EAN IS on MP
                    mp_active = is_mp_active(mp_row["MP_Status"])

                    if exp == "ACTIVE" and mp_active:
                        if disc != "—":
                            rec["Action Required"] = "Investigate stock discrepancy"
                        active_rows.append(rec)
                    elif exp == "ACTIVE" and not mp_active:
                        rec["Reason"] += " | MP shows INACTIVE despite expected ACTIVE"
                        rec["Action Required"] = "Activate listing on MP"
                        inactive_rows.append(rec)
                    else:  # exp == INACTIVE
                        rec["Action Required"] = "Delist / deactivate on MP" if mp_active else "—"
                        inactive_rows.append(rec)

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
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        wb = writer.book

        def f(**kw):
            d = {"font_name":"Arial","font_size":9,"border":1,"valign":"vcenter","align":"left"}
            d.update(kw)
            return wb.add_format(d)

        ttl  = f(bold=True, font_size=13, font_color="#0f3460", border=0)
        sub  = f(italic=True, font_size=8, font_color="#718096", border=0)
        norm = f()

        def chdr(bg):
            return f(bold=True, bg_color=bg, font_color="#ffffff", align="center", text_wrap=True)
        def cdata(bg, fc):
            return f(bg_color=bg, font_color=fc)

        # Summary
        ws = wb.add_worksheet("Summary")
        writer.sheets["Summary"] = ws
        ws.set_zoom(90)
        ws.set_column("A:A", 8); ws.set_column("B:B", 12); ws.set_column("C:F", 32)
        ws.write("A1", "PUMA Listing Audit — Summary", ttl)
        ws.write("A2", f"Generated: {datetime.now().strftime('%d %b %Y  %H:%M')}", sub)
        ws.write("A3",
            "Active = ZeCom YES + past launch + STOCK Per EAN > 0 + listed active on MP", sub)

        r = 4
        ws.write(r,0,"Region",     chdr("#0f3460"))
        ws.write(r,1,"Marketplace",chdr("#0f3460"))
        for ci,cat in enumerate(ALL_CATS):
            ws.write(r, ci+2, f"{CAT_STYLE[cat]['icon']} {cat}",
                     chdr(CAT_STYLE[cat]["hdr"]))
        ws.set_row(r, 35); r += 1

        for region, mp_res in all_results.items():
            for mp in MARKETPLACES:
                if mp not in mp_res: continue
                ws.write(r,0,region,norm); ws.write(r,1,mp,norm)
                for ci,cat in enumerate(ALL_CATS):
                    cnt = len(mp_res[mp].get(cat, pd.DataFrame()))
                    ws.write(r, ci+2, cnt,
                             cdata(CAT_STYLE[cat]["bg"], CAT_STYLE[cat]["fc"]))
                r += 1

        # Detail sheets
        for region, mp_res in all_results.items():
            for mp in MARKETPLACES:
                if mp not in mp_res: continue
                for cat in ALL_CATS:
                    df  = mp_res[mp].get(cat, pd.DataFrame())
                    sn  = f"{mp[:5]} {region} {cat[:17]}"[:31]
                    ws2 = wb.add_worksheet(sn)
                    writer.sheets[sn] = ws2
                    ws2.set_zoom(85)
                    ws2.write(0,0, f"{CAT_STYLE[cat]['icon']} {mp} [{region}] — {cat}", ttl)
                    ws2.write(1,0, f"Records: {len(df)}", sub)

                    if df.empty:
                        ws2.write(2,0,"No records in this category.",sub); continue

                    hf  = chdr(CAT_STYLE[cat]["hdr"])
                    rfd = cdata(CAT_STYLE[cat]["bg"], CAT_STYLE[cat]["fc"])
                    hi  = {"Expected Status","Reason","Action Required","Stock Discrepancy"}

                    df = df.reset_index(drop=True)
                    for ci,col in enumerate(df.columns):
                        ws2.write(2,ci,col,hf)
                        try:
                            w = max(len(str(col)),
                                    int(df[col].astype(str).str.len().max()))
                        except: w = len(str(col))
                        ws2.set_column(ci,ci,min(w+3,50))

                    ws2.freeze_panes(3,0)
                    for ri,(_,rec) in enumerate(df.iterrows()):
                        for ci,col in enumerate(df.columns):
                            v   = rec[col]
                            sv  = "" if (isinstance(v,float) and np.isnan(v)) else str(v)
                            ws2.write(ri+3, ci, sv, rfd if col in hi else norm)

    return out.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════════════════════════
for k in ("audit_results","inv_debug","load_summary"):
    if k not in st.session_state: st.session_state[k] = {}

# ══════════════════════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════════════════════
tab_upload, tab_results, tab_debug, tab_help = st.tabs([
    "📁 Upload & Run","📊 Results & Download","🔧 Debug","❓ Help",
])

# ─── HELP ────────────────────────────────────────────────────────────────────
with tab_help:
    st.markdown("""
## ✅ Active / Inactive Logic

| Condition | Result |
|---|---|
| ZeCom = YES **+** past launch **+** stock > 0 | 🟢 **ACTIVE** |
| ZeCom = NO or OFF | 🔴 **INACTIVE** |
| Launch Date is future | 🔴 **INACTIVE** |
| ZeCom = YES but stock = 0 | 🔴 **INACTIVE** |
| Tracker blank/unknown | 🔴 **INACTIVE** |
| ZeCom ACTIVE, zero variants on MP | 🟡 **Not Listed – Article** |
| ZeCom ACTIVE, article partly listed, EAN missing | 🟣 **Not Listed – EAN** |

## ⚡ Special Override (6_5_Tracker file)
- EAN present in file → treated as **ACTIVE** on that marketplace **if stock > 0**
- Stock = 0 → **INACTIVE** even with override
- EAN not in file → follows normal ZeCom rules

## 📁 Expected File Names
| File | Sheet | Key Columns |
|---|---|---|
| ZeCom (PH_MP_eCOM_Tracking_File) | PH | PIM Article#, LAZADA, SHOPEE, ZALORA, TIK TOK, Launch Dates |
| Content (Content_file) | content | Color_No, EAN, Size No. |
| Override (6_5_Tracker) | Sheet1 | Color No, EAN, Lazada Status, Shopee Status, Zalora Status |
| Inventory (Inventory_YYYYMMDD) | first sheet | EAN, STOCK Per EAN |
| Lazada (pricestock*) | template | SellerSKU, status, Quantity, Product ID |
| Shopee (Shopee*Masterfile*) | sheet 1 | SKU, Status, Stock, Product ID |
| Zalora Status (SellerStatusTemplate*) | ProductStatuses | SellerSku, Status |
| Zalora Stock (SellerStockTemplate*) | Sheet | SellerSku, Quantity |
| TikTok (TikTokSellerCenter*) | first sheet | Seller sku, Status, Stock, Product ID |

## 📦 Inventory Stock
- Column: **STOCK Per EAN**
- PH Lazada: effective = STOCK Per EAN − 1 (min 0)
- All other channels/regions: STOCK Per EAN directly
    """)

# ─── UPLOAD ──────────────────────────────────────────────────────────────────
with tab_upload:
    st.markdown("### Step 1 — Select Regions")
    selected_regions = st.multiselect("Regions:", REGIONS, default=["PH"])

    st.markdown("### Step 2 — Master Files *(all regions)*")
    c1, c2, c3 = st.columns(3)
    with c1:
        zecom_file = st.file_uploader(
            "📋 ZeCom Tracker **[required]**",
            type=["xlsx","xls","xlsm","xlsb","csv","tsv","ods"], key="zecom",
            help="PH_MP_eCOM_Tracking_File | Sheet: PH | Header row 2")
    with c2:
        content_file = st.file_uploader(
            "📦 Content Master **[required]**",
            type=["xlsx","xls","xlsm","xlsb","csv","tsv","ods"], key="content",
            help="Content_file | Sheet: content | Cols: Color_No, EAN, Size No.")
    with c3:
        override_file = st.file_uploader(
            "⚡ Special Override *(optional)*",
            type=["xlsx","xls","xlsm","xlsb","csv","tsv","ods"], key="override",
            help="6_5_Tracker | Cols: Color No, EAN, Lazada Status, Shopee Status, Zalora Status")

    st.markdown("### Step 3 — Region Files")
    region_files: dict = {}
    for region in selected_regions:
        with st.expander(f"📂 {region}", expanded=True):
            region_files[region] = {}
            c1, c2 = st.columns(2)
            with c1:
                region_files[region]["lazada"] = st.file_uploader(
                    f"Lazada ({region}) — pricestock*.xlsx",
                    type=["xlsx","xls","xlsm","xlsb","csv","tsv","ods"], key=f"laz_{region}",
                    help="Sheet: template | EAN: SellerSKU | Status: status | ID: Product ID")
                region_files[region]["shopee"] = st.file_uploader(
                    f"Shopee ({region}) — Shopee*Masterfile*.xlsx",
                    type=["xlsx","xls","xlsm","xlsb","csv","tsv","ods"], key=f"sho_{region}",
                    help="Sheet: sheet 1 | EAN: SKU | Status: Status | ID: Product ID")
                region_files[region]["tiktok"] = st.file_uploader(
                    f"TikTok ({region}) — TikTokSellerCenter*.xlsx",
                    type=["xlsx","xls","xlsm","xlsb","csv","tsv","ods"], key=f"ttk_{region}")
            with c2:
                region_files[region]["zalora_status"] = st.file_uploader(
                    f"Zalora Status ({region}) — SellerStatusTemplate*.xlsx",
                    type=["xlsx","xls","xlsm","xlsb","csv","tsv","ods"], key=f"zst_{region}",
                    help="Sheet: ProductStatuses | EAN: SellerSku | Status: Status")
                region_files[region]["zalora_stock"] = st.file_uploader(
                    f"Zalora Stock ({region}) — SellerStockTemplate*.xlsx",
                    type=["xlsx","xls","xlsm","xlsb","csv","tsv","ods"], key=f"zsk_{region}",
                    help="Sheet: Sheet | EAN: SellerSku | Stock: Quantity")
                region_files[region]["inventory"] = st.file_uploader(
                    f"Inventory ({region}) — Inventory_*.xlsx",
                    type=["xlsx","xls","xlsm","xlsb","csv","tsv","ods"], key=f"inv_{region}",
                    help="EAN: EAN | Stock: STOCK Per EAN")

    st.markdown("---")
    run_btn = st.button("🚀 Run Listing Audit", type="primary", use_container_width=True)

    if run_btn:
        errs = []
        if not zecom_file:   errs.append("ZeCom Tracker is required.")
        if not content_file: errs.append("Content Master is required.")
        if not selected_regions: errs.append("Select at least one region.")
        for e in errs: st.error(e)

        if not errs:
            prog = st.progress(0, text="Starting…")
            load_summary = {}

            with st.spinner("Loading ZeCom Tracker…"):
                zecom_df = load_zecom(zecom_file)
            st.success(f"✅ ZeCom: **{len(zecom_df):,}** articles | "
                       f"Cols: {list(zecom_df.columns)}")
            prog.progress(10)

            with st.spinner("Loading Content Master…"):
                content_df = load_content(content_file)
            st.success(f"✅ Content: **{len(content_df):,}** EANs / "
                       f"**{content_df['Article_No'].nunique():,}** articles")
            prog.progress(20)

            override_df = pd.DataFrame(
                columns=["EAN","Article_No","Has_Lazada","Has_Shopee","Has_Zalora",
                         "Lazada_ID","Shopee_ID","Zalora_ID"])
            if override_file:
                with st.spinner("Loading Special Override…"):
                    override_df = load_override(override_file)
                st.info(f"⚡ Override: **{len(override_df):,}** EANs loaded "
                        f"(Lazada: {override_df['Has_Lazada'].sum():,} | "
                        f"Shopee: {override_df['Has_Shopee'].sum():,} | "
                        f"Zalora: {override_df['Has_Zalora'].sum():,})")
            prog.progress(25)

            all_results: dict  = {}
            inv_debug_all: dict = {}
            step = 25
            step_sz = max(1, int(70/len(selected_regions)))

            for region in selected_regions:
                rf = region_files.get(region, {})
                st.markdown(f"#### 📂 Region: {region}")
                mp_dfs: dict = {}

                if rf.get("lazada"):
                    with st.spinner(f"[{region}] Lazada…"):
                        tmp = load_lazada(rf["lazada"])
                    st.write(f"  Lazada: **{len(tmp):,}** EANs | "
                             f"Status: `{tmp['MP_Status'].value_counts().to_dict()}`")
                    mp_dfs["Lazada"] = tmp

                if rf.get("shopee"):
                    with st.spinner(f"[{region}] Shopee…"):
                        tmp = load_shopee(rf["shopee"])
                    st.write(f"  Shopee: **{len(tmp):,}** EANs | "
                             f"Status: `{tmp['MP_Status'].value_counts().to_dict()}`")
                    mp_dfs["Shopee"] = tmp

                if rf.get("zalora_status") and rf.get("zalora_stock"):
                    with st.spinner(f"[{region}] Zalora…"):
                        tmp = load_zalora(rf["zalora_status"], rf["zalora_stock"])
                    st.write(f"  Zalora: **{len(tmp):,}** EANs | "
                             f"Status: `{tmp['MP_Status'].value_counts().to_dict()}`")
                    mp_dfs["Zalora"] = tmp
                elif rf.get("zalora_status"):
                    st.warning(f"[{region}] Zalora Stock file missing — skipped")

                if rf.get("tiktok"):
                    with st.spinner(f"[{region}] TikTok…"):
                        tmp = load_tiktok(rf["tiktok"])
                    st.write(f"  TikTok: **{len(tmp):,}** EANs")
                    mp_dfs["TikTok"] = tmp

                if not mp_dfs:
                    st.warning(f"[{region}] No MP files — skipping.")
                    continue

                inv_df = pd.DataFrame(columns=["EAN","Inv_Stock"])
                if rf.get("inventory"):
                    with st.spinner(f"[{region}] Inventory…"):
                        inv_df, inv_dbg = load_inventory(rf["inventory"], region)
                    inv_debug_all[region] = inv_dbg
                    st.write(f"  Inventory: **{len(inv_df):,}** EANs | "
                             f"Column: `{inv_dbg.get('Stock column used','?')}` | "
                             f"Non-zero: **{inv_dbg.get('Non-zero stock',0):,}**")
                else:
                    st.warning(f"[{region}] No Inventory — all stock = 0")

                prog.progress(step, text=f"[{region}] Auditing…")
                with st.spinner(f"[{region}] Running audit…"):
                    region_result = run_audit(
                        mp_dfs, inv_df, zecom_df, content_df, override_df, region
                    )
                all_results[region] = region_result

                for mp in MARKETPLACES:
                    if mp in region_result:
                        for cat in ALL_CATS:
                            n = len(region_result[mp].get(cat, pd.DataFrame()))
                            st.write(f"  {CAT_STYLE[cat]['icon']} {mp} {cat}: **{n:,}**")

                step = min(step+step_sz, 95)

            prog.progress(100, text="Done!")
            st.session_state.audit_results = all_results
            st.session_state.inv_debug     = inv_debug_all
            if all_results:
                st.success("🎉 Audit complete! Go to 📊 Results & Download tab.")

# ─── RESULTS ─────────────────────────────────────────────────────────────────
with tab_results:
    results = st.session_state.audit_results
    if not results:
        st.info("Run the audit first.")
    else:
        grand = {c:0 for c in ALL_CATS}
        for mp_res in results.values():
            for cats in mp_res.values():
                for c in ALL_CATS:
                    grand[c] += len(cats.get(c, pd.DataFrame()))

        st.markdown("### 📊 Overall")
        css = {CAT_ACTIVE:"cg",CAT_INACTIVE:"cr",CAT_NL_ART:"co",CAT_NL_EAN:"cp"}
        kcols = st.columns(4)
        for i,cat in enumerate(ALL_CATS):
            kcols[i].markdown(
                f"<div class='metric-box'>"
                f"<div class='metric-value {css[cat]}'>{grand[cat]:,}</div>"
                f"<div class='metric-label'>{CAT_STYLE[cat]['icon']} {cat}</div>"
                f"</div>", unsafe_allow_html=True)

        st.markdown("### 📋 Breakdown")
        rows = []
        for region, mp_res in results.items():
            for mp in MARKETPLACES:
                if mp not in mp_res: continue
                row = {"Region":region,"Marketplace":mp}
                for cat in ALL_CATS:
                    row[f"{CAT_STYLE[cat]['icon']} {cat}"] = len(
                        mp_res[mp].get(cat,pd.DataFrame()))
                rows.append(row)
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

        st.markdown("### 🔍 Drilldown")
        d1,d2,d3 = st.columns(3)
        dr_region = d1.selectbox("Region",      list(results.keys()))
        dr_mp     = d2.selectbox("Marketplace", [m for m in MARKETPLACES
                                                  if m in results.get(dr_region,{})])
        dr_cat    = d3.selectbox("Category",    ALL_CATS)
        df_view   = results[dr_region].get(dr_mp,{}).get(dr_cat, pd.DataFrame())
        st.markdown(f"**{CAT_STYLE[dr_cat]['icon']} {dr_mp} [{dr_region}] "
                    f"— {dr_cat}: {len(df_view):,} records**")
        if not df_view.empty:
            search = st.text_input("🔎 Filter by Article No / EAN", "")
            if search.strip():
                mask = pd.Series(False, index=df_view.index)
                for col in ["Article No","EAN"]:
                    if col in df_view.columns:
                        mask |= df_view[col].astype(str).str.contains(
                            search.strip(), case=False, na=False)
                df_view = df_view[mask]
                st.caption(f"{len(df_view):,} filtered")
            st.dataframe(df_view, use_container_width=True, height=450)
        else:
            st.success("✅ No records.")

        st.markdown("---")
        with st.spinner("Building Excel…"):
            xlsx = build_excel(results)
        fname = f"PUMA_Audit_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        st.download_button(
            "📥 Download Audit Report (.xlsx)", data=xlsx, file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, type="primary")
        st.caption("Sheets: Summary + one tab per MP × Region × Category")

# ─── DEBUG ───────────────────────────────────────────────────────────────────
with tab_debug:
    st.markdown("### 🔧 Inventory Debug")
    dbg = st.session_state.inv_debug
    if dbg:
        for region, info in dbg.items():
            st.markdown(f"**{region}**")
            st.dataframe(pd.DataFrame([{"Field":k,"Value":str(v)}
                                        for k,v in info.items()]),
                         use_container_width=True, hide_index=True)
    else:
        st.info("Run audit first.")
