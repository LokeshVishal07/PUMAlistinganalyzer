"""
PUMA Marketplace Listing Audit Analyzer  v8.0
==============================================
Architecture: ZeCom-driven, vectorized, 100k+ EAN capable.
New in v8.0:
  - GRAAS TC Marketplace Inventory (per MP per region): EAN, TC_Status, Max_Qty
  - GRAAS TC Product Master Inventory (per region): EAN, PM_Quantity, PM_Reserved, PM_Occupied
  - 4-rule status engine using ZeCom + PM_Quantity + Inv_Stock
  - 3-stage stock reconciliation: MP↔TC → MP↔PM → MP↔(PM-Reserved-Occupied)
  - Action To Be columns in Status + Stock sheets
  - TC Status now a governing source in PASS condition

STATUS RULES (priority order):
  0. Special Request override (highest priority)
  1. Tracker NO/OFF          → INACTIVE  "Due to Ecom No"
  2. Tracker YES, PM=0, Inv=0 → INACTIVE  "Due to 0 Stock"
  3. Tracker YES, Launch>today, PM>0, Inv>0 → INACTIVE  "Inactive – Future launch" (still listable)
  4. Tracker YES, Launch≤today, PM>0, Inv>0 → ACTIVE   "ACTIVE – All Good"

PASS = ZeCom Status = MP Status = TC Status = Final Status  (all must agree)

STOCK (3-stage):
  Stage 1: MP Stock == TC Stock              → PASS "MP matches TC"
  Stage 2: MP Stock == PM Quantity           → PASS "Matches Product Master"
  Stage 3: MP Stock == PM_Qty-Reserved-Occ  → PASS "Matched after reserved/occupied deduction"
  All fail                                   → FAIL

PH × Lazada buffer: expected = Inv_Stock - 1 (min 0) — applies to Seller Inventory only.
"""

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import re
import warnings
warnings.filterwarnings("ignore")

st.set_page_config(page_title="PUMA Listing Audit Analyzer", layout="wide", page_icon="📊")

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
  <p>ZeCom · Seller Inventory · GRAAS TC · Product Master · PH / MY / SG</p>
</div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════
REGIONS      = ["PH", "MY", "SG"]
MARKETPLACES = ["Lazada", "Shopee", "Zalora", "TikTok"]

REGION_MARKETPLACES = {
    "PH": ["Lazada", "Shopee", "Zalora"],
    "MY": ["Lazada", "Shopee", "Zalora", "TikTok"],
    "SG": ["Lazada", "Shopee", "Zalora"],
}

TODAY        = pd.Timestamp.today().normalize()
FUTURE_WINDOW = TODAY + pd.Timedelta(days=30)

# PH Lazada buffer on Seller Inventory only
STOCK_BUFFER = {("PH","Lazada"): 1}

# Seller inventory stock column per region
INVENTORY_STOCK_COL = {
    "PH": "Avail_Qty",
    "MY": "QtyAvailable",
    "SG": "QTY",
}
INVENTORY_FILE_PREFIX = {
    "inventory_"                    : "PH",
    "puma_my_b2c_channel_inventory_": "MY",
    "sg_puma sg b2c inventory"      : "SG",
}

# Product Master stock columns per region (exact headers)
PM_STOCK_COLS = {
    "PH": {
        "qty"      : "MyStock-PH quantity",
        "reserved" : "MyStock-PH reservedQuantity",
        "occupied" : "MyStock-PH occupiedQuantity",
    },
    "MY": {
        "qty"      : "MyStock-YCH-MY quantity",
        "reserved" : "MyStock-YCH-MY reservedQuantity",
        "occupied" : "MyStock-YCH-MY occupiedQuantity",
    },
    "SG": {
        "qty"      : "MyStock-YCH-SG quantity",
        "reserved" : "MyStock-YCH-SG reservedQuantity",
        "occupied" : "MyStock-YCH-SG occupiedQuantity",
    },
}

# MP confirmed columns
MP_CONFIG = {
    "Lazada": {"ean":"SellerSKU",  "status":"status",  "stock":"Quantity", "id":"Product ID"},
    "Shopee": {"ean":"SKU",        "status":"Status",   "stock":"Stock",    "id":"Product ID"},
    "Zalora": {"ean":"SellerSku",  "status":"Status",   "stock":"Quantity", "id":""},
    "TikTok": {"ean":"Seller sku", "status":"Status",   "stock":"Quantity", "id":"Product ID"},
}

INACTIVE_STATUSES = {
    "inactive","inactivated","deactivated","deleted","seller_deleted",
    "seller deleted","banned","banned by admin","unlisted","unlist",
    "suspended","blocked","violation","delisted","rejected","failed",
    "prohibited","taken down","not listed","not_listed","no","off","0","false",
}

FILE_TYPES = ["xlsx","xls","xlsm","xlsb","csv","tsv","ods"]

LISTING_COLORS = {
    "Full New Listing": ("#1a5276","#d6eaf8"),
    "Add Variant"     : ("#7b5800","#ffeb9c"),
    "Already Listed"  : ("#1e6f39","#c6efce"),
    "No Content EANs" : ("#636363","#f2f2f2"),
}

# ══════════════════════════════════════════════════════════════════════════════
# UTILITIES
# ══════════════════════════════════════════════════════════════════════════════
def _ean(v) -> str:
    s = str(v).strip().split(".")[0].strip()
    return s if s not in ("nan","None","NaT","") else ""

def _s(v) -> str:
    s = str(v).strip()
    return s if s not in ("nan","None","NaT","") else ""

def _i(v) -> int:
    try:
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return 0
        return max(0, int(float(str(v).strip().replace(",",""))))
    except:
        return 0

def _is_active(v) -> bool:
    return str(v).strip().lower() not in INACTIVE_STATUSES

def detect_inventory_region(filename: str) -> str:
    fn = filename.lower().strip()
    for prefix, region in INVENTORY_FILE_PREFIX.items():
        if fn.startswith(prefix):
            return region
    return ""

def _resolve_sheet(file, wanted: str, engines: list) -> str:
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

def norm_col(df, candidates, new_name):
    cands = {c.strip().lower() for c in candidates}
    for col in df.columns:
        if col.strip().lower() in cands:
            return df.rename(columns={col: new_name}) if col != new_name else df
    return df

def _num_col(df, col) -> pd.Series:
    """Convert a column to int series safely."""
    return pd.to_numeric(df[col], errors="coerce").fillna(0).clip(lower=0).astype(int)

# ══════════════════════════════════════════════════════════════════════════════
# FILE LOADERS
# ══════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False)
def load_content(file) -> pd.DataFrame:
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


@st.cache_data(show_spinner=False)
def load_zecom(file) -> pd.DataFrame:
    df = pd.DataFrame()
    for sh in ["PH","MY","SG","Sheet1","Sheet","Data",0]:
        try:
            file.seek(0)
            tmp = read_file(file, sheet_name=sh, header=2)
            if len(tmp) > 5:
                df = tmp; break
        except Exception:
            continue
    if df.empty:
        file.seek(0); df = read_file(file, header=2)
    df = norm_col(df, ["PIM Article#","PIM Article","Article No","ArticleNo",
                        "Color_No","ColorNo","Style Number","Article Number"], "Article_No")
    for mp, keys in [("Lazada",["LAZADA","Lazada","lazada"]),
                      ("Shopee",["SHOPEE","Shopee","shopee"]),
                      ("Zalora",["ZALORA","Zalora","zalora"]),
                      ("TikTok",["TIK TOK","TIKTOK","TikTok","tiktok","Tik Tok"])]:
        df = norm_col(df, keys, f"Tracker_{mp}")
    df = norm_col(df, ["Launch Dates","Launch Date","LaunchDate","Go Live","live date"], "Launch_Date")
    for col in [f"Tracker_{m}" for m in MARKETPLACES] + ["Launch_Date","Article_No"]:
        if col not in df.columns:
            df[col] = np.nan
    df["Article_No"]  = df["Article_No"].apply(_s)
    df["Launch_Date"] = pd.to_datetime(df["Launch_Date"], errors="coerce")
    df = df[df["Article_No"].str.match(r'^\S+.*\S+$', na=False) &
            (df["Article_No"].str.len() > 2)]
    return df.drop_duplicates("Article_No").reset_index(drop=True)


def load_special_request(file) -> pd.DataFrame:
    df = read_file(file)
    df = norm_col(df, ["Article No","ArticleNo","Color_No","PIM Article#",
                        "Color No","Style Number"], "Article_No")
    df = norm_col(df, ["Active On","Active on","active_on","Active","ActiveOn"], "Active_On")
    df = norm_col(df, ["Inactive On","Inactive on","inactive_on","Inactive","InactiveOn"], "Inactive_On")
    if "Article_No" not in df.columns:
        st.warning("Special Request: Cannot find Article No column.")
        return pd.DataFrame(columns=["Article_No","Marketplace","Status"])
    for col in ["Active_On","Inactive_On"]:
        if col not in df.columns:
            df[col] = ""
    df["Article_No"] = df["Article_No"].apply(_s)
    df = df[df["Article_No"] != ""].reset_index(drop=True)

    def parse_mp(val):
        if not val or _s(str(val)) == "": return []
        return [p.strip().upper() for p in re.split(r"[/,]", str(val)) if p.strip()]

    rows = []
    for _, row in df.iterrows():
        art = row["Article_No"]
        for mp in parse_mp(row.get("Active_On","")):
            rows.append({"Article_No":art,"Marketplace":mp,"Status":"ACTIVE"})
        for mp in parse_mp(row.get("Inactive_On","")):
            rows.append({"Article_No":art,"Marketplace":mp,"Status":"INACTIVE"})
    if not rows:
        return pd.DataFrame(columns=["Article_No","Marketplace","Status"])
    result = pd.DataFrame(rows)
    result = result.sort_values("Status").drop_duplicates(["Article_No","Marketplace"], keep="last")
    return result.reset_index(drop=True)


def load_inventory(file, region: str) -> tuple:
    filename = getattr(file, "name","") or ""
    det = detect_inventory_region(filename)
    if det and det != region:
        st.warning(f"[{region}] Inventory file '{filename}' looks like a {det} file.")
    df = read_file(file)
    all_cols = list(df.columns)
    df = norm_col(df, ["EAN","ean","Barcode","barcode","SKU","Material"], "EAN")
    if "EAN" not in df.columns:
        return pd.DataFrame(columns=["EAN","Inv_Stock"]), {"error":"EAN not found"}
    df["EAN"] = df["EAN"].apply(_ean)
    df = df[df["EAN"].str.match(r'^\d{5,}$', na=False)]
    primary = INVENTORY_STOCK_COL.get(region, "Avail_Qty")
    stock_col = (primary if primary in df.columns else
                 next((c for c in df.columns if c.strip().lower() == primary.lower()), None))
    if not stock_col:
        fallbacks = ["avail_qty","qtyavailable","qty","stock per ean",
                     "available","on hand","soh","quantity","stock"]
        stock_col = next((c for c in df.columns if c.strip().lower() in fallbacks), None)
        if stock_col:
            st.warning(f"[{region}] Inventory: using fallback col '{stock_col}' (expected '{primary}')")
    df["Inv_Stock"] = _num_col(df, stock_col) if stock_col else 0
    debug = {
        "Region":region,"File":filename,"Expected col":primary,
        "Used col":stock_col or "NOT FOUND","EAN rows":len(df),
        "Non-zero":int((df["Inv_Stock"]>0).sum()),"All cols":", ".join(all_cols),
    }
    return df[["EAN","Inv_Stock"]].drop_duplicates("EAN").reset_index(drop=True), debug


def load_product_master(file, region: str) -> pd.DataFrame:
    """
    GRAAS TC Product Master Inventory.
    EAN column: sellerSKU
    Stock columns per region (exact headers):
      PH: MyStock-PH quantity / reservedQuantity / occupiedQuantity
      MY: MyStock-YCH-MY quantity / reservedQuantity / occupiedQuantity
      SG: MyStock-YCH-SG quantity / reservedQuantity / occupiedQuantity
    Returns: EAN, PM_Quantity, PM_Reserved, PM_Occupied
    """
    df = read_file(file)
    all_cols = list(df.columns)

    # EAN column is 'sellerSKU'
    df = norm_col(df, ["sellerSKU","sellersku","SellerSKU","Seller SKU","EAN","ean","SKU"], "EAN")
    if "EAN" not in df.columns:
        st.warning(f"[{region}] Product Master: 'sellerSKU' column not found. Cols: {all_cols[:15]}")
        return pd.DataFrame(columns=["EAN","PM_Quantity","PM_Reserved","PM_Occupied"])

    df["EAN"] = df["EAN"].apply(_ean)
    df = df[df["EAN"].str.match(r'^\d{5,}$', na=False)]

    cols = PM_STOCK_COLS.get(region, PM_STOCK_COLS["PH"])

    def find_col(target):
        """Find exact or case-insensitive match."""
        if target in df.columns: return target
        return next((c for c in df.columns if c.strip().lower() == target.strip().lower()), None)

    qty_c = find_col(cols["qty"])
    res_c = find_col(cols["reserved"])
    occ_c = find_col(cols["occupied"])

    missing = [cols[k] for k,c in [("qty",qty_c),("reserved",res_c),("occupied",occ_c)] if not c]
    if missing:
        st.warning(f"[{region}] Product Master: columns not found — {missing}. Available: {all_cols[:20]}")

    df["PM_Quantity"] = _num_col(df, qty_c) if qty_c else 0
    df["PM_Reserved"] = _num_col(df, res_c) if res_c else 0
    df["PM_Occupied"] = _num_col(df, occ_c) if occ_c else 0

    return df[["EAN","PM_Quantity","PM_Reserved","PM_Occupied"]].drop_duplicates("EAN").reset_index(drop=True)


def load_tc_marketplace(file, mp: str) -> pd.DataFrame:
    """
    GRAAS TC Marketplace Inventory.
    EAN: Custom SKU
    Status: Item status  (normalize to ACTIVE/INACTIVE)
    Max_Qty: Max Quantity  (cap setting — 0 = capped/inactive)
    Returns: EAN, TC_Status, TC_Max_Qty, Marketplace
    """
    df = read_file(file)

    df = norm_col(df, ["Custom SKU","custom sku","customsku","Custom_SKU",
                        "SKU","sku","EAN","ean","sellerSKU"], "EAN")
    df = norm_col(df, ["Item status","Item Status","item_status","Status","status"], "TC_Status_Raw")
    df = norm_col(df, ["Max Quantity","Max_Quantity","max_quantity","MaxQuantity",
                        "Max Qty","MaxQty"], "TC_Max_Qty")

    for col in ["EAN","TC_Status_Raw","TC_Max_Qty"]:
        if col not in df.columns: df[col] = np.nan

    df["EAN"]         = df["EAN"].apply(_ean)
    df["TC_Status_Raw"]= df["TC_Status_Raw"].apply(_s)
    df["TC_Max_Qty"]  = _num_col(df, "TC_Max_Qty")
    # Normalize TC status using same inactive blacklist
    df["TC_Status"]   = df["TC_Status_Raw"].apply(
        lambda v: "ACTIVE" if _is_active(v) and v != "" else "INACTIVE"
    )
    df["Marketplace"] = mp

    df = df[df["EAN"].str.match(r'^\d{5,}$', na=False)]
    return df[["EAN","TC_Status","TC_Status_Raw","TC_Max_Qty","Marketplace"]].drop_duplicates("EAN")


def _load_mp(file, mp: str, extra_sheet=None) -> pd.DataFrame:
    cfg = MP_CONFIG[mp]
    df  = read_file(file, sheet_name=extra_sheet) if extra_sheet else read_file(file)
    df = norm_col(df, [cfg["ean"],"SellerSKU","SKU","Seller sku","SellerSku"], "EAN")
    df = norm_col(df, [cfg["status"],"Status","status","ItemStatus","Listing Status"], "MP_Status")
    df = norm_col(df, [cfg["stock"],"Stock","Quantity","Available","Qty"], "MP_Stock")
    if cfg["id"]:
        df = norm_col(df, [cfg["id"],"ItemId","item_id","Product ID","ProductId"], "MP_ID")
    if "MP_ID" not in df.columns: df["MP_ID"] = ""
    for col in ["EAN","MP_Status","MP_Stock"]:
        if col not in df.columns: df[col] = np.nan
    df["EAN"]       = df["EAN"].apply(_ean)
    df["MP_Stock"]  = df["MP_Stock"].apply(_i)
    df["MP_Status"] = df["MP_Status"].apply(_s)
    df["MP_ID"]     = df["MP_ID"].apply(lambda v: _s(str(v).split(".")[0]))
    df["Marketplace"] = mp
    df = df[df["EAN"].str.match(r'^\d{8,}$', na=False)]
    return df[["EAN","MP_Status","MP_Stock","MP_ID","Marketplace"]].drop_duplicates("EAN")

def load_lazada(file)  -> pd.DataFrame: return _load_mp(file,"Lazada","template")
def load_shopee(file)  -> pd.DataFrame: return _load_mp(file,"Shopee")
def load_tiktok(file)  -> pd.DataFrame: return _load_mp(file,"TikTok")

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
# AUDIT ENGINE  v8.0
# ══════════════════════════════════════════════════════════════════════════════

def run_audit(
    mp_dfs     : dict,           # {mp: DataFrame[EAN,MP_Status,MP_Stock,MP_ID]}
    tc_dfs     : dict,           # {mp: DataFrame[EAN,TC_Status,TC_Status_Raw,TC_Max_Qty]}
    inv_df     : pd.DataFrame,   # EAN, Inv_Stock
    pm_df      : pd.DataFrame,   # EAN, PM_Quantity, PM_Reserved, PM_Occupied
    zecom_df   : pd.DataFrame,
    content_df : pd.DataFrame,
    special_df : pd.DataFrame,
    region     : str,
) -> dict:

    # ── Pre-build lookup dicts (O(1) access) ──────────────────────────────
    art_eans = (content_df.groupby("Article_No")["EAN"]
                .apply(set).reset_index()
                .rename(columns={"EAN":"Content_EANs"}))
    art_eans_idx = art_eans.set_index("Article_No")["Content_EANs"].to_dict()

    ean_size = content_df.set_index("EAN")["Size"].to_dict()

    inv     = inv_df.set_index("EAN")["Inv_Stock"].to_dict() if not inv_df.empty else {}
    pm_idx  = pm_df.set_index("EAN").to_dict("index")        if not pm_df.empty else {}

    sp_idx: dict = {}
    if not special_df.empty:
        for _, row in special_df.iterrows():
            key = (row["Article_No"], row["Marketplace"])
            if sp_idx.get(key) != "INACTIVE":
                sp_idx[key] = row["Status"]

    results_listing = []
    results_status  = []
    results_stock   = []
    results_missing = []

    region_mps = REGION_MARKETPLACES.get(region, MARKETPLACES)

    for mp in region_mps:
        mp_df  = mp_dfs.get(mp, pd.DataFrame())
        tc_df  = tc_dfs.get(mp, pd.DataFrame())

        mp_idx = mp_df.set_index("EAN").to_dict("index") if not mp_df.empty else {}
        tc_mp  = tc_df.set_index("EAN").to_dict("index") if not tc_df.empty else {}

        buf = STOCK_BUFFER.get((region, mp), 0)

        for _, z in zecom_df.iterrows():
            art     = z["Article_No"]
            tracker = _s(z.get(f"Tracker_{mp}",""))
            launch  = z["Launch_Date"]
            t_upper = tracker.upper()

            tracker_active   = (t_upper == "YES")
            tracker_inactive = (t_upper in ("NO","OFF"))

            has_launch   = pd.notna(launch)
            future_launch= has_launch and (launch > TODAY)
            past_launch  = not future_launch

            # Ecom Status derived from ZeCom only
            ecom_status = "ACTIVE" if tracker_active else "INACTIVE"

            sp_key = (art, mp.upper())
            sp_val = sp_idx.get(sp_key,"")   # "ACTIVE"|"INACTIVE"|""

            content_eans: set = art_eans_idx.get(art, set())
            listed_eans  = {e for e in content_eans if e in mp_idx}
            missing_eans = content_eans - listed_eans
            n_content    = len(content_eans)
            n_listed     = len(listed_eans)
            n_missing    = len(missing_eans)

            # Listing action
            if n_content == 0:   la_action = "No Content EANs"
            elif n_listed == 0:  la_action = "Full New Listing"
            elif n_missing == 0: la_action = "Already Listed"
            else:                la_action = "Add Variant"

            results_listing.append({
                "Region":region,"Marketplace":mp,"Article No":art,
                "Ecom Status":ecom_status,
                "Tracker Status":tracker if tracker else "blank",
                "Launch Date":launch.date() if has_launch else "—",
                "Total EANs":n_content,"Listed EANs":n_listed,
                "Missing EANs":n_missing,"Listing Action":la_action,
            })

            for ean in sorted(missing_eans):
                results_missing.append({
                    "Region":region,"Marketplace":mp,"Article No":art,
                    "Missing EAN":ean,"Size":ean_size.get(ean,""),
                })

            # ── EAN-level validation (listed EANs only) ────────────────────
            for ean in listed_eans:
                mp_row  = mp_idx[ean]
                tc_row  = tc_mp.get(ean, {})
                pm_row  = pm_idx.get(ean, {})

                inv_stock = _i(inv.get(ean, 0))
                eff_stock = max(0, inv_stock - buf)   # buffer on seller inv only

                pm_qty  = _i(pm_row.get("PM_Quantity", 0))
                pm_res  = _i(pm_row.get("PM_Reserved",  0))
                pm_occ  = _i(pm_row.get("PM_Occupied",  0))
                pm_adj  = max(0, pm_qty - pm_res - pm_occ)

                tc_status_raw = _s(tc_row.get("TC_Status_Raw","—"))
                tc_status     = _s(tc_row.get("TC_Status","—"))   # ACTIVE/INACTIVE
                tc_max_qty    = _i(tc_row.get("TC_Max_Qty", 0))
                mp_stock      = _i(mp_row.get("MP_Stock", 0))
                mp_st_raw     = _s(mp_row.get("MP_Status",""))
                mp_id         = _s(mp_row.get("MP_ID",""))
                sz            = ean_size.get(ean,"")

                # ── Final Status: 4-rule engine ────────────────────────────
                # Priority 0: Special Request
                if sp_val == "INACTIVE":
                    final_status = "INACTIVE"
                    comment      = "Special Request = INACTIVE"
                elif sp_val == "ACTIVE":
                    if eff_stock > 0 or pm_qty > 0:
                        final_status = "ACTIVE"
                        comment      = "Special Request = ACTIVE + stock > 0"
                    else:
                        final_status = "INACTIVE"
                        comment      = "Special Request = ACTIVE but stock = 0"

                # Priority 1: Tracker NO/OFF
                elif tracker_inactive:
                    final_status = "INACTIVE"
                    comment      = "Due to Ecom No"

                # Priority 2: Both stocks = 0
                elif tracker_active and pm_qty == 0 and inv_stock == 0:
                    final_status = "INACTIVE"
                    comment      = "Due to 0 Stock"

                # Priority 3: Future launch (listable but inactive)
                elif tracker_active and future_launch:
                    final_status = "INACTIVE"
                    comment      = f"Inactive – Future launch ({launch.date()})"

                # Priority 4: All good → ACTIVE
                elif tracker_active and past_launch and (pm_qty > 0 or inv_stock > 0):
                    final_status = "ACTIVE"
                    comment      = "ACTIVE – All Good"

                # Fallback
                else:
                    final_status = "INACTIVE"
                    comment      = f"Tracker='{tracker}' (blank/unknown)"

                # ── PASS condition: ZeCom=MP=TC=Final ─────────────────────
                mp_active = _is_active(mp_st_raw)
                act_mp    = "ACTIVE" if mp_active else "INACTIVE"
                act_tc    = tc_status if tc_status else "—"

                # All three must match final_status
                status_ok = (act_mp == final_status and
                             (act_tc == final_status or act_tc == "—"))
                result_sv = "✓ PASS" if status_ok else "✗ FAIL"

                # Action To Be (status)
                actions = []
                if final_status == "ACTIVE"   and act_mp == "INACTIVE":
                    actions.append("Change to Active in MP")
                if final_status == "ACTIVE"   and act_tc == "INACTIVE":
                    actions.append("Change to Active in TC")
                if final_status == "INACTIVE" and act_mp == "ACTIVE":
                    actions.append("Change to Inactive in MP")
                if final_status == "INACTIVE" and act_tc == "ACTIVE":
                    actions.append("Change to Inactive in TC")
                if future_launch and not actions:
                    actions.append(f"Keep Inactive until {launch.date()}")
                if tracker_inactive and not actions:
                    actions.append("Delist / Inactivate")
                action_sv = " | ".join(actions) if actions else "No Action"

                results_status.append({
                    "Region":region,"Marketplace":mp,
                    "Article No":art,"EAN":ean,"Size":sz,"MP ID":mp_id,
                    "ZeCom Status":ecom_status,
                    "Ecom Status":ecom_status,
                    "Final Status":final_status,
                    "MP Status":act_mp,"MP Status Raw":mp_st_raw,
                    "TC Status":act_tc,"TC Status Raw":tc_status_raw,
                    "Result":result_sv,
                    "Validation Comment":comment,
                    "Action To Be":action_sv,
                })

                # ── Stock Validation: 3-stage ──────────────────────────────
                # Stage 1: MP Stock vs TC Stock (from Product Master)
                if mp_stock == pm_qty:
                    stk_result  = "✓ PASS"
                    stk_comment = "MP matches TC (Product Master)"
                    stk_step    = "Stage 1"
                # Stage 2: MP Stock vs PM Quantity
                elif mp_stock == pm_qty:
                    stk_result  = "✓ PASS"
                    stk_comment = "Matches Product Master"
                    stk_step    = "Stage 2"
                # Stage 3: MP Stock vs Adjusted (PM - Reserved - Occupied)
                elif mp_stock == pm_adj:
                    stk_result  = "✓ PASS"
                    stk_comment = "Matched after reserved/occupied deduction"
                    stk_step    = "Stage 3"
                else:
                    stk_result  = "✗ FAIL"
                    stk_comment = "Mismatch across all stages"
                    stk_step    = "All Failed"

                # Action To Be (stock)
                if stk_result == "✓ PASS":
                    action_stk = "No Action"
                elif mp_stock < pm_qty and stk_step == "All Failed":
                    action_stk = "Update Stock in MP"
                elif mp_stock != pm_qty and stk_step != "Stage 3":
                    action_stk = "Sync TC Quantity"
                elif stk_step == "Stage 3":
                    action_stk = "No Action – Reserved Qty Impact"
                else:
                    action_stk = "Re-sync API Stock"

                results_stock.append({
                    "Region":region,"Marketplace":mp,
                    "Article No":art,"EAN":ean,"Size":sz,"MP ID":mp_id,
                    "Seller Inv Stock":inv_stock,
                    "Buffer Applied":buf,
                    "Eff Seller Stock":eff_stock,
                    "PM Quantity":pm_qty,
                    "PM Reserved":pm_res,
                    "PM Occupied":pm_occ,
                    "PM Adjusted":pm_adj,
                    "TC Max Qty":tc_max_qty,
                    "MP Stock":mp_stock,
                    "Result":stk_result,
                    "Stage Passed":stk_step,
                    "Validation Comment":stk_comment,
                    "Action To Be":action_stk,
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

def build_excel(all_results: dict, regions_run: list) -> bytes:
    out = BytesIO()
    la  = pd.concat([r["listing_analysis"]  for r in all_results.values()], ignore_index=True)
    sv  = pd.concat([r["status_validation"] for r in all_results.values()], ignore_index=True)
    stk = pd.concat([r["stock_validation"]  for r in all_results.values()], ignore_index=True)
    mv  = pd.concat([r["missing_variants"]  for r in all_results.values()], ignore_index=True)

    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        wb = writer.book
        def F(**kw):
            d = {"font_name":"Arial","font_size":9,"border":1,"valign":"vcenter","align":"left"}
            d.update(kw); return wb.add_format(d)
        norm    = F()
        ttl     = F(bold=True, font_size=13, font_color="#0f3460", border=0)
        sub     = F(italic=True, font_size=8, font_color="#718096", border=0)
        pass_f  = F(bg_color="#c6efce", font_color="#276221")
        fail_f  = F(bg_color="#ffc7ce", font_color="#9c0006")
        act_f   = F(bg_color="#fff2cc", font_color="#7b5800")
        hdr_f   = F(bold=True, bg_color="#0f3460", font_color="#ffffff",
                    align="center", text_wrap=True)
        num_f   = F(align="center")

        def write_df_sheet(ws_name, title, df, res_col="Result", act_col="Action To Be",
                           widths=None):
            ws = wb.add_worksheet(ws_name)
            writer.sheets[ws_name] = ws
            ws.write(0, 0, title, ttl)
            n_pass = int((df["Result"]=="✓ PASS").sum()) if "Result" in df.columns else 0
            n_fail = int((df["Result"]=="✗ FAIL").sum()) if "Result" in df.columns else 0
            ws.write(1, 0, f"Records: {len(df):,}  PASS: {n_pass:,}  FAIL: {n_fail:,}", sub)
            if df.empty:
                ws.write(2, 0, "No data.", sub); return
            cols = list(df.columns)
            for ci, col in enumerate(cols):
                ws.write(2, ci, col, hdr_f)
                w = (widths or {}).get(col, 18)
                ws.set_column(ci, ci, w)
            ws.freeze_panes(3, 0)
            rc = cols.index(res_col) if res_col in cols else -1
            ac = cols.index(act_col) if act_col in cols else -1
            for ri, (_, row) in enumerate(df.iterrows()):
                for ci, col in enumerate(cols):
                    v = row[col]
                    sv2 = "" if (isinstance(v, float) and np.isnan(v)) else str(v)
                    if   ci == rc: f2 = pass_f if sv2 == "✓ PASS" else fail_f
                    elif ci == ac and sv2 != "No Action": f2 = act_f
                    else: f2 = norm
                    ws.write(ri+3, ci, sv2, f2)

        # Sheet 1 — Listing Analysis
        ws1 = wb.add_worksheet("1. Listing Analysis")
        writer.sheets["1. Listing Analysis"] = ws1
        ws1.write(0, 0, "📋 Listing Analysis", ttl)
        ws1.write(1, 0, f"Records: {len(la):,}", sub)
        la_cols = list(la.columns)
        widths_la = {"Region":8,"Marketplace":12,"Article No":16,"Ecom Status":13,
                     "Tracker Status":13,"Launch Date":13,"Total EANs":10,
                     "Listed EANs":10,"Missing EANs":11,"Listing Action":20}
        for ci, col in enumerate(la_cols):
            ws1.write(2, ci, col, hdr_f)
            ws1.set_column(ci, ci, widths_la.get(col, 14))
        ws1.freeze_panes(3, 0)
        ac_idx = la_cols.index("Listing Action") if "Listing Action" in la_cols else -1
        for ri, (_, row) in enumerate(la.iterrows()):
            action = str(row.get("Listing Action",""))
            colors = LISTING_COLORS.get(action, ("#000000","#ffffff"))
            af = F(bg_color=colors[1], font_color=colors[0])
            for ci, col in enumerate(la_cols):
                v = row[col]; sv2 = "" if (isinstance(v,float) and np.isnan(v)) else str(v)
                ws1.write(ri+3, ci, sv2, af if ci == ac_idx else norm)

        # Sheet 2 — Status Validation
        write_df_sheet("2. Status Validation", "✅ Status Validation", sv,
            widths={"Region":8,"Marketplace":11,"Article No":16,"EAN":16,"Size":8,
                    "MP ID":13,"ZeCom Status":13,"Ecom Status":13,"Final Status":13,
                    "MP Status":13,"MP Status Raw":13,"TC Status":13,"TC Status Raw":13,
                    "Result":10,"Validation Comment":40,"Action To Be":35})

        # Sheet 3 — Stock Validation
        write_df_sheet("3. Stock Validation", "📦 Stock Validation", stk,
            widths={"Region":8,"Marketplace":11,"Article No":16,"EAN":16,"Size":8,
                    "MP ID":13,"Seller Inv Stock":15,"Buffer Applied":13,
                    "Eff Seller Stock":15,"PM Quantity":13,"PM Reserved":13,
                    "PM Occupied":13,"PM Adjusted":13,"TC Max Qty":12,
                    "MP Stock":12,"Result":10,"Stage Passed":14,
                    "Validation Comment":38,"Action To Be":32})

        # Sheet 4 — Missing Variants
        ws4 = wb.add_worksheet("4. Missing Variants")
        writer.sheets["4. Missing Variants"] = ws4
        ws4.write(0, 0, "🟣 Missing Variants", ttl)
        ws4.write(1, 0, f"Records: {len(mv):,}", sub)
        mv_f = F(bg_color="#e8d5f5", font_color="#4a235a")
        for ci, col in enumerate(mv.columns):
            ws4.write(2, ci, col, hdr_f)
            ws4.set_column(ci, ci, {"Region":8,"Marketplace":12,"Article No":16,
                                     "Missing EAN":18,"Size":8}.get(col,14))
        ws4.freeze_panes(3, 0)
        for ri, (_, row) in enumerate(mv.iterrows()):
            for ci, col in enumerate(mv.columns):
                v = row[col]; sv2 = "" if (isinstance(v,float) and np.isnan(v)) else str(v)
                ws4.write(ri+3, ci, sv2, mv_f)

        # Sheet 5 — Summary
        ws5 = wb.add_worksheet("5. Summary Dashboard")
        writer.sheets["5. Summary Dashboard"] = ws5
        ws5.write(0, 0, "📊 Summary Dashboard", ttl)
        ws5.write(1, 0, f"Generated: {datetime.now().strftime('%d %b %Y %H:%M')}", sub)
        sum_h = ["Region","Marketplace","Total Articles","Total EANs",
                 "Active EANs","Inactive EANs",
                 "Full New Listing","Add Variant","Already Listed",
                 "Status PASS","Status FAIL","Stock PASS","Stock FAIL"]
        for ci, h in enumerate(sum_h):
            ws5.write(3, ci, h, hdr_f); ws5.set_column(ci, ci, 20)
        ws5.set_row(3, 28)
        r = 4
        for region in regions_run:
            for mp in REGION_MARKETPLACES.get(region, MARKETPLACES):
                la_rm  = la[(la["Region"]==region)&(la["Marketplace"]==mp)]
                sv_rm  = sv[(sv["Region"]==region)&(sv["Marketplace"]==mp)]
                sk_rm  = stk[(stk["Region"]==region)&(stk["Marketplace"]==mp)]
                vals = [region, mp,
                        la_rm["Article No"].nunique(), int(la_rm["Total EANs"].sum()),
                        int((sv_rm["Final Status"]=="ACTIVE").sum()),
                        int((sv_rm["Final Status"]=="INACTIVE").sum()),
                        int((la_rm["Listing Action"]=="Full New Listing").sum()),
                        int((la_rm["Listing Action"]=="Add Variant").sum()),
                        int((la_rm["Listing Action"]=="Already Listed").sum()),
                        int((sv_rm["Result"]=="✓ PASS").sum()),
                        int((sv_rm["Result"]=="✗ FAIL").sum()),
                        int((sk_rm["Result"]=="✓ PASS").sum()),
                        int((sk_rm["Result"]=="✗ FAIL").sum())]
                for ci, v in enumerate(vals):
                    ws5.write(r, ci, v, num_f if isinstance(v,int) else norm)
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

# ── HELP ─────────────────────────────────────────────────────────────────────
with tab_help:
    st.markdown("""
## ✅ Status Logic (v8.0)

| Priority | Condition | Final Status | Comment |
|---|---|---|---|
| 0 | Special Request = INACTIVE | 🔴 INACTIVE | Special Request |
| 0 | Special Request = ACTIVE + stock > 0 | 🟢 ACTIVE | Special Request |
| 1 | Tracker = NO/OFF | 🔴 INACTIVE | Due to Ecom No |
| 2 | Tracker=YES + PM Qty=0 + Inv=0 | 🔴 INACTIVE | Due to 0 Stock |
| 3 | Tracker=YES + Future Launch | 🔴 INACTIVE | Inactive – Future launch |
| 4 | Tracker=YES + Past Launch + Stock>0 | 🟢 ACTIVE | ACTIVE – All Good |

**PASS = ZeCom Status = MP Status = TC Status = Final Status**

## 📦 Stock Validation (3-Stage)
| Stage | Check | Pass Comment |
|---|---|---|
| 1 | MP Stock == Product Master Qty | MP matches TC |
| 2 | MP Stock == PM Quantity | Matches Product Master |
| 3 | MP Stock == PM Qty − Reserved − Occupied | Matched after deduction |

## 📁 New Files (v8.0)
| File | EAN Col | Key Cols |
|---|---|---|
| TC Marketplace (per MP) | Custom SKU | Item status, Max Quantity |
| Product Master (per region) | sellerSKU | MyStock-{region} quantity/reservedQuantity/occupiedQuantity |

## 📦 Inventory Stock Buffer
- **PH × Lazada only**: Seller Inventory − 1 (min 0)
- Does NOT apply to TC or Product Master
    """)

# ── UPLOAD & RUN ─────────────────────────────────────────────────────────────
with tab_upload:
    st.markdown("### Step 1 — Select Regions")
    selected_regions = st.multiselect(
        "Regions:", REGIONS, default=["PH"],
        help="Each region needs its own ZeCom, Inventory, TC, Product Master, and MP files."
    )

    st.markdown("### Step 2 — Master Files *(shared across all regions)*")
    mc1, mc2 = st.columns(2)
    with mc1:
        content_file = st.file_uploader(
            "📦 Content Master File **[required]**",
            type=FILE_TYPES, key="content",
            help="Sheet: content | Cols: Color_No, EAN, Size No."
        )
    with mc2:
        product_master_file = st.file_uploader(
            "🔵 GRAAS Product Master **[required]** *(all regions)*",
            type=FILE_TYPES, key="product_master_global",
            help="Single file covering PH + MY + SG | EAN: sellerSKU | "
                 "Cols: MyStock-PH quantity, MyStock-YCH-MY quantity, MyStock-YCH-SG quantity "
                 "+ reservedQuantity + occupiedQuantity per region"
        )

    st.markdown("### Step 3 — Region Files")
    region_files: dict = {}

    for region in selected_regions:
        with st.expander(f"📂 {region} Files", expanded=True):
            region_files[region] = {}

            # ZeCom + Special Request
            st.markdown(f"**{region} — ZeCom & Special Request**")
            c1, c2 = st.columns(2)
            with c1:
                region_files[region]["zecom"] = st.file_uploader(
                    f"📋 ZeCom Tracker ({region}) **[required]**",
                    type=FILE_TYPES, key=f"zec_{region}",
                    help="PH_MP_eCOM_Tracking_File | Sheet: PH/MY/SG | Header row 2"
                )
            with c2:
                region_files[region]["special"] = st.file_uploader(
                    f"⚡ Special Request ({region})",
                    type=FILE_TYPES, key=f"sp_{region}",
                    help="Cols: Article No | Active On | Inactive On"
                )

            # Inventory
            st.markdown(f"**{region} — Seller Inventory**")
            region_files[region]["inventory"] = st.file_uploader(
                f"📦 Seller Inventory ({region})",
                type=FILE_TYPES, key=f"inv_{region}",
                help="PH: Avail_Qty | MY: QtyAvailable | SG: QTY"
            )

            # Marketplace files
            st.markdown(f"**{region} — Marketplace Files (Seller Center)**")
            mc1, mc2 = st.columns(2)
            with mc1:
                region_files[region]["lazada"] = st.file_uploader(
                    f"Lazada ({region}) — pricestock*.xlsx",
                    type=FILE_TYPES, key=f"laz_{region}",
                    help="Sheet: template | SellerSKU, status, Quantity, Product ID"
                )
                region_files[region]["shopee"] = st.file_uploader(
                    f"Shopee ({region}) — Shopee*.xlsx",
                    type=FILE_TYPES, key=f"sho_{region}",
                    help="SKU, Status, Stock, Product ID"
                )
            with mc2:
                region_files[region]["zalora_status"] = st.file_uploader(
                    f"Zalora Status ({region}) — SellerStatusTemplate*.xlsx",
                    type=FILE_TYPES, key=f"zst_{region}"
                )
                region_files[region]["zalora_stock"] = st.file_uploader(
                    f"Zalora Stock ({region}) — SellerStockTemplate*.xlsx",
                    type=FILE_TYPES, key=f"zsk_{region}"
                )
                region_files[region]["tiktok"] = st.file_uploader(
                    f"TikTok ({region}) — Tiktoksellercenter_batchedit*.xlsx",
                    type=FILE_TYPES, key=f"ttk_{region}",
                    help="MY only | Seller sku, Status, Quantity"
                )

            # TC Marketplace files
            st.markdown(f"**{region} — GRAAS TC Marketplace Files**")
            tc1, tc2 = st.columns(2)
            rmp = REGION_MARKETPLACES.get(region, MARKETPLACES)
            with tc1:
                region_files[region]["tc_lazada"] = st.file_uploader(
                    f"TC Lazada ({region})",
                    type=FILE_TYPES, key=f"tc_laz_{region}",
                    help="EAN: Custom SKU | Status: Item status | Cap: Max Quantity"
                )
                region_files[region]["tc_shopee"] = st.file_uploader(
                    f"TC Shopee ({region})",
                    type=FILE_TYPES, key=f"tc_sho_{region}"
                )
            with tc2:
                region_files[region]["tc_zalora"] = st.file_uploader(
                    f"TC Zalora ({region})",
                    type=FILE_TYPES, key=f"tc_zal_{region}"
                )
                if "TikTok" in rmp:
                    region_files[region]["tc_tiktok"] = st.file_uploader(
                        f"TC TikTok ({region})",
                        type=FILE_TYPES, key=f"tc_ttk_{region}"
                    )

    st.markdown("---")
    run_btn = st.button("🚀 Run Audit", type="primary", width='stretch')

    if run_btn:
        errors = []
        if not content_file:       errors.append("Content Master file is required.")
        if not selected_regions:   errors.append("Select at least one region.")
        if not product_master_file:
            errors.append("GRAAS Product Master file is required (Step 2).")
        for rg in selected_regions:
            if not region_files.get(rg,{}).get("zecom"):
                errors.append(f"[{rg}] ZeCom Tracker is required.")
        for e in errors: st.error(e)

        if not errors:
            prog = st.progress(0, text="Loading Content Master…")
            run_log = []

            with st.spinner("Loading Content Master…"):
                content_df = load_content(content_file)
            st.success(f"✅ Content: **{len(content_df):,}** EANs / "
                       f"**{content_df['Article_No'].nunique():,}** articles")
            prog.progress(10)

            all_results   = {}
            inv_debug_all = {}
            step    = 10
            step_sz = max(1, int(85/max(len(selected_regions),1)))

            for region in selected_regions:
                rf = region_files.get(region, {})
                st.markdown(f"#### 📂 {region}")

                # ZeCom
                with st.spinner(f"[{region}] ZeCom…"):
                    zecom_df = load_zecom(rf["zecom"])
                st.write(f"  ZeCom: **{len(zecom_df):,}** articles")

                # Special Request
                special_df = pd.DataFrame(columns=["Article_No","Marketplace","Status"])
                if rf.get("special"):
                    with st.spinner(f"[{region}] Special Request…"):
                        special_df = load_special_request(rf["special"])
                    st.write(f"  Special Request: **{len(special_df):,}** overrides")

                # Seller Inventory
                inv_df = pd.DataFrame(columns=["EAN","Inv_Stock"])
                if rf.get("inventory"):
                    with st.spinner(f"[{region}] Seller Inventory…"):
                        inv_df, inv_dbg = load_inventory(rf["inventory"], region)
                    inv_debug_all[region] = inv_dbg
                    st.write(f"  Seller Inv: **{len(inv_df):,}** EANs | "
                             f"col=`{inv_dbg.get('Used col','?')}` | "
                             f"non-zero=**{inv_dbg.get('Non-zero',0):,}**")
                else:
                    st.warning(f"[{region}] No Seller Inventory — stock = 0")

                # Product Master (global file, load per region using region-specific columns)
                with st.spinner(f"[{region}] Product Master ({region} columns)…"):
                    pm_df = load_product_master(product_master_file, region)
                st.write(f"  Product Master [{region}]: **{len(pm_df):,}** EANs | "
                         f"non-zero qty=**{int((pm_df['PM_Quantity']>0).sum()):,}**")

                # TC Marketplace files
                tc_dfs: dict = {}
                tc_map = {"tc_lazada":"Lazada","tc_shopee":"Shopee",
                          "tc_zalora":"Zalora","tc_tiktok":"TikTok"}
                for key, mp_name in tc_map.items():
                    if rf.get(key):
                        with st.spinner(f"[{region}] TC {mp_name}…"):
                            tc_dfs[mp_name] = load_tc_marketplace(rf[key], mp_name)
                        st.write(f"  TC {mp_name}: **{len(tc_dfs[mp_name]):,}** EANs | "
                                 f"status: `{dict(tc_dfs[mp_name]['TC_Status'].value_counts().head(3))}`")

                # Marketplace Seller Center files
                mp_dfs: dict = {}
                if rf.get("lazada"):
                    with st.spinner(f"[{region}] Lazada…"):
                        tmp = load_lazada(rf["lazada"])
                    st.write(f"  Lazada MP: **{len(tmp):,}** EANs | "
                             f"`{dict(tmp['MP_Status'].value_counts().head(3))}`")
                    mp_dfs["Lazada"] = tmp
                if rf.get("shopee"):
                    with st.spinner(f"[{region}] Shopee…"):
                        tmp = load_shopee(rf["shopee"])
                    st.write(f"  Shopee MP: **{len(tmp):,}** EANs | "
                             f"`{dict(tmp['MP_Status'].value_counts().head(3))}`")
                    mp_dfs["Shopee"] = tmp
                if rf.get("zalora_status") and rf.get("zalora_stock"):
                    with st.spinner(f"[{region}] Zalora…"):
                        tmp = load_zalora(rf["zalora_status"], rf["zalora_stock"])
                    st.write(f"  Zalora MP: **{len(tmp):,}** EANs")
                    mp_dfs["Zalora"] = tmp
                elif rf.get("zalora_status"):
                    st.warning(f"[{region}] Zalora Stock file missing — skipped")
                if rf.get("tiktok"):
                    with st.spinner(f"[{region}] TikTok…"):
                        tmp = load_tiktok(rf["tiktok"])
                    st.write(f"  TikTok MP: **{len(tmp):,}** EANs")
                    mp_dfs["TikTok"] = tmp

                if not mp_dfs:
                    st.warning(f"[{region}] No marketplace files — skipping.")
                    step += step_sz
                    prog.progress(min(step,95))
                    continue

                prog.progress(min(step+step_sz//2, 95), text=f"[{region}] Auditing…")

                with st.spinner(f"[{region}] Running audit engine…"):
                    region_result = run_audit(
                        mp_dfs, tc_dfs, inv_df, pm_df,
                        zecom_df, content_df, special_df, region
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
                prog.progress(min(step,95))

            prog.progress(100, text="Complete!")
            st.session_state.audit_results = all_results
            st.session_state.inv_debug     = inv_debug_all
            st.session_state.run_log       = run_log
            st.session_state.regions_run   = selected_regions
            if all_results:
                st.success("🎉 Audit complete! Go to 📊 Results & Download tab.")

# ── RESULTS ──────────────────────────────────────────────────────────────────
with tab_results:
    results = st.session_state.audit_results
    if not results:
        st.info("Run the audit first.")
    else:
        regions_run = st.session_state.get("regions_run", list(results.keys()))
        la  = pd.concat([r["listing_analysis"]  for r in results.values()], ignore_index=True)
        sv  = pd.concat([r["status_validation"] for r in results.values()], ignore_index=True)
        stk = pd.concat([r["stock_validation"]  for r in results.values()], ignore_index=True)
        mv  = pd.concat([r["missing_variants"]  for r in results.values()], ignore_index=True)

        st.markdown("### 📊 Overall Summary")
        k = st.columns(8)
        kpis = [
            ("Total Articles",   la["Article No"].nunique(),                    "cb"),
            ("Total Listed EANs",len(sv),                                       "cb"),
            ("Full New Listing", (la["Listing Action"]=="Full New Listing").sum(),"co"),
            ("Add Variant",      (la["Listing Action"]=="Add Variant").sum(),    "cp"),
            ("Already Listed",   (la["Listing Action"]=="Already Listed").sum(), "cg"),
            ("Active EANs",      (sv["Final Status"]=="ACTIVE").sum(),           "cg"),
            ("Status Failures",  (sv["Result"]=="✗ FAIL").sum(),                "cr"),
            ("Stock Failures",   (stk["Result"]=="✗ FAIL").sum(),               "cr"),
        ]
        for i,(label,val,css) in enumerate(kpis):
            k[i].markdown(
                f"<div class='metric-box'><div class='mv {css}'>{int(val):,}</div>"
                f"<div class='ml'>{label}</div></div>", unsafe_allow_html=True)

        st.markdown("### 📋 Region × Marketplace Breakdown")
        brows = []
        for region in regions_run:
            for mp in REGION_MARKETPLACES.get(region, MARKETPLACES):
                la_rm = la[(la["Region"]==region)&(la["Marketplace"]==mp)]
                sv_rm = sv[(sv["Region"]==region)&(sv["Marketplace"]==mp)]
                sk_rm = stk[(stk["Region"]==region)&(stk["Marketplace"]==mp)]
                brows.append({
                    "Region":region,"Marketplace":mp,
                    "Articles":la_rm["Article No"].nunique(),
                    "Full New Listing":(la_rm["Listing Action"]=="Full New Listing").sum(),
                    "Add Variant":(la_rm["Listing Action"]=="Add Variant").sum(),
                    "Already Listed":(la_rm["Listing Action"]=="Already Listed").sum(),
                    "Active":(sv_rm["Final Status"]=="ACTIVE").sum(),
                    "Inactive":(sv_rm["Final Status"]=="INACTIVE").sum(),
                    "Status PASS":(sv_rm["Result"]=="✓ PASS").sum(),
                    "Status FAIL":(sv_rm["Result"]=="✗ FAIL").sum(),
                    "Stock FAIL":(sk_rm["Result"]=="✗ FAIL").sum(),
                })
        st.dataframe(pd.DataFrame(brows), width='stretch', hide_index=True)

        st.markdown("### 🔍 Drilldown")
        t1, t2, t3, t4 = st.tabs([
            "📋 Listing Analysis","✅ Status Validation",
            "📦 Stock Validation","🟣 Missing Variants"
        ])

        def drilldown_filters(key_prefix, df):
            c1,c2,c3 = st.columns(3)
            ro = df["Region"].unique().tolist()
            mo = df["Marketplace"].unique().tolist()
            dr = c1.multiselect("Region",      ro, default=ro, key=f"{key_prefix}_r")
            dm = c2.multiselect("Marketplace", mo, default=mo, key=f"{key_prefix}_m")
            search = c3.text_input("🔎 Search Article/EAN","", key=f"{key_prefix}_s")
            out = df[df["Region"].isin(dr)&df["Marketplace"].isin(dm)]
            if search.strip():
                mask = pd.Series(False, index=out.index)
                for col in ["Article No","EAN","Missing EAN"]:
                    if col in out.columns:
                        mask |= out[col].astype(str).str.contains(
                            search.strip(), case=False, na=False)
                out = out[mask]
            return out

        with t1:
            fla = drilldown_filters("la", la)
            st.caption(f"{len(fla):,} records")
            st.dataframe(fla, width='stretch', height=400)
        with t2:
            fsv = drilldown_filters("sv", sv)
            rf2 = st.multiselect("Result Filter",["✓ PASS","✗ FAIL"],
                                  default=["✓ PASS","✗ FAIL"], key="sv_res")
            fsv = fsv[fsv["Result"].isin(rf2)]
            st.caption(f"{len(fsv):,} records")
            st.dataframe(fsv, width='stretch', height=400)
        with t3:
            fstk = drilldown_filters("stk", stk)
            rf3  = st.multiselect("Result Filter",["✓ PASS","✗ FAIL"],
                                   default=["✓ PASS","✗ FAIL"], key="stk_res")
            fstk = fstk[fstk["Result"].isin(rf3)]
            st.caption(f"{len(fstk):,} records")
            st.dataframe(fstk, width='stretch', height=400)
        with t4:
            fmv = drilldown_filters("mv", mv)
            st.caption(f"{len(fmv):,} missing variants")
            st.dataframe(fmv, width='stretch', height=400)

        st.markdown("---")
        st.markdown("### 💾 Download Full Report")
        with st.spinner("Building Excel…"):
            xlsx = build_excel(results, regions_run)
        fname = f"PUMA_Audit_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        st.download_button(
            "📥 Download Audit Report (.xlsx)", data=xlsx, file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width='stretch', type="primary"
        )
        st.caption("5-sheet Excel: Listing Analysis | Status Validation | "
                   "Stock Validation | Missing Variants | Summary Dashboard")

# ── DEBUG ─────────────────────────────────────────────────────────────────────
with tab_debug:
    st.markdown("### 🔧 Inventory Debug")
    dbg = st.session_state.inv_debug
    if dbg:
        for region, info in dbg.items():
            st.markdown(f"**{region}**")
            st.dataframe(pd.DataFrame([{"Field":k,"Value":str(v)}
                                        for k,v in info.items()]),
                         width='stretch', hide_index=True)
    else:
        st.info("Run audit first.")
    st.markdown("### 📋 Run Log")
    log = st.session_state.run_log
    if log:
        for line in log: st.text(line)
    else:
        st.info("Run audit first.")
