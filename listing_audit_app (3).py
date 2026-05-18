"""
PUMA Marketplace Listing Audit Analyzer  v7.0
==============================================
Architecture: ZeCom-driven, vectorized, 100k+ EAN capable.

MASTER FILE  (all regions): Content File only
PER-REGION   : ZeCom Tracker, Special Request, Inventory,
               Lazada, Shopee, Zalora (status+stock), TikTok

STATUS RULES
  ACTIVE   = Tracker=YES AND Launch<=today AND Stock>0 AND Special!=Inactive
  INACTIVE = Tracker=NO/OFF OR Stock=0 OR Launch>today OR Special=Inactive

LISTING ELIGIBILITY (can appear on MP)
  Tracker=YES AND Stock>0 AND Launch<=today+30d AND Special!=Inactive

LISTING ANALYSIS (article level, content EANs vs MP EANs)
  All content EANs on MP            -> Already Listed
  Some content EANs on MP           -> Add Variant
  Zero content EANs on MP           -> Full New Listing

STOCK BUFFER: PH x Lazada only -> expected = inventory - 1 (min 0)

OUTPUT: 5-sheet Excel
  1. Listing Analysis   2. Status Validation   3. Stock Validation
  4. Missing Variants   5. Summary Dashboard
"""

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import re
import warnings
warnings.filterwarnings("ignore")

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
  <h1>PUMA Marketplace Listing Audit Analyzer</h1>
  <p>ZeCom-Driven | Vectorized | PH / MY / SG | Lazada / Shopee / Zalora / TikTok</p>
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

TODAY         = pd.Timestamp.today().normalize()
FUTURE_WINDOW = TODAY + pd.Timedelta(days=30)

STOCK_BUFFER = {("PH", "Lazada"): 1}

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

MP_CONFIG = {
    "Lazada": {"ean": "SellerSKU",  "status": "status",  "stock": "Quantity",  "id": "Product ID"},
    "Shopee": {"ean": "SKU",        "status": "Status",   "stock": "Stock",     "id": "Product ID"},
    "Zalora": {"ean": "SellerSku",  "status": "Status",   "stock": "Quantity",  "id": ""},
    "TikTok": {"ean": "Seller sku", "status": "Status",   "stock": "Quantity",  "id": "Product ID"},
}

INACTIVE_STATUSES = {
    "inactive","inactivated","deactivated","deleted","seller_deleted",
    "seller deleted","banned","banned by admin","unlisted","unlist",
    "suspended","blocked","violation","delisted","rejected","failed",
    "prohibited","taken down","not listed","not_listed","no","off","0","false",
}

FILE_TYPES = ["xlsx","xls","xlsm","xlsb","csv","tsv","ods"]

LISTING_COLORS = {
    "Full New Listing": ("#1a5276", "#d6eaf8"),
    "Add Variant"     : ("#7b5800", "#ffeb9c"),
    "Already Listed"  : ("#1e6f39", "#c6efce"),
    "No Content EANs" : ("#636363", "#f2f2f2"),
}

# ══════════════════════════════════════════════════════════════════════════════
# UTILITIES
# ══════════════════════════════════════════════════════════════════════════════
def _ean(v) -> str:
    s = str(v).strip().split(".")[0].strip()
    return s if s not in ("nan", "None", "NaT", "") else ""

def _s(v) -> str:
    s = str(v).strip()
    return s if s not in ("nan", "None", "NaT", "") else ""

def _i(v) -> int:
    try:
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return 0
        return max(0, int(float(str(v).strip().replace(",", ""))))
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

    if ext in ("csv", "tsv"):
        sep = "\t" if ext == "tsv" else ","
        for enc in ("utf-8", "utf-8-sig", "latin-1"):
            try:
                file.seek(0)
                df = pd.read_csv(file, sep=sep, header=header,
                                 dtype=str, encoding=enc, low_memory=False)
                df.columns = [str(c).strip() for c in df.columns]
                return df.dropna(how="all").reset_index(drop=True)
            except UnicodeDecodeError:
                continue
            except Exception as e:
                st.error(f"Cannot read '{name}': {e}")
                return pd.DataFrame()
        return pd.DataFrame()

    ENGINE_MAP = {
        "xlsx": ["openpyxl", "xlrd"],
        "xlsm": ["openpyxl"],
        "xls" : ["xlrd", "openpyxl"],
        "xlsb": ["pyxlsb", "openpyxl"],
        "ods" : ["odf", "openpyxl"],
    }
    engines = ENGINE_MAP.get(ext, ["openpyxl", "xlrd"])

    if isinstance(sheet_name, str):
        sheet_name = _resolve_sheet(file, sheet_name, engines)

    for eng in engines:
        try:
            file.seek(0)
            df = pd.read_excel(file, sheet_name=sheet_name,
                               header=header, engine=eng)
            df.columns = [str(c).strip() for c in df.columns]
            return df.dropna(how="all").reset_index(drop=True)
        except ImportError:
            continue
        except Exception:
            continue

    try:
        file.seek(0)
        df = pd.read_excel(file, sheet_name=sheet_name, header=header)
        df.columns = [str(c).strip() for c in df.columns]
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

# ══════════════════════════════════════════════════════════════════════════════
# FILE LOADERS
# ══════════════════════════════════════════════════════════════════════════════
def load_content(file) -> pd.DataFrame:
    df = read_file(file, sheet_name="content")
    if df.empty:
        file.seek(0)
        df = read_file(file)
    df = norm_col(df, ["Color_No","ColorNo","Color No","Article No",
                        "ArticleNo","PIM Article#","Style Number"], "Article_No")
    df = norm_col(df, ["EAN","ean","Barcode","GTIN","UPC","Child SKU"], "EAN")
    if "Article_No" not in df.columns or "EAN" not in df.columns:
        st.error("Content: Missing Article_No or EAN column.")
        return pd.DataFrame(columns=["Article_No","EAN","Size"])
    df["Article_No"] = df["Article_No"].apply(_s).astype(str)
    df["EAN"]        = df["EAN"].apply(_ean).astype(str)
    size_col = next((c for c in df.columns
                     if "size" in c.lower() and c not in ("Article_No","EAN")), None)
    df["Size"] = df[size_col].apply(_s) if size_col else ""
    df = df[(df["Article_No"] != "") & (df["EAN"] != "")]
    return df[["Article_No","EAN","Size"]].drop_duplicates("EAN").reset_index(drop=True)


def load_zecom(file) -> pd.DataFrame:
    df = pd.DataFrame()
    for sh in ["PH","MY","SG","Sheet1","Sheet","Data",0]:
        try:
            file.seek(0)
            tmp = read_file(file, sheet_name=sh, header=2)
            if len(tmp) > 5:
                df = tmp
                break
        except Exception:
            continue
    if df.empty:
        file.seek(0)
        df = read_file(file, header=2)
    df = norm_col(df, ["PIM Article#","PIM Article","Article No","ArticleNo",
                        "Color_No","ColorNo","Style Number","Article Number"], "Article_No")
    for mp, keys in [("Lazada", ["LAZADA","Lazada","lazada"]),
                      ("Shopee", ["SHOPEE","Shopee","shopee"]),
                      ("Zalora", ["ZALORA","Zalora","zalora"]),
                      ("TikTok", ["TIK TOK","TIKTOK","TikTok","tiktok","Tik Tok"])]:
        df = norm_col(df, keys, f"Tracker_{mp}")
    df = norm_col(df, ["Launch Dates","Launch Date","LaunchDate","Go Live","live date"], "Launch_Date")
    for col in [f"Tracker_{m}" for m in MARKETPLACES] + ["Launch_Date","Article_No"]:
        if col not in df.columns:
            df[col] = np.nan
    df["Article_No"]  = df["Article_No"].apply(_s).astype(str)
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
        if not val or _s(str(val)) == "":
            return []
        return [p.strip().upper() for p in re.split(r"[/,]", str(val)) if p.strip()]

    rows = []
    for _, row in df.iterrows():
        art = row["Article_No"]
        for mp in parse_mp(row.get("Active_On", "")):
            rows.append({"Article_No": art, "Marketplace": mp, "Status": "ACTIVE"})
        for mp in parse_mp(row.get("Inactive_On", "")):
            rows.append({"Article_No": art, "Marketplace": mp, "Status": "INACTIVE"})
    if not rows:
        return pd.DataFrame(columns=["Article_No","Marketplace","Status"])
    result = pd.DataFrame(rows)
    result = result.sort_values("Status").drop_duplicates(["Article_No","Marketplace"], keep="last")
    return result.reset_index(drop=True)


def load_inventory(file, region: str) -> tuple:
    filename = getattr(file, "name", "") or ""
    det = detect_inventory_region(filename)
    if det and det != region:
        st.warning(f"[{region}] Inventory file '{filename}' looks like a {det} file.")
    df = read_file(file)
    all_cols = list(df.columns)
    df = norm_col(df, ["EAN","ean","Barcode","barcode","SKU","Material"], "EAN")
    if "EAN" not in df.columns:
        return pd.DataFrame(columns=["EAN","Inv_Stock"]), {"error": "EAN not found"}
    df["EAN"] = df["EAN"].apply(_ean).astype(str)
    df = df[df["EAN"].str.match(r'^\d{5,}$', na=False)]
    primary = INVENTORY_STOCK_COL.get(region, "Avail_Qty")
    stock_col = (primary if primary in df.columns else
                 next((c for c in df.columns if c.strip().lower() == primary.lower()), None))
    if not stock_col:
        fallbacks = ["avail_qty","stock per ean","qtyavailable","qty",
                     "available","on hand","soh","quantity","stock"]
        stock_col = next((c for c in df.columns if c.strip().lower() in fallbacks), None)
        if stock_col:
            st.warning(f"[{region}] Inventory: using fallback col '{stock_col}' (expected '{primary}')")
    if stock_col and stock_col in df.columns:
        df["Inv_Stock"] = pd.to_numeric(df[stock_col], errors="coerce").fillna(0).clip(lower=0).astype(int)
    else:
        df["Inv_Stock"] = 0
    debug = {
        "Region": region, "File": filename,
        "Expected col": primary, "Used col": stock_col or "NOT FOUND",
        "EAN rows": len(df), "Non-zero": int((df["Inv_Stock"] > 0).sum()),
        "All cols": ", ".join(all_cols),
    }
    return df[["EAN","Inv_Stock"]].drop_duplicates("EAN").reset_index(drop=True), debug


def _load_mp(file, mp: str, extra_sheet=None) -> pd.DataFrame:
    cfg = MP_CONFIG[mp]
    df  = read_file(file, sheet_name=extra_sheet) if extra_sheet else read_file(file)
    df = norm_col(df, [cfg["ean"],"SellerSKU","SKU","Seller sku","SellerSku"], "EAN")
    df = norm_col(df, [cfg["status"],"Status","status","ItemStatus","Listing Status"], "MP_Status")
    df = norm_col(df, [cfg["stock"],"Stock","Quantity","Available","Qty"], "MP_Stock")
    if cfg["id"]:
        df = norm_col(df, [cfg["id"],"ItemId","item_id","Product ID","ProductId"], "MP_ID")
    if "MP_ID" not in df.columns:
        df["MP_ID"] = ""
    for col in ["EAN","MP_Status","MP_Stock"]:
        if col not in df.columns:
            df[col] = np.nan
    df["EAN"]       = df["EAN"].apply(_ean).astype(str)
    df["MP_Stock"]  = df["MP_Stock"].apply(_i)
    df["MP_Status"] = df["MP_Status"].apply(_s)
    df["MP_ID"]     = df["MP_ID"].apply(lambda v: _s(str(v).split(".")[0]))
    df["Marketplace"] = mp
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
        if col not in ds.columns:
            ds[col] = np.nan
    if "EAN"      not in dstk.columns: dstk["EAN"]      = np.nan
    if "MP_Stock" not in dstk.columns: dstk["MP_Stock"]  = 0
    ds["EAN"]       = ds["EAN"].apply(_ean).astype(str)
    dstk["EAN"]     = dstk["EAN"].apply(_ean).astype(str)
    ds["MP_Status"] = ds["MP_Status"].apply(_s)
    merged = ds.merge(dstk[["EAN","MP_Stock"]].drop_duplicates("EAN"), on="EAN", how="left")
    merged["MP_Stock"]    = merged["MP_Stock"].apply(_i)
    merged["MP_ID"]       = ""
    merged["Marketplace"] = "Zalora"
    merged["EAN"]         = merged["EAN"].astype(str)
    merged = merged[merged["EAN"].str.match(r'^\d{8,}$', na=False)]
    return merged[["EAN","MP_Status","MP_Stock","MP_ID","Marketplace"]].drop_duplicates("EAN")


# ══════════════════════════════════════════════════════════════════════════════
# AUDIT ENGINE
# ══════════════════════════════════════════════════════════════════════════════
def run_audit(mp_dfs, inv_df, zecom_df, content_df, special_df, region) -> dict:

    art_eans = (content_df.groupby("Article_No")["EAN"]
                .apply(set).reset_index()
                .rename(columns={"EAN": "Content_EANs"}))
    art_eans_idx = art_eans.set_index("Article_No")["Content_EANs"].to_dict()
    ean_size     = content_df.set_index("EAN")["Size"].to_dict()
    inv          = inv_df.set_index("EAN")["Inv_Stock"].to_dict() if not inv_df.empty else {}

    sp_idx = {}
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
        mp_df = mp_dfs.get(mp, pd.DataFrame())
        mp_idx = mp_df.set_index("EAN").to_dict("index") if not mp_df.empty else {}
        buf = STOCK_BUFFER.get((region, mp), 0)

        for _, z in zecom_df.iterrows():
            art     = z["Article_No"]
            tracker = _s(z.get(f"Tracker_{mp}", ""))
            launch  = z["Launch_Date"]
            t_upper = tracker.upper()

            tracker_active   = (t_upper == "YES")
            tracker_inactive = (t_upper in ("NO", "OFF"))
            has_launch       = pd.notna(launch)
            future_launch    = has_launch and (launch > TODAY)
            future_30        = has_launch and (TODAY < launch <= FUTURE_WINDOW)
            far_future       = has_launch and (launch > FUTURE_WINDOW)

            sp_key = (art, mp.upper())
            sp_val = sp_idx.get(sp_key, "")

            content_eans = art_eans_idx.get(art, set())
            listed_eans  = {e for e in content_eans if e in mp_idx}
            missing_eans = content_eans - listed_eans
            n_content    = len(content_eans)
            n_listed     = len(listed_eans)
            n_missing    = len(missing_eans)

            if n_content == 0:   la_action = "No Content EANs"
            elif n_listed == 0:  la_action = "Full New Listing"
            elif n_missing == 0: la_action = "Already Listed"
            else:                la_action = "Add Variant"

            results_listing.append({
                "Region": region, "Marketplace": mp, "Article No": art,
                "Tracker Status": tracker if tracker else "blank",
                "Launch Date": launch.date() if has_launch else "-",
                "Total EANs": n_content, "Listed EANs": n_listed,
                "Missing EANs": n_missing, "Listing Action": la_action,
            })

            for ean in sorted(missing_eans):
                results_missing.append({
                    "Region": region, "Marketplace": mp, "Article No": art,
                    "Missing EAN": ean, "Size": ean_size.get(ean, ""),
                })

            for ean in listed_eans:
                mp_row   = mp_idx[ean]
                inv_stock = _i(inv.get(ean, 0))
                eff_stock = max(0, inv_stock - buf)
                mp_stock  = _i(mp_row.get("MP_Stock", 0))
                mp_st_raw = _s(mp_row.get("MP_Status", ""))
                mp_id     = _s(mp_row.get("MP_ID", ""))
                sz        = ean_size.get(ean, "")

                # Expected status
                if sp_val == "INACTIVE":
                    exp_status = "INACTIVE"
                    reason     = "Special Request = INACTIVE"
                elif sp_val == "ACTIVE":
                    if eff_stock > 0:
                        exp_status = "ACTIVE"
                        reason     = "Special Request = ACTIVE + stock > 0"
                    else:
                        exp_status = "INACTIVE"
                        reason     = "Special Request = ACTIVE but stock = 0"
                elif tracker_inactive:
                    exp_status = "INACTIVE"
                    reason     = f"Tracker = {tracker}"
                elif far_future:
                    exp_status = "INACTIVE"
                    reason     = f"Launch Date {launch.date()} > today+30d"
                elif not tracker_active:
                    exp_status = "INACTIVE"
                    reason     = f"Tracker = '{tracker}' (blank/unknown)"
                elif eff_stock <= 0:
                    exp_status = "INACTIVE"
                    reason     = "Tracker=YES but stock=0"
                elif future_30:
                    exp_status = "INACTIVE"
                    reason     = f"Future launch {launch.date()} (listable)"
                else:
                    exp_status = "ACTIVE"
                    reason     = f"Tracker=YES, stock={eff_stock}, launch OK"

                act_status = "ACTIVE" if _is_active(mp_st_raw) else "INACTIVE"
                result_sv  = "PASS" if exp_status == act_status else "FAIL"

                results_status.append({
                    "Region": region, "Marketplace": mp,
                    "Article No": art, "EAN": ean, "Size": sz, "MP ID": mp_id,
                    "Expected Status": exp_status, "Actual Status": act_status,
                    "MP Status Raw": mp_st_raw, "Result": result_sv, "Reason": reason,
                })

                # Stock validation
                exp_stock    = eff_stock
                stock_result = "PASS" if exp_stock == mp_stock else "FAIL"
                disc         = mp_stock - exp_stock

                results_stock.append({
                    "Region": region, "Marketplace": mp,
                    "Article No": art, "EAN": ean, "Size": sz, "MP ID": mp_id,
                    "Inventory Stock": inv_stock, "Buffer Applied": buf,
                    "Expected Stock": exp_stock, "Marketplace Stock": mp_stock,
                    "Result": stock_result, "Discrepancy": disc,
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
    import xlsxwriter
    out = BytesIO()
    la  = pd.concat([r["listing_analysis"]  for r in all_results.values()], ignore_index=True)
    sv  = pd.concat([r["status_validation"] for r in all_results.values()], ignore_index=True)
    stk = pd.concat([r["stock_validation"]  for r in all_results.values()], ignore_index=True)
    mv  = pd.concat([r["missing_variants"]  for r in all_results.values()], ignore_index=True)

    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        wb = writer.book

        def F(**kw):
            d = {"font_name":"Arial","font_size":9,"border":1,"valign":"vcenter","align":"left"}
            d.update(kw)
            return wb.add_format(d)

        norm   = F()
        ttl    = F(bold=True, font_size=13, font_color="#0f3460", border=0)
        sub    = F(italic=True, font_size=8, font_color="#718096", border=0)
        pass_f = F(bg_color="#c6efce", font_color="#276221")
        fail_f = F(bg_color="#ffc7ce", font_color="#9c0006")
        hdr_f  = F(bold=True, bg_color="#0f3460", font_color="#ffffff",
                   align="center", text_wrap=True)
        num_f  = F(align="center")

        def safe_val(v):
            return "" if (isinstance(v, float) and np.isnan(v)) else str(v)

        # Sheet 1 — Listing Analysis
        ws1 = wb.add_worksheet("1. Listing Analysis")
        writer.sheets["1. Listing Analysis"] = ws1
        ws1.write(0, 0, "Listing Analysis", ttl)
        ws1.write(1, 0, f"Records: {len(la):,}", sub)
        la_widths = {"Region":8,"Marketplace":12,"Article No":16,"Tracker Status":14,
                     "Launch Date":13,"Total EANs":11,"Listed EANs":11,
                     "Missing EANs":12,"Listing Action":20}
        if not la.empty:
            for ci, col in enumerate(la.columns):
                ws1.write(2, ci, col, hdr_f)
                ws1.set_column(ci, ci, la_widths.get(col, 15))
            ws1.freeze_panes(3, 0)
            ac_idx = list(la.columns).index("Listing Action") if "Listing Action" in la.columns else -1
            for ri, (_, row) in enumerate(la.iterrows()):
                action = str(row.get("Listing Action", ""))
                colors = LISTING_COLORS.get(action, ("#000000", "#ffffff"))
                af = F(bg_color=colors[1], font_color=colors[0])
                for ci, col in enumerate(la.columns):
                    ws1.write(ri+3, ci, safe_val(row[col]), af if ci == ac_idx else norm)

        # Sheet 2 — Status Validation
        ws2 = wb.add_worksheet("2. Status Validation")
        writer.sheets["2. Status Validation"] = ws2
        ws2.write(0, 0, "Status Validation", ttl)
        n_pass = int((sv["Result"]=="PASS").sum()) if "Result" in sv.columns else 0
        n_fail = int((sv["Result"]=="FAIL").sum()) if "Result" in sv.columns else 0
        ws2.write(1, 0, f"Records: {len(sv):,}  PASS: {n_pass:,}  FAIL: {n_fail:,}", sub)
        sv_widths = {"Region":8,"Marketplace":11,"Article No":16,"EAN":16,"Size":8,
                     "MP ID":13,"Expected Status":16,"Actual Status":14,
                     "MP Status Raw":14,"Result":10,"Reason":42}
        if not sv.empty:
            for ci, col in enumerate(sv.columns):
                ws2.write(2, ci, col, hdr_f)
                ws2.set_column(ci, ci, sv_widths.get(col, 15))
            ws2.freeze_panes(3, 0)
            rc = list(sv.columns).index("Result") if "Result" in sv.columns else -1
            for ri, (_, row) in enumerate(sv.iterrows()):
                res = str(row.get("Result", ""))
                rf  = pass_f if res == "PASS" else fail_f
                for ci, col in enumerate(sv.columns):
                    ws2.write(ri+3, ci, safe_val(row[col]), rf if ci == rc else norm)

        # Sheet 3 — Stock Validation
        ws3 = wb.add_worksheet("3. Stock Validation")
        writer.sheets["3. Stock Validation"] = ws3
        ws3.write(0, 0, "Stock Validation", ttl)
        n_sp = int((stk["Result"]=="PASS").sum()) if "Result" in stk.columns else 0
        n_sf = int((stk["Result"]=="FAIL").sum()) if "Result" in stk.columns else 0
        ws3.write(1, 0, f"Records: {len(stk):,}  PASS: {n_sp:,}  FAIL: {n_sf:,}", sub)
        stk_widths = {"Region":8,"Marketplace":11,"Article No":16,"EAN":16,"Size":8,
                      "MP ID":13,"Inventory Stock":16,"Buffer Applied":13,
                      "Expected Stock":14,"Marketplace Stock":16,"Result":10,"Discrepancy":12}
        if not stk.empty:
            for ci, col in enumerate(stk.columns):
                ws3.write(2, ci, col, hdr_f)
                ws3.set_column(ci, ci, stk_widths.get(col, 15))
            ws3.freeze_panes(3, 0)
            rc3 = list(stk.columns).index("Result") if "Result" in stk.columns else -1
            for ri, (_, row) in enumerate(stk.iterrows()):
                res = str(row.get("Result", ""))
                rf  = pass_f if res == "PASS" else fail_f
                for ci, col in enumerate(stk.columns):
                    ws3.write(ri+3, ci, safe_val(row[col]), rf if ci == rc3 else norm)

        # Sheet 4 — Missing Variants
        ws4 = wb.add_worksheet("4. Missing Variants")
        writer.sheets["4. Missing Variants"] = ws4
        ws4.write(0, 0, "Missing Variants", ttl)
        ws4.write(1, 0, f"Records: {len(mv):,}", sub)
        mv_f = F(bg_color="#e8d5f5", font_color="#4a235a")
        mv_widths = {"Region":8,"Marketplace":12,"Article No":16,"Missing EAN":18,"Size":8}
        if not mv.empty:
            for ci, col in enumerate(mv.columns):
                ws4.write(2, ci, col, hdr_f)
                ws4.set_column(ci, ci, mv_widths.get(col, 14))
            ws4.freeze_panes(3, 0)
            for ri, (_, row) in enumerate(mv.iterrows()):
                for ci, col in enumerate(mv.columns):
                    ws4.write(ri+3, ci, safe_val(row[col]), mv_f)

        # Sheet 5 — Summary
        ws5 = wb.add_worksheet("5. Summary Dashboard")
        writer.sheets["5. Summary Dashboard"] = ws5
        ws5.write(0, 0, "Summary Dashboard", ttl)
        ws5.write(1, 0, f"Generated: {datetime.now().strftime('%d %b %Y %H:%M')}", sub)
        sum_h = ["Region","Marketplace","Total Articles","Total EANs",
                 "Active EANs","Inactive EANs","Full New Listing",
                 "Add Variant","Already Listed","Status PASS","Status FAIL",
                 "Stock PASS","Stock FAIL"]
        for ci, h in enumerate(sum_h):
            ws5.write(3, ci, h, hdr_f)
            ws5.set_column(ci, ci, 18)
        ws5.set_row(3, 28)

        r = 4
        for region in regions_run:
            for mp in REGION_MARKETPLACES.get(region, MARKETPLACES):
                la_rm  = la[(la["Region"]==region)&(la["Marketplace"]==mp)] if not la.empty else pd.DataFrame()
                sv_rm  = sv[(sv["Region"]==region)&(sv["Marketplace"]==mp)] if not sv.empty else pd.DataFrame()
                sk_rm  = stk[(stk["Region"]==region)&(stk["Marketplace"]==mp)] if not stk.empty else pd.DataFrame()

                def sc(df, col, val=None):
                    if df.empty or col not in df.columns: return 0
                    return int(df[col].nunique()) if val is None else int((df[col]==val).sum())

                vals = [region, mp,
                        sc(la_rm,"Article No"), int(la_rm["Total EANs"].sum()) if "Total EANs" in la_rm.columns else 0,
                        sc(sv_rm,"Expected Status","ACTIVE"), sc(sv_rm,"Expected Status","INACTIVE"),
                        sc(la_rm,"Listing Action","Full New Listing"), sc(la_rm,"Listing Action","Add Variant"),
                        sc(la_rm,"Listing Action","Already Listed"),
                        sc(sv_rm,"Result","PASS"), sc(sv_rm,"Result","FAIL"),
                        sc(sk_rm,"Result","PASS"), sc(sk_rm,"Result","FAIL")]
                for ci, v in enumerate(vals):
                    ws5.write(r, ci, v, num_f if isinstance(v, int) else norm)
                r += 1

    return out.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════════════════════════
for k in ("audit_results", "inv_debug", "run_log", "regions_run"):
    if k not in st.session_state:
        st.session_state[k] = {} if k not in ("run_log", "regions_run") else []

# ══════════════════════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════════════════════
tab_upload, tab_results, tab_debug, tab_help = st.tabs([
    "Upload & Run", "Results & Download", "Debug", "Help"
])

# ── HELP ─────────────────────────────────────────────────────────────────────
with tab_help:
    st.markdown("""
## Status Logic

| Condition | Expected Status |
|---|---|
| Tracker=YES + Past Launch + Stock>0 | ACTIVE |
| Tracker=NO or OFF | INACTIVE |
| Inventory Stock = 0 | INACTIVE |
| Launch Date > today+30d | INACTIVE |
| Launch within 30 days (future) | INACTIVE (listable) |
| Special Request = INACTIVE | INACTIVE (overrides all) |
| Special Request = ACTIVE + stock>0 | ACTIVE |

## Listing Analysis

| Situation | Action |
|---|---|
| Zero content EANs on MP | Full New Listing |
| Some content EANs on MP | Add Variant |
| All content EANs on MP | Already Listed |

## File Guide

| File | Scope | Key Columns |
|---|---|---|
| Content Master | All regions | Color_No, EAN, Size No. |
| ZeCom Tracker | Per region | PIM Article#, LAZADA, SHOPEE, ZALORA, TIK TOK, Launch Dates |
| Special Request | Per region | Article No, Active On, Inactive On |
| Inventory PH (Inventory_*) | Per region | EAN, Avail_Qty |
| Inventory MY (PUMA_MY_B2C_*) | Per region | EAN, QtyAvailable |
| Inventory SG (SG_PUMA SG B2C*) | Per region | EAN, QTY |
| Lazada | Per region | Sheet=template, SellerSKU, status, Quantity |
| Shopee | Per region | SKU, Status, Stock |
| Zalora | Per region (2 files) | SellerSku, Status / Quantity |
| TikTok | MY only | Seller sku, Status, Quantity |

## Stock Buffer
- PH x Lazada only: Expected Stock = Inventory - 1 (min 0)
- All other channels: use inventory directly
    """)

# ── UPLOAD & RUN ─────────────────────────────────────────────────────────────
with tab_upload:
    st.markdown("### Step 1 — Select Regions")
    selected_regions = st.multiselect(
        "Regions:", REGIONS, default=["PH"],
        help="Each region needs its own ZeCom, Inventory, and MP files."
    )

    st.markdown("### Step 2 — Master File *(shared across all regions)*")
    mc1, mc2 = st.columns(2)
    with mc1:
        content_file = st.file_uploader(
            "Content Master File [required]",
            type=FILE_TYPES, key="content",
            help="Sheet: content | Cols: Color_No, EAN, Size No."
        )

    st.markdown("### Step 3 — Region Files")
    region_files = {}

    for region in selected_regions:
        with st.expander(f"{region} Files", expanded=True):
            region_files[region] = {}

            st.markdown(f"**{region} — ZeCom & Special Request**")
            c1, c2 = st.columns(2)
            with c1:
                region_files[region]["zecom"] = st.file_uploader(
                    f"ZeCom Tracker ({region}) [required]",
                    type=FILE_TYPES, key=f"zec_{region}",
                    help="PH_MP_eCOM_Tracking_File | Header row 2"
                )
            with c2:
                region_files[region]["special"] = st.file_uploader(
                    f"Special Request ({region})",
                    type=FILE_TYPES, key=f"sp_{region}",
                    help="Cols: Article No | Active On | Inactive On"
                )

            st.markdown(f"**{region} — Inventory**")
            region_files[region]["inventory"] = st.file_uploader(
                f"Seller Inventory ({region})",
                type=FILE_TYPES, key=f"inv_{region}",
                help="PH: Avail_Qty | MY: QtyAvailable | SG: QTY"
            )

            st.markdown(f"**{region} — Marketplace Files**")
            mc1, mc2 = st.columns(2)
            with mc1:
                region_files[region]["lazada"] = st.file_uploader(
                    f"Lazada ({region}) — pricestock*.xlsx",
                    type=FILE_TYPES, key=f"laz_{region}"
                )
                region_files[region]["shopee"] = st.file_uploader(
                    f"Shopee ({region}) — Shopee*.xlsx",
                    type=FILE_TYPES, key=f"sho_{region}"
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
                    help="MY only"
                )

    st.markdown("---")
    run_btn = st.button("Run Audit", type="primary", use_container_width=True)

    if run_btn:
        errors = []
        if not content_file:       errors.append("Content Master file is required.")
        if not selected_regions:   errors.append("Select at least one region.")
        for rg in selected_regions:
            if not region_files.get(rg, {}).get("zecom"):
                errors.append(f"[{rg}] ZeCom Tracker is required.")
        for e in errors:
            st.error(e)

        if not errors:
            prog = st.progress(0, text="Loading Content Master...")
            run_log = []

            with st.spinner("Loading Content Master..."):
                content_df = load_content(content_file)
            st.success(f"Content: {len(content_df):,} EANs / {content_df['Article_No'].nunique():,} articles")
            prog.progress(10)

            all_results  = {}
            inv_debug_all = {}
            step    = 10
            step_sz = max(1, int(85 / max(len(selected_regions), 1)))

            for region in selected_regions:
                rf = region_files.get(region, {})
                st.markdown(f"#### {region}")

                with st.spinner(f"[{region}] ZeCom..."):
                    zecom_df = load_zecom(rf["zecom"])
                st.write(f"  ZeCom: {len(zecom_df):,} articles")

                special_df = pd.DataFrame(columns=["Article_No","Marketplace","Status"])
                if rf.get("special"):
                    with st.spinner(f"[{region}] Special Request..."):
                        special_df = load_special_request(rf["special"])
                    st.write(f"  Special Request: {len(special_df):,} overrides")

                inv_df = pd.DataFrame(columns=["EAN","Inv_Stock"])
                if rf.get("inventory"):
                    with st.spinner(f"[{region}] Inventory..."):
                        inv_df, inv_dbg = load_inventory(rf["inventory"], region)
                    inv_debug_all[region] = inv_dbg
                    st.write(f"  Inventory: {len(inv_df):,} EANs | col={inv_dbg.get('Used col','?')} | non-zero={inv_dbg.get('Non-zero',0):,}")
                else:
                    st.warning(f"[{region}] No Inventory file — stock = 0")

                mp_dfs = {}
                if rf.get("lazada"):
                    with st.spinner(f"[{region}] Lazada..."):
                        tmp = load_lazada(rf["lazada"])
                    st.write(f"  Lazada: {len(tmp):,} EANs")
                    mp_dfs["Lazada"] = tmp
                if rf.get("shopee"):
                    with st.spinner(f"[{region}] Shopee..."):
                        tmp = load_shopee(rf["shopee"])
                    st.write(f"  Shopee: {len(tmp):,} EANs")
                    mp_dfs["Shopee"] = tmp
                if rf.get("zalora_status") and rf.get("zalora_stock"):
                    with st.spinner(f"[{region}] Zalora..."):
                        tmp = load_zalora(rf["zalora_status"], rf["zalora_stock"])
                    st.write(f"  Zalora: {len(tmp):,} EANs")
                    mp_dfs["Zalora"] = tmp
                elif rf.get("zalora_status"):
                    st.warning(f"[{region}] Zalora Stock file missing — skipped")
                if rf.get("tiktok"):
                    with st.spinner(f"[{region}] TikTok..."):
                        tmp = load_tiktok(rf["tiktok"])
                    st.write(f"  TikTok: {len(tmp):,} EANs")
                    mp_dfs["TikTok"] = tmp

                if not mp_dfs:
                    st.warning(f"[{region}] No marketplace files — skipping.")
                    step += step_sz
                    prog.progress(min(step, 95))
                    continue

                prog.progress(min(step + step_sz // 2, 95), text=f"[{region}] Auditing...")
                with st.spinner(f"[{region}] Running audit..."):
                    region_result = run_audit(mp_dfs, inv_df, zecom_df, content_df, special_df, region)

                all_results[region] = region_result
                la = region_result["listing_analysis"]
                sv = region_result["status_validation"]

                def _n(df, col, val):
                    return int((df[col]==val).sum()) if col in df.columns and not df.empty else 0

                st.success(
                    f"[{region}] Done — "
                    f"Full New: {_n(la,'Listing Action','Full New Listing'):,} | "
                    f"Add Variant: {_n(la,'Listing Action','Add Variant'):,} | "
                    f"Already Listed: {_n(la,'Listing Action','Already Listed'):,} | "
                    f"Status FAIL: {_n(sv,'Result','FAIL'):,}"
                )
                step += step_sz
                prog.progress(min(step, 95))

            prog.progress(100, text="Complete!")
            st.session_state["audit_results"] = all_results
            st.session_state["inv_debug"]     = inv_debug_all
            st.session_state["run_log"]       = run_log
            st.session_state["regions_run"]   = selected_regions
            if all_results:
                st.success("Audit complete! Go to Results & Download tab.")

# ── RESULTS ──────────────────────────────────────────────────────────────────
with tab_results:
    results = st.session_state["audit_results"]
    if not results:
        st.info("Run the audit first.")
    else:
        regions_run = st.session_state.get("regions_run", list(results.keys()))
        la  = pd.concat([r["listing_analysis"]  for r in results.values()], ignore_index=True)
        sv  = pd.concat([r["status_validation"] for r in results.values()], ignore_index=True)
        stk = pd.concat([r["stock_validation"]  for r in results.values()], ignore_index=True)
        mv  = pd.concat([r["missing_variants"]  for r in results.values()], ignore_index=True)

        def _sc(df, col, val=None):
            if df.empty or col not in df.columns: return 0
            return int(df[col].nunique()) if val is None else int((df[col]==val).sum())

        st.markdown("### Overall Summary")
        k = st.columns(8)
        kpis = [
            ("Total Articles",    _sc(la,"Article No"),                        "cb"),
            ("Total Listed EANs", len(sv),                                     "cb"),
            ("Full New Listing",  _sc(la,"Listing Action","Full New Listing"), "co"),
            ("Add Variant",       _sc(la,"Listing Action","Add Variant"),      "cp"),
            ("Already Listed",    _sc(la,"Listing Action","Already Listed"),   "cg"),
            ("Active EANs",       _sc(sv,"Expected Status","ACTIVE"),          "cg"),
            ("Status Failures",   _sc(sv,"Result","FAIL"),                     "cr"),
            ("Stock Failures",    _sc(stk,"Result","FAIL"),                    "cr"),
        ]
        for i, (label, val, css) in enumerate(kpis):
            k[i].markdown(
                f"<div class='metric-box'><div class='mv {css}'>{int(val):,}</div>"
                f"<div class='ml'>{label}</div></div>", unsafe_allow_html=True)

        st.markdown("### Region x Marketplace Breakdown")
        brows = []
        for region in regions_run:
            for mp in REGION_MARKETPLACES.get(region, MARKETPLACES):
                la_rm  = la[(la["Region"]==region)&(la["Marketplace"]==mp)] if not la.empty else pd.DataFrame()
                sv_rm  = sv[(sv["Region"]==region)&(sv["Marketplace"]==mp)] if not sv.empty else pd.DataFrame()
                sk_rm  = stk[(stk["Region"]==region)&(stk["Marketplace"]==mp)] if not stk.empty else pd.DataFrame()
                brows.append({
                    "Region": region, "Marketplace": mp,
                    "Articles":      _sc(la_rm,"Article No"),
                    "Full New":      _sc(la_rm,"Listing Action","Full New Listing"),
                    "Add Variant":   _sc(la_rm,"Listing Action","Add Variant"),
                    "Already Listed":_sc(la_rm,"Listing Action","Already Listed"),
                    "Active":        _sc(sv_rm,"Expected Status","ACTIVE"),
                    "Inactive":      _sc(sv_rm,"Expected Status","INACTIVE"),
                    "Status PASS":   _sc(sv_rm,"Result","PASS"),
                    "Status FAIL":   _sc(sv_rm,"Result","FAIL"),
                    "Stock FAIL":    _sc(sk_rm,"Result","FAIL"),
                })
        st.dataframe(pd.DataFrame(brows), use_container_width=True)

        st.markdown("### Drilldown")
        t1, t2, t3, t4 = st.tabs([
            "Listing Analysis", "Status Validation", "Stock Validation", "Missing Variants"
        ])

        def drilldown_filters(key_prefix, df):
            c1, c2, c3 = st.columns(3)
            ro = df["Region"].unique().tolist() if not df.empty else []
            mo = df["Marketplace"].unique().tolist() if not df.empty else []
            dr = c1.multiselect("Region",      ro, default=ro, key=f"{key_prefix}_r")
            dm = c2.multiselect("Marketplace", mo, default=mo, key=f"{key_prefix}_m")
            search = c3.text_input("Search Article/EAN", "", key=f"{key_prefix}_s")
            if df.empty: return df
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
            st.dataframe(fla, use_container_width=True, height=400)
        with t2:
            fsv = drilldown_filters("sv", sv)
            rf2 = st.multiselect("Result", ["PASS","FAIL"], default=["PASS","FAIL"], key="sv_res")
            if not fsv.empty and "Result" in fsv.columns:
                fsv = fsv[fsv["Result"].isin(rf2)]
            st.caption(f"{len(fsv):,} records")
            st.dataframe(fsv, use_container_width=True, height=400)
        with t3:
            fstk = drilldown_filters("stk", stk)
            rf3  = st.multiselect("Result", ["PASS","FAIL"], default=["PASS","FAIL"], key="stk_res")
            if not fstk.empty and "Result" in fstk.columns:
                fstk = fstk[fstk["Result"].isin(rf3)]
            st.caption(f"{len(fstk):,} records")
            st.dataframe(fstk, use_container_width=True, height=400)
        with t4:
            fmv = drilldown_filters("mv", mv)
            st.caption(f"{len(fmv):,} missing variants")
            st.dataframe(fmv, use_container_width=True, height=400)

        st.markdown("---")
        st.markdown("### Download Full Report")
        with st.spinner("Building Excel..."):
            xlsx = build_excel(results, regions_run)
        fname = f"PUMA_Audit_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        st.download_button(
            "Download Audit Report (.xlsx)", data=xlsx, file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, type="primary"
        )

# ── DEBUG ─────────────────────────────────────────────────────────────────────
with tab_debug:
    st.markdown("### Inventory Debug")
    dbg = st.session_state["inv_debug"]
    if dbg:
        for region, info in dbg.items():
            st.markdown(f"**{region}**")
            st.dataframe(pd.DataFrame([{"Field":k,"Value":str(v)} for k,v in info.items()]),
                         use_container_width=True)
    else:
        st.info("Run audit first.")
    st.markdown("### Run Log")
    log = st.session_state["run_log"]
    if log:
        for line in log: st.text(line)
    else:
        st.info("Run audit first.")
