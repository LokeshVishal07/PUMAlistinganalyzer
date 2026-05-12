import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import re
from datetime import datetime
import warnings
warnings.filterwarnings("ignore")

st.set_page_config(page_title="PUMA Listing Audit Analyzer", layout="wide", page_icon="📊")

# ─── Styling ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
        padding: 2rem; border-radius: 12px; text-align: center; margin-bottom: 2rem;
        color: white;
    }
    .main-header h1 { font-size: 2.2rem; font-weight: 700; margin: 0; letter-spacing: 1px; }
    .main-header p { color: #a0aec0; margin-top: 0.5rem; }
    .section-card {
        background: #f8fafc; border: 1px solid #e2e8f0;
        border-radius: 10px; padding: 1.2rem; margin-bottom: 1rem;
    }
    .section-title { font-weight: 700; font-size: 1rem; color: #2d3748; margin-bottom: 0.5rem; }
    .metric-box {
        background: white; border-radius: 8px; padding: 1rem;
        border: 1px solid #e2e8f0; text-align: center;
    }
    .metric-value { font-size: 2rem; font-weight: 800; }
    .metric-label { font-size: 0.8rem; color: #718096; margin-top: 0.2rem; }
    .status-active { color: #38a169; }
    .status-inactive { color: #e53e3e; }
    .status-missing { color: #d69e2e; }
    .region-badge {
        display: inline-block; padding: 3px 10px; border-radius: 20px;
        font-size: 0.75rem; font-weight: 600; margin: 2px;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class='main-header'>
    <h1>📊 PUMA Listing Audit Analyzer</h1>
    <p>Marketplace Listing Validation · Status & Stock Audit · MY / SG / PH</p>
</div>
""", unsafe_allow_html=True)

# ─── Constants ──────────────────────────────────────────────────────────────
REGIONS = ["MY", "SG", "PH"]
MARKETPLACES = ["Lazada", "Shopee", "Zalora", "TikTok"]

CHANNEL_BUFFER_PH = {"lazada": 1, "shopee": 0, "zalora": 0}

# ─── File Detection ──────────────────────────────────────────────────────────
def detect_file_type(filename: str) -> tuple[str, str]:
    """Returns (marketplace_or_type, region)"""
    fn = filename.lower()
    if fn.startswith("pricestock"):
        return "lazada", _detect_region(filename)
    if fn.startswith("sellerstatustemplate"):
        return "zalora_status", _detect_region(filename)
    if fn.startswith("sellerstocktemplate"):
        return "zalora_stock", _detect_region(filename)
    if "shopee" in fn and ("masterfile" in fn or "master" in fn):
        return "shopee", _detect_region(filename)
    if fn.startswith("tiktoksellercenter") or "batchedit" in fn:
        return "tiktok", _detect_region(filename)
    if fn.startswith("inventory_"):
        return "inventory", "PH"
    if "puma_my_b2c_channel_inventory" in fn:
        return "inventory", "MY"
    if "sg_puma" in fn and "inventory" in fn:
        return "inventory", "SG"
    if "zecom" in fn or "tracker" in fn or "zrm" in fn:
        return "zecom", "ALL"
    if "content" in fn or "master" in fn or "puma" in fn:
        return "content", "ALL"
    return "unknown", "unknown"

def _detect_region(filename: str) -> str:
    fn = filename.upper()
    for r in REGIONS:
        if r in fn:
            return r
    return "unknown"

# ─── Loaders ─────────────────────────────────────────────────────────────────
def safe_read(file, **kwargs) -> pd.DataFrame:
    try:
        return pd.read_excel(file, **kwargs)
    except Exception:
        try:
            file.seek(0)
            return pd.read_csv(file, **kwargs)
        except Exception as e:
            st.warning(f"Could not read file: {e}")
            return pd.DataFrame()

def normalize_col(df: pd.DataFrame, candidates: list, new_name: str):
    candidates = [str(x).strip().lower() for x in candidates]

    for c in df.columns:
        if str(c).strip().lower() in candidates:
            df = df.rename(columns={c: new_name})
            break

    return df

def load_lazada(file) -> pd.DataFrame:
    df = safe_read(file)
    df = normalize_col(df, ["SellerSKU", "Seller SKU", "seller_sku", "SKU"], "EAN")
    df = normalize_col(df, ["Status", "status", "ItemStatus"], "Marketplace_Status")
    df = normalize_col(df, ["Available", "Stock", "Quantity", "available", "quantity"], "Marketplace_Stock")
    df["Marketplace"] = "Lazada"
    return df[["EAN", "Marketplace_Status", "Marketplace_Stock", "Marketplace"]].dropna(subset=["EAN"])

def load_shopee(file) -> pd.DataFrame:
    df = safe_read(file)
    df = normalize_col(df, ["SKU", "sku", "Seller SKU", "SellerSKU"], "EAN")
    df = normalize_col(df, ["Status", "status", "Listing Status"], "Marketplace_Status")
    df = normalize_col(df, ["Stock", "Quantity", "Available Stock", "stock", "quantity"], "Marketplace_Stock")
    df["Marketplace"] = "Shopee"
    return df[["EAN", "Marketplace_Status", "Marketplace_Stock", "Marketplace"]].dropna(subset=["EAN"])

def load_zalora(status_file, stock_file) -> pd.DataFrame:
    ds = safe_read(status_file)
    dstk = safe_read(stock_file)
    ds = normalize_col(ds, ["SellerSku", "Seller Sku", "SellerSKU", "Seller SKU"], "EAN")
    ds = normalize_col(ds, ["Status", "status", "ItemStatus"], "Marketplace_Status")
    dstk = normalize_col(dstk, ["SellerSku", "Seller Sku", "SellerSKU", "Seller SKU"], "EAN")
    dstk = normalize_col(dstk, ["Stock", "Quantity", "Available", "quantity"], "Marketplace_Stock")
    merged = ds.merge(dstk[["EAN", "Marketplace_Stock"]], on="EAN", how="left")
    merged["Marketplace"] = "Zalora"
    return merged[["EAN", "Marketplace_Status", "Marketplace_Stock", "Marketplace"]].dropna(subset=["EAN"])

def load_tiktok(file) -> pd.DataFrame:
    df = safe_read(file)
    df = normalize_col(df, ["Seller sku", "Seller SKU", "SellerSKU", "SKU", "sku"], "EAN")
    df = normalize_col(df, ["Status", "status", "Product Status"], "Marketplace_Status")
    df = normalize_col(df, ["Stock", "Quantity", "Available", "quantity"], "Marketplace_Stock")
    df["Marketplace"] = "TikTok"
    return df[["EAN", "Marketplace_Status", "Marketplace_Stock", "Marketplace"]].dropna(subset=["EAN"])

def load_inventory(file, region: str) -> pd.DataFrame:
    df = safe_read(file)
    df = normalize_col(df, ["EAN", "ean", "Barcode", "barcode", "SKU", "sku"], "EAN")
    # Try to find stock columns for each channel
    lazada_cols = [c for c in df.columns if "lazada" in c.lower()]
    shopee_cols = [c for c in df.columns if "shopee" in c.lower()]
    zalora_cols = [c for c in df.columns if "zalora" in c.lower()]
    total_cols  = [c for c in df.columns if any(x in c.lower() for x in ["total", "qty", "quantity", "available", "stock"])]

    for c in lazada_cols:  df.rename(columns={c: "Inv_Lazada"}, inplace=True); break
    for c in shopee_cols:  df.rename(columns={c: "Inv_Shopee"}, inplace=True); break
    for c in zalora_cols:  df.rename(columns={c: "Inv_Zalora"}, inplace=True); break
    if "Inv_Lazada" not in df.columns and total_cols:
        df["Inv_Total"] = pd.to_numeric(df[total_cols[0]], errors="coerce").fillna(0)

    # PH channel buffer
    if region == "PH":
        if "Inv_Lazada" in df.columns:
            df["Inv_Lazada"] = pd.to_numeric(df["Inv_Lazada"], errors="coerce").fillna(0)
            df["Inv_Lazada"] = (df["Inv_Lazada"] - CHANNEL_BUFFER_PH["lazada"]).clip(lower=0)
    df["Region"] = region
    return df.dropna(subset=["EAN"])

def load_zecom(file) -> pd.DataFrame:
    df = safe_read(file)
    # Normalize article/color no
    df = normalize_col(df, ["PIM Article#", "PIM Article", "Article No", "ArticleNo", "Color_No",
                             "ColorNo", "Article Number"], "Article_No")
    # Find marketplace columns
    for mp in ["Lazada", "Shopee", "Zalora"]:
        cols = [c for c in df.columns if mp.lower() in str(c).lower()]
        if cols:
            df.rename(columns={cols[0]: f"Tracker_{mp}"}, inplace=True)
    # Launch date
    date_cols = [c for c in df.columns if "launch" in c.lower() or "date" in c.lower()]
    if date_cols:
        df.rename(columns={date_cols[0]: "Launch_Date"}, inplace=True)
        df["Launch_Date"] = pd.to_datetime(df["Launch_Date"], errors="coerce")
    return df.dropna(subset=["Article_No"])

def load_content(file) -> pd.DataFrame:
    df = safe_read(file)
    df = normalize_col(df, ["EAN", "ean", "Barcode", "barcode", "Child SKU", "ChildSKU"], "EAN")
    df = normalize_col(df, ["Color_No", "ColorNo", "Article No", "ArticleNo",
                             "PIM Article#", "Article Number", "Parent SKU"], "Article_No")
    return df.dropna(subset=["EAN"])

# ─── Validation Engine ────────────────────────────────────────────────────────
def is_active_status(val) -> bool:
    if pd.isna(val): return False
    return str(val).strip().upper() in ["ACTIVE", "1", "TRUE", "YES", "ENABLED", "PUBLISHED", "LISTED"]

def tracker_status(val) -> str:
    if pd.isna(val): return "UNKNOWN"
    v = str(val).strip().upper()
    if v == "YES": return "ACTIVE"
    if v in ["NO", "OFF"]: return "INACTIVE"
    return "UNKNOWN"

def determine_expected_status(row, mp: str) -> tuple[str, list[str]]:
    reasons = []
    tracker_col = f"Tracker_{mp}"
    t_status = tracker_status(row.get(tracker_col, np.nan))
    launch = row.get("Launch_Date", np.nan)
    today = pd.Timestamp.today()

    if t_status == "INACTIVE":
        reasons.append(f"Tracker Status is NO/OFF for {mp}")
        return "INACTIVE", reasons
    if pd.notna(launch) and launch > today:
        reasons.append(f"Future Launch Date: {launch.date()}")
        return "INACTIVE", reasons
    if t_status == "ACTIVE":
        return "ACTIVE", reasons
    reasons.append(f"No tracker entry for {mp}")
    return "INACTIVE", reasons

def get_stock_for_mp(inv_row, mp: str, region: str) -> int:
    col_map = {"Lazada": "Inv_Lazada", "Shopee": "Inv_Shopee", "Zalora": "Inv_Zalora", "TikTok": "Inv_Total"}
    col = col_map.get(mp)
    if col and col in inv_row:
        v = inv_row[col]
    elif "Inv_Total" in inv_row:
        v = inv_row["Inv_Total"]
    else:
        return 0
    return max(0, int(pd.to_numeric(v, errors="coerce") or 0))

def run_audit(marketplace_dfs: dict, inventory_df: pd.DataFrame,
              zecom_df: pd.DataFrame, content_df: pd.DataFrame, region: str) -> pd.DataFrame:
    results = []
    today = pd.Timestamp.today()

    # Build EAN → Article_No map from content
    ean_to_art = dict(zip(
        content_df["EAN"].astype(str).str.strip(),
        content_df["Article_No"].astype(str).str.strip()
    ))

    # Build Article_No → zecom row
    zecom_idx = zecom_df.set_index(zecom_df["Article_No"].astype(str).str.strip())

    # Build EAN → inventory row
    inv_idx = inventory_df.set_index(inventory_df["EAN"].astype(str).str.strip()) if not inventory_df.empty else pd.DataFrame()

    for mp, mp_df in marketplace_dfs.items():
        for _, row in mp_df.iterrows():
            ean = str(row["EAN"]).strip()
            mp_status = row.get("Marketplace_Status", np.nan)
            mp_stock  = row.get("Marketplace_Stock", 0)
            art_no    = ean_to_art.get(ean, "NOT IN CONTENT")
            zecom_row = zecom_idx.loc[art_no] if art_no in zecom_idx.index else pd.Series(dtype=object)
            inv_row   = inv_idx.loc[ean] if ean in inv_idx.index else pd.Series(dtype=object)
            if isinstance(inv_row, pd.DataFrame): inv_row = inv_row.iloc[0]
            if isinstance(zecom_row, pd.DataFrame): zecom_row = zecom_row.iloc[0]

            # Expected status from tracker
            expected_status, expected_reasons = determine_expected_status(zecom_row, mp)

            # Stock
            inv_stock = get_stock_for_mp(inv_row, mp, region)
            in_stock  = inv_stock > 0

            # Final expected (must be in stock to be active)
            final_expected = expected_status
            if expected_status == "ACTIVE" and not in_stock:
                final_expected = "INACTIVE"
                expected_reasons.append("No Inventory Stock")

            # Marketplace actual
            is_active_mp = is_active_status(mp_status)
            actual_status = "ACTIVE" if is_active_mp else "INACTIVE"

            # Match?
            status_match = (final_expected == actual_status)

            # Build reasons for not listed / mismatch
            audit_reasons = []
            if art_no == "NOT IN CONTENT":
                audit_reasons.append("EAN not found in Content Master")
            if zecom_row.empty:
                audit_reasons.append("Article not in ZeCom Tracker")
            else:
                audit_reasons.extend(expected_reasons)
            if not status_match:
                if final_expected == "ACTIVE" and actual_status == "INACTIVE":
                    audit_reasons.append("Should be ACTIVE but listed as INACTIVE on marketplace")
                elif final_expected == "INACTIVE" and actual_status == "ACTIVE":
                    audit_reasons.append("Should be INACTIVE but listed as ACTIVE on marketplace")

            results.append({
                "Region": region,
                "Marketplace": mp,
                "EAN (Seller SKU)": ean,
                "Article No (Color No)": art_no,
                "Marketplace Status": actual_status,
                "Marketplace Stock": mp_stock,
                "Inventory Stock": inv_stock,
                "In Stock": "YES" if in_stock else "NO",
                "Tracker Expected Status": final_expected,
                "Status Match": "✓ OK" if status_match else "✗ MISMATCH",
                "Audit Result": "PASS" if status_match else "FAIL",
                "Reasons": " | ".join(audit_reasons) if audit_reasons else "-",
            })
    return pd.DataFrame(results)

# ─── Excel Export ─────────────────────────────────────────────────────────────
def to_excel(all_results: dict) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb = writer.book
        # Formats
        hdr_fmt = wb.add_format({"bold": True, "bg_color": "#0f3460", "font_color": "white",
                                  "border": 1, "align": "center", "valign": "vcenter", "font_name": "Arial", "font_size": 10})
        pass_fmt = wb.add_format({"bg_color": "#c6efce", "font_color": "#276221", "font_name": "Arial", "font_size": 9, "border": 1})
        fail_fmt = wb.add_format({"bg_color": "#ffc7ce", "font_color": "#9c0006", "font_name": "Arial", "font_size": 9, "border": 1})
        warn_fmt = wb.add_format({"bg_color": "#ffeb9c", "font_color": "#9c5700", "font_name": "Arial", "font_size": 9, "border": 1})
        norm_fmt = wb.add_format({"font_name": "Arial", "font_size": 9, "border": 1})
        title_fmt = wb.add_format({"bold": True, "font_size": 14, "font_color": "#0f3460", "font_name": "Arial"})
        summary_hdr = wb.add_format({"bold": True, "bg_color": "#16213e", "font_color": "white",
                                      "font_name": "Arial", "font_size": 10, "border": 1, "align": "center"})
        summary_val = wb.add_format({"font_name": "Arial", "font_size": 10, "border": 1, "align": "center"})

        combined = pd.concat(all_results.values(), ignore_index=True) if all_results else pd.DataFrame()

        # ── Summary Sheet ──
        ws = wb.add_worksheet("📋 Summary")
        writer.sheets["📋 Summary"] = ws
        ws.set_column("A:A", 22)
        ws.set_column("B:G", 15)
        ws.write("A1", "PUMA Listing Audit – Summary Report", title_fmt)
        ws.write("A2", f"Generated: {datetime.now().strftime('%d %b %Y %H:%M')}", norm_fmt)

        row = 4
        if not combined.empty:
            ws.write(row, 0, "Region", summary_hdr)
            ws.write(row, 1, "Marketplace", summary_hdr)
            ws.write(row, 2, "Total SKUs", summary_hdr)
            ws.write(row, 3, "PASS", summary_hdr)
            ws.write(row, 4, "FAIL", summary_hdr)
            ws.write(row, 5, "ACTIVE", summary_hdr)
            ws.write(row, 6, "INACTIVE", summary_hdr)
            ws.write(row, 7, "In Stock", summary_hdr)
            ws.write(row, 8, "No Stock", summary_hdr)
            row += 1
            for (region, mp), grp in combined.groupby(["Region", "Marketplace"]):
                ws.write(row, 0, region, summary_val)
                ws.write(row, 1, mp, summary_val)
                ws.write(row, 2, len(grp), summary_val)
                ws.write(row, 3, (grp["Audit Result"] == "PASS").sum(), pass_fmt)
                ws.write(row, 4, (grp["Audit Result"] == "FAIL").sum(), fail_fmt)
                ws.write(row, 5, (grp["Marketplace Status"] == "ACTIVE").sum(), summary_val)
                ws.write(row, 6, (grp["Marketplace Status"] == "INACTIVE").sum(), summary_val)
                ws.write(row, 7, (grp["In Stock"] == "YES").sum(), summary_val)
                ws.write(row, 8, (grp["In Stock"] == "NO").sum(), summary_val)
                row += 1

        # ── Detail sheets per region ──
        for region, df in all_results.items():
            sname = f"{region} Detail"
            df.to_excel(writer, sheet_name=sname, index=False, startrow=1)
            ws2 = writer.sheets[sname]
            ws2.write(0, 0, f"Region: {region} – Detailed Audit", title_fmt)
            cols = list(df.columns)
            for ci, col in enumerate(cols):
                ws2.write(1, ci, col, hdr_fmt)
                ws2.set_column(ci, ci, 22)
            for ri, record in df.iterrows():
                for ci, col in enumerate(cols):
                    val = record[col]
                    if col == "Audit Result":
                        fmt = pass_fmt if val == "PASS" else fail_fmt
                    elif col == "Status Match":
                        fmt = pass_fmt if "OK" in str(val) else fail_fmt
                    elif col == "In Stock":
                        fmt = pass_fmt if val == "YES" else warn_fmt
                    else:
                        fmt = norm_fmt
                    ws2.write(ri + 2, ci, str(val) if not pd.isna(val) else "", fmt)

        # ── Mismatch / Fail Sheet ──
        if not combined.empty:
            fails = combined[combined["Audit Result"] == "FAIL"]
            if not fails.empty:
                fails.to_excel(writer, sheet_name="⚠ Mismatches", index=False, startrow=1)
                ws3 = writer.sheets["⚠ Mismatches"]
                ws3.write(0, 0, "Status Mismatches & Issues", title_fmt)
                for ci, col in enumerate(fails.columns):
                    ws3.write(1, ci, col, hdr_fmt)
                    ws3.set_column(ci, ci, 22)

    return output.getvalue()

# ─── UI ────────────────────────────────────────────────────────────────────────
tab_upload, tab_results, tab_help = st.tabs(["📁 Upload Files", "📊 Results & Download", "❓ Help & Guide"])

# Session state
if "audit_results" not in st.session_state:
    st.session_state.audit_results = {}
if "uploaded_files_info" not in st.session_state:
    st.session_state.uploaded_files_info = []

with tab_help:
    st.markdown("## 📖 File Naming Guide")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("""
**Lazada** → file starts with `pricestock`  
**Shopee** → file contains `Shopee` + `Masterfile` or `Master`  
**Zalora Status** → file starts with `SellerStatusTemplate`  
**Zalora Stock** → file starts with `SellerStockTemplate`  
**TikTok** → file starts with `TikTokSellerCenter` or contains `batchedit`
        """)
    with col2:
        st.markdown("""
**Inventory PH** → starts with `Inventory_`  
**Inventory MY** → starts with `PUMA_MY_B2C_Channel_Inventory_`  
**Inventory SG** → starts with `SG_PUMA SG B2C Inventory Rpt_New_`  
**ZeCom Tracker** → contains `zecom` or `tracker`  
**Content/Master** → contains `content` or `master` or `puma`
        """)
    st.markdown("""
---
## ✅ Active / Inactive Logic
| Condition | Expected Status |
|---|---|
| Tracker = YES + In Stock + Past Launch | **ACTIVE** |
| Tracker = NO or OFF | **INACTIVE** |
| No Stock | **INACTIVE** |
| Future Launch Date | **INACTIVE** |
| Special Brand Request | **INACTIVE** |

## 🔑 Key Identifiers
- **EAN** = Seller SKU = Child level identifier  
- **Article No / Color No** = PIM Article# = Parent level identifier  
- Zalora: `SellerSku`, Lazada: `SellerSKU`, Shopee: `SKU`, TikTok: `Seller sku`
    """)

with tab_upload:
    st.markdown("### Step 1 — Select Region(s)")
    selected_regions = st.multiselect("Regions to audit:", REGIONS, default=REGIONS)

    st.markdown("### Step 2 — Upload Master Files (shared across regions)")
    col_z, col_c = st.columns(2)
    with col_z:
        zecom_file = st.file_uploader("📋 ZeCom Tracker (Excel)", type=["xlsx", "xls", "csv"],
                                      key="zecom", help="Contains Tracker Status YES/NO/OFF per marketplace")
    with col_c:
        content_file = st.file_uploader("📦 Content Master File (Excel)", type=["xlsx", "xls", "csv"],
                                        key="content", help="Master dump with EAN → Article No mapping")

    st.markdown("### Step 3 — Upload Region Files")
    region_files = {}
    for region in selected_regions:
        with st.expander(f"📂 {region} Files", expanded=True):
            region_files[region] = {}
            c1, c2 = st.columns(2)
            with c1:
                region_files[region]["lazada"] = st.file_uploader(
                    f"Lazada ({region}) — pricestock*.xlsx", type=["xlsx","xls","csv"], key=f"lazada_{region}")
                region_files[region]["shopee"] = st.file_uploader(
                    f"Shopee ({region}) — Shopee*Masterfile*.xlsx", type=["xlsx","xls","csv"], key=f"shopee_{region}")
            with c2:
                region_files[region]["zalora_status"] = st.file_uploader(
                    f"Zalora Status ({region}) — SellerStatusTemplate*.xlsx", type=["xlsx","xls","csv"], key=f"zstatus_{region}")
                region_files[region]["zalora_stock"] = st.file_uploader(
                    f"Zalora Stock ({region}) — SellerStockTemplate*.xlsx", type=["xlsx","xls","csv"], key=f"zstock_{region}")
                region_files[region]["tiktok"] = st.file_uploader(
                    f"TikTok ({region}) — TikTokSellerCenter*.xlsx", type=["xlsx","xls","csv"], key=f"tiktok_{region}")
            region_files[region]["inventory"] = st.file_uploader(
                f"Inventory ({region})", type=["xlsx","xls","csv"], key=f"inv_{region}",
                help="PH: Inventory_* | MY: PUMA_MY_B2C_Channel_Inventory_* | SG: SG_PUMA SG B2C Inventory Rpt_New_*")

    st.markdown("---")
    run_btn = st.button("🚀 Run Listing Audit", type="primary", use_container_width=True)

    if run_btn:
        if not zecom_file or not content_file:
            st.error("Please upload both ZeCom Tracker and Content Master File.")
        else:
            with st.spinner("Loading master files..."):
                zecom_df  = load_zecom(zecom_file)
                content_df = load_content(content_file)
                st.success(f"✅ Content: {len(content_df):,} SKUs | ZeCom Tracker: {len(zecom_df):,} articles")

            all_results = {}
            for region in selected_regions:
                rf = region_files.get(region, {})
                marketplace_dfs = {}

                if rf.get("lazada"):
                    marketplace_dfs["Lazada"] = load_lazada(rf["lazada"])
                if rf.get("shopee"):
                    marketplace_dfs["Shopee"] = load_shopee(rf["shopee"])
                if rf.get("zalora_status") and rf.get("zalora_stock"):
                    marketplace_dfs["Zalora"] = load_zalora(rf["zalora_status"], rf["zalora_stock"])
                elif rf.get("zalora_status"):
                    st.warning(f"[{region}] Zalora: Stock file missing, skipping Zalora audit.")
                if rf.get("tiktok"):
                    marketplace_dfs["TikTok"] = load_tiktok(rf["tiktok"])

                inv_df = pd.DataFrame()
                if rf.get("inventory"):
                    inv_df = load_inventory(rf["inventory"], region)
                    st.info(f"[{region}] Inventory: {len(inv_df):,} rows loaded")

                if marketplace_dfs:
                    total_mp_skus = sum(len(v) for v in marketplace_dfs.values())
                    st.info(f"[{region}] Marketplace SKUs: {total_mp_skus:,} across {list(marketplace_dfs.keys())}")
                    with st.spinner(f"Running audit for {region}..."):
                        result_df = run_audit(marketplace_dfs, inv_df, zecom_df, content_df, region)
                    all_results[region] = result_df
                    st.success(f"✅ [{region}] Audit complete — {len(result_df):,} records")
                else:
                    st.warning(f"[{region}] No marketplace files uploaded, skipping.")

            st.session_state.audit_results = all_results
            if all_results:
                st.success("🎉 All audits complete! Go to **Results & Download** tab.")

with tab_results:
    results = st.session_state.audit_results
    if not results:
        st.info("Run the audit first in the **Upload Files** tab.")
    else:
        combined = pd.concat(results.values(), ignore_index=True)

        # ── KPI Summary ──
        st.markdown("### 📊 Audit Overview")
        k1, k2, k3, k4, k5 = st.columns(5)
        total   = len(combined)
        passed  = (combined["Audit Result"] == "PASS").sum()
        failed  = (combined["Audit Result"] == "FAIL").sum()
        active  = (combined["Marketplace Status"] == "ACTIVE").sum()
        instock = (combined["In Stock"] == "YES").sum()
        k1.markdown(f"<div class='metric-box'><div class='metric-value'>{total:,}</div><div class='metric-label'>Total SKUs</div></div>", unsafe_allow_html=True)
        k2.markdown(f"<div class='metric-box'><div class='metric-value status-active'>{passed:,}</div><div class='metric-label'>✓ Pass</div></div>", unsafe_allow_html=True)
        k3.markdown(f"<div class='metric-box'><div class='metric-value status-inactive'>{failed:,}</div><div class='metric-label'>✗ Fail / Mismatch</div></div>", unsafe_allow_html=True)
        k4.markdown(f"<div class='metric-box'><div class='metric-value'>{active:,}</div><div class='metric-label'>Active on MP</div></div>", unsafe_allow_html=True)
        k5.markdown(f"<div class='metric-box'><div class='metric-value status-missing'>{instock:,}</div><div class='metric-label'>In Stock</div></div>", unsafe_allow_html=True)

        st.markdown("### 📋 Region × Marketplace Breakdown")
        pivot = combined.groupby(["Region", "Marketplace"]).agg(
            Total=("EAN (Seller SKU)", "count"),
            Pass=("Audit Result", lambda x: (x == "PASS").sum()),
            Fail=("Audit Result", lambda x: (x == "FAIL").sum()),
            Active=("Marketplace Status", lambda x: (x == "ACTIVE").sum()),
            Inactive=("Marketplace Status", lambda x: (x == "INACTIVE").sum()),
            In_Stock=("In Stock", lambda x: (x == "YES").sum()),
            No_Stock=("In Stock", lambda x: (x == "NO").sum()),
        ).reset_index()
        st.dataframe(pivot, use_container_width=True)

        st.markdown("### 🔍 Detailed Results")
        c1, c2, c3 = st.columns(3)
        filt_region = c1.multiselect("Region", options=combined["Region"].unique().tolist(), default=combined["Region"].unique().tolist())
        filt_mp     = c2.multiselect("Marketplace", options=combined["Marketplace"].unique().tolist(), default=combined["Marketplace"].unique().tolist())
        filt_result = c3.multiselect("Audit Result", options=["PASS", "FAIL"], default=["PASS","FAIL"])

        filtered = combined[
            combined["Region"].isin(filt_region) &
            combined["Marketplace"].isin(filt_mp) &
            combined["Audit Result"].isin(filt_result)
        ]
        st.dataframe(filtered, use_container_width=True, height=400)

        st.markdown("### ⚠️ Top Failure Reasons")
        fails = combined[combined["Audit Result"] == "FAIL"]
        if not fails.empty:
            all_reasons = fails["Reasons"].str.split(" | ").explode().str.strip()
            reason_counts = all_reasons.value_counts().head(10).reset_index()
            reason_counts.columns = ["Reason", "Count"]
            st.dataframe(reason_counts, use_container_width=True)

        st.markdown("---")
        st.markdown("### 💾 Download Report")
        excel_bytes = to_excel(results)
        fname = f"PUMA_Listing_Audit_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        st.download_button(
            label="📥 Download Full Audit Report (.xlsx)",
            data=excel_bytes,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )
        st.caption("Report includes: Summary sheet, per-region detail, and all mismatches tab.")
