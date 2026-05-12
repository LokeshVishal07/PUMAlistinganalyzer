import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
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
    .metric-box {
        background: white; border-radius: 8px; padding: 1rem;
        border: 1px solid #e2e8f0; text-align: center;
    }
    .metric-value { font-size: 2rem; font-weight: 800; }
    .metric-label { font-size: 0.8rem; color: #718096; margin-top: 0.2rem; }
    .status-active  { color: #38a169; }
    .status-inactive{ color: #e53e3e; }
    .status-missing { color: #d69e2e; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class='main-header'>
    <h1>📊 PUMA Listing Audit Analyzer</h1>
    <p>Marketplace Listing Validation · Status & Stock Audit · MY / SG / PH</p>
</div>
""", unsafe_allow_html=True)

# ─── Constants ───────────────────────────────────────────────────────────────
REGIONS = ["MY", "SG", "PH"]
CHANNEL_BUFFER_PH = {"lazada": 1, "shopee": 0, "zalora": 0}

# ─── Core Helpers ─────────────────────────────────────────────────────────────

def str_col(c) -> str:
    """Convert ANY column label (int, float, str, NaN) to a clean string."""
    try:
        s = str(c).strip()
        return s if s not in ("", "nan", "None") else ""
    except Exception:
        return ""

def col_lower(c) -> str:
    return str_col(c).lower()

def sanitize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Force every column label to a non-empty string. Must be called after every read."""
    new_cols = []
    for i, c in enumerate(df.columns):
        s = str_col(c)
        new_cols.append(s if s else f"_col_{i}")
    df.columns = new_cols
    return df

def find_header_row(file, max_scan: int = 15) -> int:
    """
    Detect which row is the real header by finding the row with the most
    non-empty string cells. Returns 0-based row index.
    """
    try:
        preview = pd.read_excel(file, header=None, nrows=max_scan)
        file.seek(0)
        best_row, best_score = 0, -1
        for i, row in preview.iterrows():
            score = row.dropna().apply(
                lambda v: isinstance(v, str) and len(v.strip()) > 0
            ).sum()
            if score > best_score:
                best_score, best_row = score, int(i)
        return best_row
    except Exception:
        try:
            file.seek(0)
        except Exception:
            pass
        return 0

def safe_read(file) -> pd.DataFrame:
    """
    Read Excel or CSV, auto-detect header row, sanitize all column labels.
    Guaranteed to return a DataFrame with string column names.
    """
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
                st.warning(f"Could not read file: {e}")
                return pd.DataFrame()

    df = sanitize_columns(df)
    df = df.dropna(how="all").reset_index(drop=True)
    return df

def normalize_col(df: pd.DataFrame, candidates: list, new_name: str) -> pd.DataFrame:
    """Rename the first column whose lowercase-stripped name matches a candidate."""
    c_lower = [str(x).strip().lower() for x in candidates]
    for col in df.columns:
        if col_lower(col) in c_lower:
            df = df.rename(columns={col: new_name})
            break
    return df

# ─── File Loaders ─────────────────────────────────────────────────────────────

def load_lazada(file) -> pd.DataFrame:
    df = safe_read(file)
    df = normalize_col(df, ["SellerSKU", "Seller SKU", "Seller Sku", "seller_sku", "SKU", "sku"], "EAN")
    df = normalize_col(df, ["Status", "status", "ItemStatus", "Item Status", "Listing Status"], "Marketplace_Status")
    df = normalize_col(df, ["Available", "Stock", "Quantity", "available", "quantity", "Qty"], "Marketplace_Stock")
    for col in ["EAN", "Marketplace_Status", "Marketplace_Stock"]:
        if col not in df.columns:
            df[col] = np.nan
    df["Marketplace"] = "Lazada"
    return df[["EAN", "Marketplace_Status", "Marketplace_Stock", "Marketplace"]].dropna(subset=["EAN"])

def load_shopee(file) -> pd.DataFrame:
    df = safe_read(file)
    df = normalize_col(df, ["SKU", "sku", "Seller SKU", "SellerSKU", "Seller Sku"], "EAN")
    df = normalize_col(df, ["Status", "status", "Listing Status", "Product Status"], "Marketplace_Status")
    df = normalize_col(df, ["Stock", "Quantity", "Available Stock", "stock", "quantity", "Qty"], "Marketplace_Stock")
    for col in ["EAN", "Marketplace_Status", "Marketplace_Stock"]:
        if col not in df.columns:
            df[col] = np.nan
    df["Marketplace"] = "Shopee"
    return df[["EAN", "Marketplace_Status", "Marketplace_Stock", "Marketplace"]].dropna(subset=["EAN"])

def load_zalora(status_file, stock_file) -> pd.DataFrame:
    ds   = safe_read(status_file)
    dstk = safe_read(stock_file)
    ds   = normalize_col(ds,   ["SellerSku", "Seller Sku", "SellerSKU", "Seller SKU", "seller_sku"], "EAN")
    ds   = normalize_col(ds,   ["Status", "status", "ItemStatus", "Item Status"], "Marketplace_Status")
    dstk = normalize_col(dstk, ["SellerSku", "Seller Sku", "SellerSKU", "Seller SKU", "seller_sku"], "EAN")
    dstk = normalize_col(dstk, ["Stock", "Quantity", "Available", "quantity", "Qty"], "Marketplace_Stock")
    for col in ["EAN", "Marketplace_Status"]:
        if col not in ds.columns:
            ds[col] = np.nan
    if "EAN" not in dstk.columns:
        dstk["EAN"] = np.nan
    if "Marketplace_Stock" not in dstk.columns:
        dstk["Marketplace_Stock"] = np.nan
    merged = ds.merge(dstk[["EAN", "Marketplace_Stock"]], on="EAN", how="left")
    merged["Marketplace"] = "Zalora"
    return merged[["EAN", "Marketplace_Status", "Marketplace_Stock", "Marketplace"]].dropna(subset=["EAN"])

def load_tiktok(file) -> pd.DataFrame:
    df = safe_read(file)
    df = normalize_col(df, ["Seller sku", "Seller SKU", "SellerSKU", "Seller Sku", "SKU", "sku"], "EAN")
    df = normalize_col(df, ["Status", "status", "Product Status", "Item Status"], "Marketplace_Status")
    df = normalize_col(df, ["Stock", "Quantity", "Available", "quantity", "Qty"], "Marketplace_Stock")
    for col in ["EAN", "Marketplace_Status", "Marketplace_Stock"]:
        if col not in df.columns:
            df[col] = np.nan
    df["Marketplace"] = "TikTok"
    return df[["EAN", "Marketplace_Status", "Marketplace_Stock", "Marketplace"]].dropna(subset=["EAN"])

def load_inventory(file, region: str) -> pd.DataFrame:
    df = safe_read(file)
    df = normalize_col(df, ["EAN", "ean", "Barcode", "barcode", "SKU", "sku", "Item Code"], "EAN")

    lazada_cols = [c for c in df.columns if "lazada" in col_lower(c)]
    shopee_cols = [c for c in df.columns if "shopee" in col_lower(c)]
    zalora_cols = [c for c in df.columns if "zalora" in col_lower(c)]
    total_cols  = [c for c in df.columns if any(
        x in col_lower(c) for x in ["total", "qty", "quantity", "available", "stock", "on hand"]
    )]

    if lazada_cols:
        df = df.rename(columns={lazada_cols[0]: "Inv_Lazada"})
    if shopee_cols:
        df = df.rename(columns={shopee_cols[0]: "Inv_Shopee"})
    if zalora_cols:
        df = df.rename(columns={zalora_cols[0]: "Inv_Zalora"})
    if "Inv_Lazada" not in df.columns and total_cols:
        df["Inv_Total"] = pd.to_numeric(df[total_cols[0]], errors="coerce").fillna(0)

    if region == "PH" and "Inv_Lazada" in df.columns:
        df["Inv_Lazada"] = (
            pd.to_numeric(df["Inv_Lazada"], errors="coerce")
            .fillna(0)
            .sub(CHANNEL_BUFFER_PH["lazada"])
            .clip(lower=0)
        )

    if "EAN" not in df.columns:
        df["EAN"] = np.nan
    df["Region"] = region
    return df.dropna(subset=["EAN"])

def load_zecom(file) -> pd.DataFrame:
    df = safe_read(file)
    df = normalize_col(df, [
        "PIM Article#", "PIM Article", "Article No", "ArticleNo",
        "Color_No", "ColorNo", "Article Number", "Color No", "SKU"
    ], "Article_No")

    # Tracker columns for each marketplace
    for mp in ["Lazada", "Shopee", "Zalora"]:
        cols = [c for c in df.columns if mp.lower() in col_lower(c)]
        if cols:
            df = df.rename(columns={cols[0]: f"Tracker_{mp}"})

    # Launch date — only match string column names containing "launch" or "date"
    date_cols = [
        c for c in df.columns
        if "launch" in col_lower(c) or "date" in col_lower(c)
    ]
    if date_cols:
        df = df.rename(columns={date_cols[0]: "Launch_Date"})
        df["Launch_Date"] = pd.to_datetime(df["Launch_Date"], errors="coerce")

    if "Article_No" not in df.columns:
        df["Article_No"] = np.nan
    return df.dropna(subset=["Article_No"])

def load_content(file) -> pd.DataFrame:
    df = safe_read(file)
    df = normalize_col(df, [
        "EAN", "ean", "Barcode", "barcode", "Child SKU", "ChildSKU",
        "Seller SKU", "SellerSKU"
    ], "EAN")
    df = normalize_col(df, [
        "Color_No", "ColorNo", "Article No", "ArticleNo",
        "PIM Article#", "Article Number", "Parent SKU", "Color No"
    ], "Article_No")
    for col in ["EAN", "Article_No"]:
        if col not in df.columns:
            df[col] = np.nan
    return df.dropna(subset=["EAN"])

# ─── Validation Logic ─────────────────────────────────────────────────────────

def is_active_status(val) -> bool:
    if pd.isna(val):
        return False
    return str(val).strip().upper() in [
        "ACTIVE", "1", "TRUE", "YES", "ENABLED", "PUBLISHED", "LISTED", "NORMAL"
    ]

def tracker_status(val) -> str:
    if pd.isna(val):
        return "UNKNOWN"
    v = str(val).strip().upper()
    if v == "YES":
        return "ACTIVE"
    if v in ["NO", "OFF"]:
        return "INACTIVE"
    return "UNKNOWN"

def determine_expected_status(zecom_row: pd.Series, mp: str):
    reasons      = []
    tracker_col  = f"Tracker_{mp}"
    t_val        = zecom_row[tracker_col] if tracker_col in zecom_row.index else np.nan
    t_status     = tracker_status(t_val)
    launch_val   = zecom_row["Launch_Date"] if "Launch_Date" in zecom_row.index else np.nan
    today        = pd.Timestamp.today()

    if t_status == "INACTIVE":
        reasons.append(f"Tracker Status is NO/OFF for {mp}")
        return "INACTIVE", reasons

    if pd.notna(launch_val):
        try:
            if pd.Timestamp(launch_val) > today:
                reasons.append(f"Future Launch Date: {pd.Timestamp(launch_val).date()}")
                return "INACTIVE", reasons
        except Exception:
            pass

    if t_status == "ACTIVE":
        return "ACTIVE", reasons

    reasons.append(f"No tracker entry for {mp}")
    return "INACTIVE", reasons

def get_stock_for_mp(inv_row: pd.Series, mp: str) -> int:
    col_map = {
        "Lazada": "Inv_Lazada",
        "Shopee": "Inv_Shopee",
        "Zalora": "Inv_Zalora",
        "TikTok": "Inv_Total",
    }
    col = col_map.get(mp, "")
    val = inv_row[col] if col in inv_row.index else (
          inv_row["Inv_Total"] if "Inv_Total" in inv_row.index else np.nan)
    v   = pd.to_numeric(val, errors="coerce")
    return max(0, int(v)) if pd.notna(v) else 0

# ─── Audit Engine ─────────────────────────────────────────────────────────────

def run_audit(marketplace_dfs: dict, inventory_df: pd.DataFrame,
              zecom_df: pd.DataFrame, content_df: pd.DataFrame, region: str) -> pd.DataFrame:
    results = []

    ean_to_art = dict(zip(
        content_df["EAN"].astype(str).str.strip(),
        content_df["Article_No"].astype(str).str.strip()
    ))
    zecom_idx = zecom_df.set_index(zecom_df["Article_No"].astype(str).str.strip())
    inv_idx   = (
        inventory_df.set_index(inventory_df["EAN"].astype(str).str.strip())
        if not inventory_df.empty else pd.DataFrame()
    )

    empty_series = pd.Series(dtype=object)

    for mp, mp_df in marketplace_dfs.items():
        for _, row in mp_df.iterrows():
            ean = str(row.get("EAN", "")).strip()
            if not ean or ean.lower() in ("nan", "none", ""):
                continue

            mp_status = row.get("Marketplace_Status", np.nan)
            mp_stock  = row.get("Marketplace_Stock", 0)
            art_no    = ean_to_art.get(ean, "NOT IN CONTENT")

            zecom_row = zecom_idx.loc[art_no] if art_no in zecom_idx.index else empty_series
            inv_row   = inv_idx.loc[ean]       if ean   in inv_idx.index   else empty_series
            if isinstance(zecom_row, pd.DataFrame): zecom_row = zecom_row.iloc[0]
            if isinstance(inv_row,   pd.DataFrame): inv_row   = inv_row.iloc[0]

            expected_status, expected_reasons = determine_expected_status(zecom_row, mp)

            inv_stock = get_stock_for_mp(inv_row, mp) if not inv_row.empty else 0
            in_stock  = inv_stock > 0

            final_expected = expected_status
            if expected_status == "ACTIVE" and not in_stock:
                final_expected = "INACTIVE"
                expected_reasons.append("No Inventory Stock")

            actual_status = "ACTIVE" if is_active_status(mp_status) else "INACTIVE"
            status_match  = (final_expected == actual_status)

            audit_reasons = []
            if art_no == "NOT IN CONTENT":
                audit_reasons.append("EAN not found in Content Master")
            if zecom_row.empty:
                audit_reasons.append("Article not in ZeCom Tracker")
            else:
                audit_reasons.extend(expected_reasons)
            if not status_match:
                if final_expected == "ACTIVE" and actual_status == "INACTIVE":
                    audit_reasons.append("Should be ACTIVE but INACTIVE on marketplace")
                else:
                    audit_reasons.append("Should be INACTIVE but ACTIVE on marketplace")

            results.append({
                "Region":                  region,
                "Marketplace":             mp,
                "EAN (Seller SKU)":        ean,
                "Article No (Color No)":   art_no,
                "Marketplace Status":      actual_status,
                "Marketplace Stock":       mp_stock,
                "Inventory Stock":         inv_stock,
                "In Stock":                "YES" if in_stock else "NO",
                "Tracker Expected Status": final_expected,
                "Status Match":            "OK" if status_match else "MISMATCH",
                "Audit Result":            "PASS" if status_match else "FAIL",
                "Reasons":                 " | ".join(audit_reasons) if audit_reasons else "-",
            })

    return pd.DataFrame(results)

# ─── Excel Export ─────────────────────────────────────────────────────────────

def to_excel(all_results: dict) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb = writer.book
        hdr_fmt  = wb.add_format({"bold": True, "bg_color": "#0f3460", "font_color": "white",
                                   "border": 1, "align": "center", "valign": "vcenter",
                                   "font_name": "Arial", "font_size": 10})
        pass_fmt = wb.add_format({"bg_color": "#c6efce", "font_color": "#276221",
                                   "font_name": "Arial", "font_size": 9, "border": 1})
        fail_fmt = wb.add_format({"bg_color": "#ffc7ce", "font_color": "#9c0006",
                                   "font_name": "Arial", "font_size": 9, "border": 1})
        warn_fmt = wb.add_format({"bg_color": "#ffeb9c", "font_color": "#9c5700",
                                   "font_name": "Arial", "font_size": 9, "border": 1})
        norm_fmt = wb.add_format({"font_name": "Arial", "font_size": 9, "border": 1})
        ttl_fmt  = wb.add_format({"bold": True, "font_size": 14,
                                   "font_color": "#0f3460", "font_name": "Arial"})
        shdr_fmt = wb.add_format({"bold": True, "bg_color": "#16213e", "font_color": "white",
                                   "font_name": "Arial", "font_size": 10, "border": 1, "align": "center"})
        sval_fmt = wb.add_format({"font_name": "Arial", "font_size": 10, "border": 1, "align": "center"})

        combined = pd.concat(all_results.values(), ignore_index=True) if all_results else pd.DataFrame()

        # Summary
        ws = wb.add_worksheet("Summary")
        writer.sheets["Summary"] = ws
        ws.set_column("A:I", 18)
        ws.write("A1", "PUMA Listing Audit - Summary Report", ttl_fmt)
        ws.write("A2", f"Generated: {datetime.now().strftime('%d %b %Y %H:%M')}", norm_fmt)
        r = 4
        if not combined.empty:
            for ci, h in enumerate(["Region","Marketplace","Total SKUs",
                                     "PASS","FAIL","ACTIVE","INACTIVE","In Stock","No Stock"]):
                ws.write(r, ci, h, shdr_fmt)
            r += 1
            for (reg, mp), grp in combined.groupby(["Region","Marketplace"]):
                ws.write(r, 0, reg,      sval_fmt)
                ws.write(r, 1, mp,       sval_fmt)
                ws.write(r, 2, len(grp), sval_fmt)
                ws.write(r, 3, int((grp["Audit Result"]       == "PASS").sum()),     pass_fmt)
                ws.write(r, 4, int((grp["Audit Result"]       == "FAIL").sum()),     fail_fmt)
                ws.write(r, 5, int((grp["Marketplace Status"] == "ACTIVE").sum()),   sval_fmt)
                ws.write(r, 6, int((grp["Marketplace Status"] == "INACTIVE").sum()), sval_fmt)
                ws.write(r, 7, int((grp["In Stock"]           == "YES").sum()),      sval_fmt)
                ws.write(r, 8, int((grp["In Stock"]           == "NO").sum()),       sval_fmt)
                r += 1

        # Per-region sheets
        for region, df in all_results.items():
            sname = f"{region} Detail"
            df.to_excel(writer, sheet_name=sname, index=False, startrow=1)
            ws2 = writer.sheets[sname]
            ws2.write(0, 0, f"Region: {region} - Detailed Audit", ttl_fmt)
            for ci, col in enumerate(df.columns):
                ws2.write(1, ci, col, hdr_fmt)
                ws2.set_column(ci, ci, 22)
            for ri, record in df.iterrows():
                for ci, col in enumerate(df.columns):
                    val      = record[col]
                    safe_val = "" if (isinstance(val, float) and np.isnan(val)) else str(val)
                    if col == "Audit Result":
                        fmt = pass_fmt if val == "PASS" else fail_fmt
                    elif col == "Status Match":
                        fmt = pass_fmt if val == "OK" else fail_fmt
                    elif col == "In Stock":
                        fmt = pass_fmt if val == "YES" else warn_fmt
                    else:
                        fmt = norm_fmt
                    ws2.write(ri + 2, ci, safe_val, fmt)

        # Mismatches
        if not combined.empty:
            fails = combined[combined["Audit Result"] == "FAIL"]
            if not fails.empty:
                fails.to_excel(writer, sheet_name="Mismatches", index=False, startrow=1)
                ws3 = writer.sheets["Mismatches"]
                ws3.write(0, 0, "Status Mismatches and Issues", ttl_fmt)
                for ci, col in enumerate(fails.columns):
                    ws3.write(1, ci, col, hdr_fmt)
                    ws3.set_column(ci, ci, 22)

    return output.getvalue()

# ─── UI ───────────────────────────────────────────────────────────────────────
tab_upload, tab_results, tab_help = st.tabs([
    "📁 Upload Files", "📊 Results & Download", "❓ Help & Guide"
])

if "audit_results" not in st.session_state:
    st.session_state.audit_results = {}

# HELP
with tab_help:
    st.markdown("## 📖 File Naming Guide")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("""
**Lazada** → starts with `pricestock`  
**Shopee** → contains `Shopee` + `Masterfile`  
**Zalora Status** → starts with `SellerStatusTemplate`  
**Zalora Stock** → starts with `SellerStockTemplate`  
**TikTok** → starts with `TikTokSellerCenter` or contains `batchedit`
        """)
    with c2:
        st.markdown("""
**Inventory PH** → starts with `Inventory_`  
**Inventory MY** → starts with `PUMA_MY_B2C_Channel_Inventory_`  
**Inventory SG** → starts with `SG_PUMA SG B2C Inventory Rpt_New_`  
**ZeCom Tracker** → contains `zecom` or `tracker`  
**Content File** → master dump: EAN (child) ↔ Article No (parent)
        """)
    st.markdown("""
---
## ✅ Active / Inactive Logic

| Condition | Expected Status |
|---|---|
| Tracker = YES + In Stock + Past Launch Date | **ACTIVE** |
| Tracker = NO or OFF | **INACTIVE** |
| No inventory stock | **INACTIVE** |
| Future launch date | **INACTIVE** |

## 🔑 Key Identifiers
- **EAN** = Seller SKU → child-level identifier  
- **Article No / Color No** = PIM Article# → parent-level identifier  
- PH Lazada channel buffer: **-1 unit** applied automatically
    """)

# UPLOAD
with tab_upload:
    st.markdown("### Step 1 — Select Region(s)")
    selected_regions = st.multiselect("Regions to audit:", REGIONS, default=REGIONS)

    st.markdown("### Step 2 — Upload Master Files")
    col_z, col_c = st.columns(2)
    with col_z:
        zecom_file = st.file_uploader(
            "📋 ZeCom Tracker", type=["xlsx","xls","csv"], key="zecom",
            help="Tracker with YES/NO/OFF per marketplace per article"
        )
    with col_c:
        content_file = st.file_uploader(
            "📦 Content Master File", type=["xlsx","xls","csv"], key="content",
            help="Master dump: EAN (child) ↔ Article No (parent)"
        )

    st.markdown("### Step 3 — Upload Region Files")
    region_files = {}
    for region in selected_regions:
        with st.expander(f"📂 {region} Files", expanded=True):
            region_files[region] = {}
            c1, c2 = st.columns(2)
            with c1:
                region_files[region]["lazada"] = st.file_uploader(
                    f"Lazada ({region})  —  pricestock*.xlsx",
                    type=["xlsx","xls","csv"], key=f"laz_{region}"
                )
                region_files[region]["shopee"] = st.file_uploader(
                    f"Shopee ({region})  —  Shopee*Masterfile*.xlsx",
                    type=["xlsx","xls","csv"], key=f"sho_{region}"
                )
                region_files[region]["tiktok"] = st.file_uploader(
                    f"TikTok ({region})  —  TikTokSellerCenter*.xlsx",
                    type=["xlsx","xls","csv"], key=f"ttk_{region}"
                )
            with c2:
                region_files[region]["zalora_status"] = st.file_uploader(
                    f"Zalora Status ({region})  —  SellerStatusTemplate*.xlsx",
                    type=["xlsx","xls","csv"], key=f"zst_{region}"
                )
                region_files[region]["zalora_stock"] = st.file_uploader(
                    f"Zalora Stock ({region})  —  SellerStockTemplate*.xlsx",
                    type=["xlsx","xls","csv"], key=f"zsk_{region}"
                )
                region_files[region]["inventory"] = st.file_uploader(
                    f"Inventory ({region})",
                    type=["xlsx","xls","csv"], key=f"inv_{region}",
                    help="PH: Inventory_*  |  MY: PUMA_MY_B2C_Channel_Inventory_*  |  SG: SG_PUMA SG B2C Inventory Rpt_New_*"
                )

    st.markdown("---")
    run_btn = st.button("🚀 Run Listing Audit", type="primary", use_container_width=True)

    if run_btn:
        if not zecom_file or not content_file:
            st.error("Please upload both ZeCom Tracker and Content Master File before running.")
        else:
            with st.spinner("Loading master files…"):
                zecom_df   = load_zecom(zecom_file)
                content_df = load_content(content_file)
            st.success(
                f"Content: {len(content_df):,} SKUs  |  ZeCom Tracker: {len(zecom_df):,} articles"
            )

            all_results = {}
            for region in selected_regions:
                rf = region_files.get(region, {})
                marketplace_dfs = {}

                if rf.get("lazada"):
                    with st.spinner(f"[{region}] Loading Lazada…"):
                        marketplace_dfs["Lazada"] = load_lazada(rf["lazada"])
                if rf.get("shopee"):
                    with st.spinner(f"[{region}] Loading Shopee…"):
                        marketplace_dfs["Shopee"] = load_shopee(rf["shopee"])
                if rf.get("zalora_status") and rf.get("zalora_stock"):
                    with st.spinner(f"[{region}] Loading Zalora…"):
                        marketplace_dfs["Zalora"] = load_zalora(
                            rf["zalora_status"], rf["zalora_stock"]
                        )
                elif rf.get("zalora_status"):
                    st.warning(f"[{region}] Zalora Stock file missing — Zalora skipped.")
                if rf.get("tiktok"):
                    with st.spinner(f"[{region}] Loading TikTok…"):
                        marketplace_dfs["TikTok"] = load_tiktok(rf["tiktok"])

                inv_df = pd.DataFrame()
                if rf.get("inventory"):
                    with st.spinner(f"[{region}] Loading Inventory…"):
                        inv_df = load_inventory(rf["inventory"], region)
                    st.info(f"[{region}] Inventory: {len(inv_df):,} rows")

                if marketplace_dfs:
                    total_skus = sum(len(v) for v in marketplace_dfs.values())
                    st.info(
                        f"[{region}] Marketplace SKUs: {total_skus:,} "
                        f"across {list(marketplace_dfs.keys())}"
                    )
                    with st.spinner(f"Running audit for {region}…"):
                        result_df = run_audit(
                            marketplace_dfs, inv_df, zecom_df, content_df, region
                        )
                    all_results[region] = result_df
                    st.success(f"[{region}] Audit complete — {len(result_df):,} records")
                else:
                    st.warning(f"[{region}] No marketplace files uploaded — skipping.")

            st.session_state.audit_results = all_results
            if all_results:
                st.success("All audits complete! Switch to the Results & Download tab.")

# RESULTS
with tab_results:
    results = st.session_state.audit_results
    if not results:
        st.info("Run the audit first from the Upload Files tab.")
    else:
        combined = pd.concat(results.values(), ignore_index=True)

        st.markdown("### 📊 Audit Overview")
        k1, k2, k3, k4, k5 = st.columns(5)
        total   = len(combined)
        passed  = int((combined["Audit Result"] == "PASS").sum())
        failed  = int((combined["Audit Result"] == "FAIL").sum())
        active  = int((combined["Marketplace Status"] == "ACTIVE").sum())
        instock = int((combined["In Stock"] == "YES").sum())
        k1.markdown(f"<div class='metric-box'><div class='metric-value'>{total:,}</div><div class='metric-label'>Total SKUs</div></div>", unsafe_allow_html=True)
        k2.markdown(f"<div class='metric-box'><div class='metric-value status-active'>{passed:,}</div><div class='metric-label'>✓ Pass</div></div>", unsafe_allow_html=True)
        k3.markdown(f"<div class='metric-box'><div class='metric-value status-inactive'>{failed:,}</div><div class='metric-label'>✗ Fail</div></div>", unsafe_allow_html=True)
        k4.markdown(f"<div class='metric-box'><div class='metric-value'>{active:,}</div><div class='metric-label'>Active on MP</div></div>", unsafe_allow_html=True)
        k5.markdown(f"<div class='metric-box'><div class='metric-value status-missing'>{instock:,}</div><div class='metric-label'>In Stock</div></div>", unsafe_allow_html=True)

        st.markdown("### 📋 Region × Marketplace Breakdown")
        pivot = combined.groupby(["Region","Marketplace"]).agg(
            Total    =("EAN (Seller SKU)", "count"),
            Pass     =("Audit Result",       lambda x: int((x=="PASS").sum())),
            Fail     =("Audit Result",       lambda x: int((x=="FAIL").sum())),
            Active   =("Marketplace Status", lambda x: int((x=="ACTIVE").sum())),
            Inactive =("Marketplace Status", lambda x: int((x=="INACTIVE").sum())),
            In_Stock =("In Stock",           lambda x: int((x=="YES").sum())),
            No_Stock =("In Stock",           lambda x: int((x=="NO").sum())),
        ).reset_index()
        st.dataframe(pivot, use_container_width=True)

        st.markdown("### 🔍 Detailed Results")
        fc1, fc2, fc3 = st.columns(3)
        filt_region = fc1.multiselect("Region",
            options=combined["Region"].unique().tolist(),
            default=combined["Region"].unique().tolist())
        filt_mp = fc2.multiselect("Marketplace",
            options=combined["Marketplace"].unique().tolist(),
            default=combined["Marketplace"].unique().tolist())
        filt_result = fc3.multiselect("Audit Result",
            options=["PASS","FAIL"], default=["PASS","FAIL"])

        filtered = combined[
            combined["Region"].isin(filt_region) &
            combined["Marketplace"].isin(filt_mp) &
            combined["Audit Result"].isin(filt_result)
        ]
        st.dataframe(filtered, use_container_width=True, height=420)

        st.markdown("### ⚠️ Top Failure Reasons")
        fails = combined[combined["Audit Result"] == "FAIL"]
        if not fails.empty:
            reason_counts = (
                fails["Reasons"].str.split(" | ").explode().str.strip()
                .value_counts().head(10).reset_index()
            )
            reason_counts.columns = ["Reason", "Count"]
            st.dataframe(reason_counts, use_container_width=True)
        else:
            st.success("No failures found!")

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
            type="primary",
        )
        st.caption("Report includes: Summary, per-region detail sheets, and Mismatches tab.")
