"""
PUMA Listing Audit Analyzer  v3.0
==================================
SOURCE OF TRUTH ORDER:
  1. ZeCom Tracker   → master list of Article_No + tracker YES/NO/OFF per marketplace
  2. Special Override → uploaded sheet that overrides tracker status (ACTIVE / INACTIVE)
  3. Content File     → Article_No (parent) → EAN variants (children)
  4. Inventory File   → EAN → stock per channel
  5. Marketplace File → what is actually listed (EAN level)

OUTPUT PER MARKETPLACE (article-level driven by ZeCom):
  • Active           – tracker ACTIVE, listed on MP as active, in stock
  • Inactive         – tracker says INACTIVE or no stock or future launch
  • Not Listed - Article Level  – should be active, but ZERO variants on MP at all
  • Not Listed - Variant Level  – should be active, some variants on MP but specific EAN missing

Special Override File:
  Columns: Article_No  |  Status (ACTIVE / INACTIVE)
  Applies to ALL regions.  Overrides ZeCom tracker for the listed articles.
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
.main-header{background:linear-gradient(135deg,#1a1a2e 0%,#16213e 50%,#0f3460 100%);
    padding:2rem;border-radius:12px;text-align:center;margin-bottom:1.5rem;color:white;}
.main-header h1{font-size:2rem;font-weight:700;margin:0;}
.main-header p{color:#a0aec0;margin-top:.4rem;font-size:.9rem;}
.metric-box{background:white;border-radius:8px;padding:.9rem;
    border:1px solid #e2e8f0;text-align:center;margin-bottom:.5rem;}
.metric-value{font-size:1.7rem;font-weight:800;}
.metric-label{font-size:.75rem;color:#718096;margin-top:.15rem;}
.cg{color:#38a169;} .cr{color:#e53e3e;} .co{color:#d69e2e;} .cp{color:#805ad5;}
.inv-debug{background:#fffbea;border:1px solid #f6e05e;border-radius:6px;
    padding:.6rem .9rem;font-size:.8rem;margin-top:.5rem;}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class='main-header'>
  <h1>📊 PUMA Listing Audit Analyzer</h1>
  <p>ZeCom-driven · Article &amp; EAN Level · MY / SG / PH · Lazada / Shopee / Zalora / TikTok</p>
</div>""", unsafe_allow_html=True)

# ── Constants ─────────────────────────────────────────────────────────────────
REGIONS      = ["MY","SG","PH"]
MARKETPLACES = ["Lazada","Shopee","Zalora","TikTok"]
CHANNEL_BUFFER_PH = {"Lazada":1,"Shopee":0,"Zalora":0,"TikTok":0}

# ══════════════════════════════════════════════════════════════════════════════
# HELPERS — bullet-proof column handling
# ══════════════════════════════════════════════════════════════════════════════
def _sc(c)->str:
    try:
        s=str(c).strip()
        return "" if s in ("nan","None","") else s
    except:
        return ""

def _cl(c)->str:
    return _sc(c).lower()

def sanitize_cols(df:pd.DataFrame)->pd.DataFrame:
    df.columns=[_sc(c) or f"_col_{i}" for i,c in enumerate(df.columns)]
    return df

def find_header_row(file,max_scan:int=20)->int:
    try:
        prev=pd.read_excel(file,header=None,nrows=max_scan)
        file.seek(0)
        best_r,best_s=0,-1
        for i,row in prev.iterrows():
            sc=int(row.dropna().apply(lambda v:isinstance(v,str) and len(v.strip())>0).sum())
            if sc>best_s:
                best_s,best_r=sc,int(i)
        return best_r
    except:
        try: file.seek(0)
        except: pass
        return 0

def safe_read(file)->pd.DataFrame:
    try:
        hdr=find_header_row(file)
        df=pd.read_excel(file,header=hdr)
    except:
        try:
            file.seek(0); df=pd.read_excel(file,header=0)
        except:
            try:
                file.seek(0); df=pd.read_csv(file,header=0)
            except Exception as e:
                st.warning(f"Cannot read file: {e}"); return pd.DataFrame()
    df=sanitize_cols(df)
    return df.dropna(how="all").reset_index(drop=True)

def norm_col(df:pd.DataFrame,candidates:list,new_name:str)->pd.DataFrame:
    cands=[str(x).strip().lower() for x in candidates]
    for col in df.columns:
        if _cl(col) in cands:
            if col!=new_name:
                df=df.rename(columns={col:new_name})
            break
    return df

def cs(s:pd.Series)->pd.Series:
    """Clean string series → upper stripped, blanks as empty string."""
    return s.astype(str).str.strip().str.upper().replace({"NAN":"","NONE":"","NAT":""})

def to_num(series:pd.Series)->pd.Series:
    return pd.to_numeric(series,errors="coerce").fillna(0).clip(lower=0).astype(int)

# ══════════════════════════════════════════════════════════════════════════════
# FILE LOADERS
# ══════════════════════════════════════════════════════════════════════════════

# ── 1. ZeCom Tracker ──────────────────────────────────────────────────────────
def load_zecom(file)->pd.DataFrame:
    """
    Returns one row per Article_No with columns:
      Article_No, Tracker_Lazada, Tracker_Shopee, Tracker_Zalora,
      Tracker_TikTok, Launch_Date
    Tracker values: YES → active, NO/OFF → inactive, blank/NaN → unknown
    """
    df=safe_read(file)

    df=norm_col(df,[
        "PIM Article#","PIM Article","Article No","ArticleNo",
        "Color_No","ColorNo","Color No","Article Number",
        "Style Number","StyleNo","Style No","Parent SKU"
    ],"Article_No")

    if "Article_No" not in df.columns:
        st.error("ZeCom Tracker: cannot find Article No column. "
                 "Checked: PIM Article#, Color_No, Article No, Style Number")
        return pd.DataFrame(columns=["Article_No"])

    # Tracker columns — scan every column for marketplace names
    for mp in MARKETPLACES:
        # look for a column whose name contains the mp name (case-insensitive)
        # but exclude Article_No itself and already-renamed tracker cols
        existing=[f"Tracker_{m}" for m in MARKETPLACES]
        cols=[c for c in df.columns
              if mp.lower() in _cl(c) and c not in existing and c!="Article_No"]
        if cols:
            df=df.rename(columns={cols[0]:f"Tracker_{mp}"})

    # Guarantee all tracker columns exist
    for mp in MARKETPLACES:
        if f"Tracker_{mp}" not in df.columns:
            df[f"Tracker_{mp}"]=np.nan

    # Launch date
    skip={"Article_No"}|{f"Tracker_{m}" for m in MARKETPLACES}
    date_cols=[c for c in df.columns
               if ("launch" in _cl(c) or "go live" in _cl(c) or "golive" in _cl(c))
               and c not in skip]
    if date_cols:
        df=df.rename(columns={date_cols[0]:"Launch_Date"})
        df["Launch_Date"]=pd.to_datetime(df["Launch_Date"],errors="coerce")
    else:
        df["Launch_Date"]=pd.NaT

    df["Article_No"]=cs(df["Article_No"])
    df=df[df["Article_No"].str.len()>0].drop_duplicates(subset=["Article_No"])
    return df.reset_index(drop=True)


# ── 2. Special Override ────────────────────────────────────────────────────────
def load_special_override(file)->pd.DataFrame:
    """
    Columns: Article_No | Status (ACTIVE / INACTIVE)
    Overrides ZeCom for all regions.
    """
    df=safe_read(file)
    df=norm_col(df,[
        "PIM Article#","PIM Article","Article No","ArticleNo",
        "Color_No","ColorNo","Article Number","Style Number"
    ],"Article_No")
    df=norm_col(df,["Status","status","Active","Listing Status","Override"],"Status")

    if "Article_No" not in df.columns or "Status" not in df.columns:
        st.warning("Special Override: needs Article_No and Status columns.")
        return pd.DataFrame(columns=["Article_No","Status"])

    df["Article_No"]=cs(df["Article_No"])
    df["Status"]=cs(df["Status"])
    df=df[df["Article_No"].str.len()>0]
    return df[["Article_No","Status"]].drop_duplicates(subset=["Article_No"]).reset_index(drop=True)


# ── 3. Content File ────────────────────────────────────────────────────────────
def load_content(file)->pd.DataFrame:
    """
    Returns one row per EAN with at minimum: EAN, Article_No.
    Also captures size / color / description if present.
    """
    df=safe_read(file)

    df=norm_col(df,[
        "EAN","ean","Barcode","barcode","Child SKU","ChildSKU",
        "Seller SKU","SellerSKU","Item EAN","GTIN","UPC"
    ],"EAN")

    df=norm_col(df,[
        "Color_No","ColorNo","Article No","ArticleNo",
        "PIM Article#","Article Number","Parent SKU",
        "Color No","Style Number","StyleNo"
    ],"Article_No")

    # Optional enrichment columns
    extra=[]
    for label,keys in [("Size",["size","sz"]),
                       ("Color",["colour","color"]),
                       ("Description",["desc","name","product name"])]:
        cols=[c for c in df.columns
              if any(k in _cl(c) for k in keys)
              and c not in ("EAN","Article_No")]
        if cols:
            df=df.rename(columns={cols[0]:label})
            extra.append(label)

    for col in ["EAN","Article_No"]:
        if col not in df.columns:
            df[col]=np.nan

    df["EAN"]=cs(df["EAN"])
    df["Article_No"]=cs(df["Article_No"])
    df=df[(df["EAN"].str.len()>0)&(df["Article_No"].str.len()>0)]

    keep=["EAN","Article_No"]+extra
    keep=[c for c in keep if c in df.columns]
    return df[keep].drop_duplicates(subset=["EAN"]).reset_index(drop=True)


# ── 4. Inventory File ─────────────────────────────────────────────────────────
def load_inventory(file,region:str)->tuple[pd.DataFrame,dict]:
    """
    Returns (DataFrame, debug_info).
    DataFrame columns: EAN, Inv_Lazada, Inv_Shopee, Inv_Zalora, Inv_TikTok
    Applies PH Lazada buffer.
    debug_info: dict with detected column names for each channel.
    """
    df=safe_read(file)

    # Show user the columns found (helps debug mismatches)
    all_cols=list(df.columns)

    # EAN
    df=norm_col(df,[
        "EAN","ean","Barcode","barcode","Item EAN","GTIN",
        "SKU","sku","Item Code","Seller SKU","SellerSKU",
        "Material","Material Number"
    ],"EAN")

    if "EAN" not in df.columns:
        st.warning(f"[{region}] Inventory: cannot find EAN/Barcode column.\n"
                   f"Columns found: {all_cols[:20]}")
        return pd.DataFrame(columns=["EAN","Inv_Lazada","Inv_Shopee","Inv_Zalora","Inv_TikTok"]),{}

    df["EAN"]=cs(df["EAN"])
    df=df[df["EAN"].str.len()>0]

    debug={}

    def find_first(keywords:list)->str|None:
        """Return first column name that contains any keyword."""
        for kw in keywords:
            for c in df.columns:
                if kw in _cl(c) and c!="EAN":
                    return c
        return None

    # Try channel-specific columns first
    laz_c=find_first(["lazada"])
    sho_c=find_first(["shopee"])
    zal_c=find_first(["zalora"])
    ttk_c=find_first(["tiktok","tik tok"])

    # Fallback: single total / available / SOH column
    tot_c=find_first([
        "available","on hand","onhand","total","qty","quantity",
        "stock","soh","free","unrestricted"
    ])

    debug={
        "Lazada col" : laz_c or f"not found → using '{tot_c}'",
        "Shopee col" : sho_c or f"not found → using '{tot_c}'",
        "Zalora col" : zal_c or f"not found → using '{tot_c}'",
        "TikTok col" : ttk_c or f"not found → using '{tot_c}'",
        "Total col"  : tot_c or "not found",
        "All columns": ", ".join(all_cols[:30])
    }

    result=pd.DataFrame({"EAN":df["EAN"]})
    result["Inv_Lazada"]=to_num(df[laz_c] if laz_c else (df[tot_c] if tot_c else pd.Series(0,index=df.index)))
    result["Inv_Shopee"]=to_num(df[sho_c] if sho_c else (df[tot_c] if tot_c else pd.Series(0,index=df.index)))
    result["Inv_Zalora"]=to_num(df[zal_c] if zal_c else (df[tot_c] if tot_c else pd.Series(0,index=df.index)))
    result["Inv_TikTok"]=to_num(df[ttk_c] if ttk_c else (df[tot_c] if tot_c else pd.Series(0,index=df.index)))

    if region=="PH":
        result["Inv_Lazada"]=(result["Inv_Lazada"]-CHANNEL_BUFFER_PH["Lazada"]).clip(lower=0)

    result=result.drop_duplicates(subset=["EAN"]).reset_index(drop=True)
    return result, debug


# ── 5. Marketplace Loaders ────────────────────────────────────────────────────
def _mp_load(df,ean_cands,status_cands,stock_cands,mp_name)->pd.DataFrame:
    df=norm_col(df,ean_cands,"EAN")
    df=norm_col(df,status_cands,"MP_Status")
    df=norm_col(df,stock_cands,"MP_Stock")
    for col in ["EAN","MP_Status","MP_Stock"]:
        if col not in df.columns: df[col]=np.nan
    df["EAN"]=cs(df["EAN"])
    df["MP_Stock"]=pd.to_numeric(df["MP_Stock"],errors="coerce").fillna(0)
    df["Marketplace"]=mp_name
    df=df[df["EAN"].str.len()>0]
    return df[["EAN","MP_Status","MP_Stock","Marketplace"]].drop_duplicates(subset=["EAN"])

def load_lazada(file)->pd.DataFrame:
    df=safe_read(file)
    return _mp_load(df,
        ["SellerSKU","Seller SKU","Seller Sku","seller_sku","SKU","sku","ItemId","item_id"],
        ["Status","status","ItemStatus","Item Status","Listing Status","Active","active"],
        ["Available","Stock","Quantity","available","quantity","Qty","FreeQty","sellable_quantity","free qty"],
        "Lazada")

def load_shopee(file)->pd.DataFrame:
    df=safe_read(file)
    return _mp_load(df,
        ["SKU","sku","Seller SKU","SellerSKU","Seller Sku","variation_sku","Item SKU","Model SKU"],
        ["Status","status","Listing Status","Product Status","Item Status","Publish"],
        ["Stock","Quantity","Available Stock","stock","quantity","Qty","Current Stock","Available"],
        "Shopee")

def load_zalora(sf,stf)->pd.DataFrame:
    ds=safe_read(sf);  dstk=safe_read(stf)
    ds  =norm_col(ds,  ["SellerSku","Seller Sku","SellerSKU","Seller SKU","seller_sku","SKU"],"EAN")
    ds  =norm_col(ds,  ["Status","status","ItemStatus","Item Status","Active"],"MP_Status")
    dstk=norm_col(dstk,["SellerSku","Seller Sku","SellerSKU","Seller SKU","seller_sku","SKU"],"EAN")
    dstk=norm_col(dstk,["Stock","Quantity","Available","Qty","quantity"],"MP_Stock")
    for col in ["EAN","MP_Status"]:
        if col not in ds.columns: ds[col]=np.nan
    if "EAN"      not in dstk.columns: dstk["EAN"]=np.nan
    if "MP_Stock" not in dstk.columns: dstk["MP_Stock"]=np.nan
    ds["EAN"]  =cs(ds["EAN"])
    dstk["EAN"]=cs(dstk["EAN"])
    merged=ds.merge(dstk[["EAN","MP_Stock"]].drop_duplicates("EAN"),on="EAN",how="left")
    merged["MP_Stock"]=pd.to_numeric(merged["MP_Stock"],errors="coerce").fillna(0)
    merged["Marketplace"]="Zalora"
    merged=merged[merged["EAN"].str.len()>0]
    return merged[["EAN","MP_Status","MP_Stock","Marketplace"]].drop_duplicates(subset=["EAN"])

def load_tiktok(file)->pd.DataFrame:
    df=safe_read(file)
    return _mp_load(df,
        ["Seller sku","Seller SKU","SellerSKU","Seller Sku","SKU","sku","Variation SKU"],
        ["Status","status","Product Status","Item Status","Active"],
        ["Stock","Quantity","Available","Qty","quantity","Available Stock"],
        "TikTok")


# ══════════════════════════════════════════════════════════════════════════════
# STATUS LOGIC
# ══════════════════════════════════════════════════════════════════════════════
def resolve_tracker(val)->str:
    """Convert tracker cell → 'ACTIVE' | 'INACTIVE' | 'UNKNOWN'"""
    if pd.isna(val): return "UNKNOWN"
    v=str(val).strip().upper()
    if v=="YES":              return "ACTIVE"
    if v in ("NO","OFF"):     return "INACTIVE"
    return "UNKNOWN"

def is_mp_active(val)->bool:
    if pd.isna(val): return False
    return str(val).strip().upper() in [
        "ACTIVE","1","TRUE","YES","ENABLED","PUBLISHED","LISTED","NORMAL","ACTIVATED","FOR SALE"
    ]

def expected_status(zecom_row:pd.Series, mp:str, override_map:dict)->tuple[str,list]:
    """
    Returns (expected: 'ACTIVE'|'INACTIVE', reasons: list[str])
    Priority: Special Override > ZeCom Tracker > Launch Date
    """
    art=str(zecom_row.get("Article_No","")).strip()
    reasons=[]

    # 1. Special override takes priority
    if art in override_map:
        ovr=override_map[art]
        reasons.append(f"Special Override: {ovr}")
        return ovr, reasons

    # 2. Tracker status
    t_col=f"Tracker_{mp}"
    t_val=zecom_row.get(t_col,np.nan)
    t_status=resolve_tracker(t_val)

    if t_status=="INACTIVE":
        reasons.append(f"Tracker Status is NO/OFF for {mp}")
        return "INACTIVE", reasons

    # 3. Launch date check
    launch=zecom_row.get("Launch_Date",pd.NaT)
    today=pd.Timestamp.today()
    if pd.notna(launch):
        try:
            if pd.Timestamp(launch)>today:
                reasons.append(f"Future Launch Date: {pd.Timestamp(launch).date()}")
                return "INACTIVE", reasons
        except: pass

    if t_status=="ACTIVE":
        return "ACTIVE", reasons

    reasons.append(f"Tracker blank/unknown for {mp}")
    return "INACTIVE", reasons


# ══════════════════════════════════════════════════════════════════════════════
# AUDIT ENGINE  — ZeCom article-driven
# ══════════════════════════════════════════════════════════════════════════════
def run_audit(
    mp_dfs:dict,           # {mp_name: DataFrame with EAN/MP_Status/MP_Stock}
    inv_df:pd.DataFrame,   # EAN-level inventory
    zecom_df:pd.DataFrame, # Article_No-level tracker
    content_df:pd.DataFrame,  # EAN → Article_No mapping
    override_map:dict,     # Article_No → 'ACTIVE'/'INACTIVE'
    region:str
)->dict:
    """
    Returns dict with keys per marketplace:
      {mp: {"Active":df, "Inactive":df, "Not Listed - Article Level":df, "Not Listed - Variant Level":df}}
    DRIVEN BY ZECOM: we iterate every Article_No in ZeCom, not by what's on the marketplace.
    """
    today=pd.Timestamp.today()

    # ── Build lookup maps ────────────────────────────────────────────────────
    # Article_No → list of EANs  (from content)
    art_to_eans: dict[str,list] = {}
    for _,row in content_df.iterrows():
        art=str(row["Article_No"]).strip()
        ean=str(row["EAN"]).strip()
        if art and ean:
            art_to_eans.setdefault(art,[]).append(ean)

    # EAN → content extra info (size, color, description)
    ean_meta: dict[str,dict] = {}
    extra_cols=[c for c in content_df.columns if c not in ("EAN","Article_No")]
    for _,row in content_df.iterrows():
        ean=str(row["EAN"]).strip()
        if ean:
            ean_meta[ean]={c:row[c] for c in extra_cols if c in row.index}

    # Inventory index: EAN → row
    inv_idx={}
    if not inv_df.empty:
        for _,row in inv_df.iterrows():
            inv_idx[str(row["EAN"]).strip()]=row

    inv_col={"Lazada":"Inv_Lazada","Shopee":"Inv_Shopee",
              "Zalora":"Inv_Zalora","TikTok":"Inv_TikTok"}

    def get_inv(ean:str,mp:str)->int:
        row=inv_idx.get(ean)
        if row is None: return 0
        col=inv_col.get(mp,"Inv_Lazada")
        v=row.get(col,0)
        r=pd.to_numeric(v,errors="coerce")
        return int(r) if pd.notna(r) else 0

    # MP lookup: EAN → MP row
    mp_ean_idx: dict[str,dict] = {}  # will be rebuilt per mp below

    results={}

    for mp in MARKETPLACES:
        mp_df=mp_dfs.get(mp,pd.DataFrame())
        # Build EAN → mp row index
        mp_ean_idx={}
        if not mp_df.empty:
            for _,row in mp_df.iterrows():
                e=str(row["EAN"]).strip()
                if e: mp_ean_idx[e]=row

        active_rows=[]
        inactive_rows=[]
        not_listed_article_rows=[]
        not_listed_ean_rows=[]

        # ── Iterate EVERY article in ZeCom ───────────────────────────────────
        for _,z_row in zecom_df.iterrows():
            art=str(z_row["Article_No"]).strip()
            if not art: continue

            exp_status, exp_reasons = expected_status(z_row, mp, override_map)

            variants=art_to_eans.get(art,[])          # all EANs for this article
            listed_eans=[e for e in variants if e in mp_ean_idx]  # EANs on MP

            # ── Article has NO variants in content ─────────────────────────
            # Still report at article level
            if not variants:
                base={
                    "Region":region,"Marketplace":mp,
                    "Article No":art,"Expected Status":exp_status,
                    "Total Variants":0,"Listed Variants":0,
                    "Reason":"; ".join(exp_reasons) if exp_reasons else "-",
                    "Note":"Article not in Content File"
                }
                if exp_status=="INACTIVE":
                    inactive_rows.append(base)
                else:
                    not_listed_article_rows.append({**base,"Note":"No variants in Content File"})
                continue

            # ── INACTIVE articles ──────────────────────────────────────────
            if exp_status=="INACTIVE":
                # Report each variant with its stock and MP status
                for ean in variants:
                    mp_row=mp_ean_idx.get(ean)
                    inv_stock=get_inv(ean,mp)
                    meta=ean_meta.get(ean,{})
                    row_data={
                        "Region":region,"Marketplace":mp,
                        "Article No":art,"EAN (Seller SKU)":ean,
                        **{k:v for k,v in meta.items()},
                        "Expected Status":"INACTIVE",
                        "MP Listed":"YES" if mp_row is not None else "NO",
                        "MP Status":str(mp_row["MP_Status"]) if mp_row is not None else "-",
                        "MP Stock":int(mp_row["MP_Stock"]) if mp_row is not None else 0,
                        "Inventory Stock":inv_stock,
                        "Reason":"; ".join(exp_reasons) if exp_reasons else "-",
                    }
                    inactive_rows.append(row_data)
                continue

            # ── ACTIVE articles ────────────────────────────────────────────
            # Check stock across variants
            total_variants=len(variants)
            listed_count=len(listed_eans)

            if listed_count==0:
                # NOT LISTED at article level — no variants on MP at all
                # Still show all variants
                for ean in variants:
                    inv_stock=get_inv(ean,mp)
                    meta=ean_meta.get(ean,{})
                    not_listed_article_rows.append({
                        "Region":region,"Marketplace":mp,
                        "Article No":art,"EAN (Seller SKU)":ean,
                        **{k:v for k,v in meta.items()},
                        "Expected Status":"ACTIVE",
                        "MP Listed":"NO","MP Status":"-","MP Stock":0,
                        "Inventory Stock":inv_stock,
                        "Reason":"Article should be ACTIVE but has ZERO variants listed on MP",
                    })
            else:
                # Some or all variants listed — process each EAN
                for ean in variants:
                    mp_row=mp_ean_idx.get(ean)
                    inv_stock=get_inv(ean,mp)
                    meta=ean_meta.get(ean,{})

                    if mp_row is not None:
                        # EAN IS on marketplace
                        mp_active=is_mp_active(mp_row["MP_Status"])
                        row_data={
                            "Region":region,"Marketplace":mp,
                            "Article No":art,"EAN (Seller SKU)":ean,
                            **{k:v for k,v in meta.items()},
                            "Expected Status":"ACTIVE",
                            "MP Listed":"YES",
                            "MP Status":str(mp_row["MP_Status"]),
                            "MP Stock":int(mp_row["MP_Stock"]),
                            "Inventory Stock":inv_stock,
                            "Reason":"-",
                        }
                        if mp_active and inv_stock>0:
                            row_data["Reason"]="-"
                            active_rows.append(row_data)
                        elif mp_active and inv_stock==0:
                            row_data["Reason"]="Listed ACTIVE but NO inventory stock"
                            active_rows.append(row_data)   # still active on MP
                        else:
                            row_data["Reason"]="Listed INACTIVE on MP despite tracker ACTIVE"
                            inactive_rows.append(row_data)
                    else:
                        # EAN is NOT on marketplace — variant missing
                        not_listed_ean_rows.append({
                            "Region":region,"Marketplace":mp,
                            "Article No":art,"EAN (Seller SKU)":ean,
                            **{k:v for k,v in meta.items()},
                            "Expected Status":"ACTIVE",
                            "MP Listed":"NO","MP Status":"-","MP Stock":0,
                            "Inventory Stock":inv_stock,
                            "Reason":f"Variant (size/EAN) missing on MP — "
                                     f"{listed_count}/{total_variants} variants listed",
                        })

        results[mp]={
            "Active"                    : pd.DataFrame(active_rows),
            "Inactive"                  : pd.DataFrame(inactive_rows),
            "Not Listed - Article Level": pd.DataFrame(not_listed_article_rows),
            "Not Listed - Variant Level": pd.DataFrame(not_listed_ean_rows),
        }

    return results


# ══════════════════════════════════════════════════════════════════════════════
# EXCEL EXPORT
# ══════════════════════════════════════════════════════════════════════════════
CAT_COLORS={
    "Active"                    :{"hdr":"#1e6f39","tab":"#c6efce","txt":"#276221"},
    "Inactive"                  :{"hdr":"#9c0006","tab":"#ffc7ce","txt":"#9c0006"},
    "Not Listed - Article Level":{"hdr":"#7b5800","tab":"#ffeb9c","txt":"#7b5800"},
    "Not Listed - Variant Level":{"hdr":"#4a235a","tab":"#e8d5f5","txt":"#4a235a"},
}

def make_excel(all_results:dict)->bytes:
    """
    Sheets:
      📋 Summary
      For each marketplace × category: e.g. "Lazada - Active", "Lazada - Inactive" etc.
    """
    output=BytesIO()
    with pd.ExcelWriter(output,engine="xlsxwriter") as writer:
        wb=writer.book
        ttl_fmt=wb.add_format({"bold":True,"font_size":13,"font_color":"#0f3460","font_name":"Arial"})
        sub_fmt=wb.add_format({"font_size":9,"italic":True,"font_color":"#4a5568","font_name":"Arial"})
        norm_fmt=wb.add_format({"font_name":"Arial","font_size":9,"border":1})

        def hdr_fmt(bg,fc="#ffffff"):
            return wb.add_format({"bold":True,"bg_color":bg,"font_color":fc,
                                  "border":1,"align":"center","valign":"vcenter",
                                  "font_name":"Arial","font_size":9,"text_wrap":True})
        def data_fmt(bg,fc="#000000"):
            return wb.add_format({"bg_color":bg,"font_color":fc,"border":1,
                                  "font_name":"Arial","font_size":9})

        # ── Summary sheet ──────────────────────────────────────────────────
        ws=wb.add_worksheet("Summary")
        writer.sheets["Summary"]=ws
        ws.set_column("A:A",18); ws.set_column("B:B",25)
        ws.set_column("C:H",14)
        ws.write("A1","PUMA Listing Audit — Summary",ttl_fmt)
        ws.write("A2",f"Generated: {datetime.now().strftime('%d %b %Y  %H:%M')}",sub_fmt)

        cats=["Active","Inactive","Not Listed - Article Level","Not Listed - Variant Level"]
        r=4
        # Header
        ws.write(r,0,"Region",hdr_fmt("#0f3460"))
        ws.write(r,1,"Marketplace",hdr_fmt("#0f3460"))
        for ci,cat in enumerate(cats):
            cc=CAT_COLORS[cat]
            ws.write(r,ci+2,cat,hdr_fmt(cc["hdr"]))
        ws.set_row(r,30)
        r+=1

        for region, mp_results in all_results.items():
            for mp in MARKETPLACES:
                if mp not in mp_results: continue
                cats_data=mp_results[mp]
                ws.write(r,0,region,norm_fmt)
                ws.write(r,1,mp,norm_fmt)
                for ci,cat in enumerate(cats):
                    cnt=len(cats_data.get(cat,pd.DataFrame()))
                    cc=CAT_COLORS[cat]
                    fmt=data_fmt(cc["tab"],cc["txt"])
                    ws.write(r,ci+2,cnt,fmt)
                r+=1

        # ── Per-MP × Category sheets ───────────────────────────────────────
        for region, mp_results in all_results.items():
            for mp in MARKETPLACES:
                if mp not in mp_results: continue
                for cat in cats:
                    df=mp_results[mp].get(cat,pd.DataFrame())
                    cc=CAT_COLORS[cat]
                    # Sheet name max 31 chars
                    sname=f"{mp[:4]} {region} {cat[:18]}"
                    sname=sname[:31]

                    if df.empty:
                        ws2=wb.add_worksheet(sname)
                        writer.sheets[sname]=ws2
                        ws2.write(0,0,f"{mp} [{region}] — {cat}",ttl_fmt)
                        ws2.write(1,0,"No records in this category.",sub_fmt)
                        continue

                    df.to_excel(writer,sheet_name=sname,index=False,startrow=2)
                    ws2=writer.sheets[sname]
                    ws2.write(0,0,f"{mp} [{region}] — {cat}",ttl_fmt)
                    ws2.write(1,0,f"Total records: {len(df)}",sub_fmt)

                    hf=hdr_fmt(cc["hdr"])
                    rf=data_fmt(cc["tab"],cc["txt"])
                    nf=norm_fmt

                    for ci,col in enumerate(df.columns):
                        ws2.write(2,ci,col,hf)
                        ws2.set_column(ci,ci,20)

                    for ri,(_,rec) in enumerate(df.iterrows()):
                        for ci,col in enumerate(df.columns):
                            val=rec[col]
                            safe=("" if (isinstance(val,float) and np.isnan(val))
                                  else str(val))
                            if col in ("Expected Status","Reason"):
                                ws2.write(ri+3,ci,safe,rf)
                            else:
                                ws2.write(ri+3,ci,safe,nf)

    return output.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════════════════════════
if "audit_results"  not in st.session_state: st.session_state.audit_results={}
if "inv_debug"      not in st.session_state: st.session_state.inv_debug={}

# ══════════════════════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════════════════════
tab_upload,tab_results,tab_debug,tab_help = st.tabs([
    "📁 Upload & Run","📊 Results & Download","🔧 Inventory Debug","❓ Help"
])

# ── HELP ─────────────────────────────────────────────────────────────────────
with tab_help:
    st.markdown("""
## 📖 File Naming Reference

| File | Expected Name Pattern |
|---|---|
| Lazada | starts with `pricestock` |
| Shopee | contains `Shopee` + `Masterfile` |
| Zalora Status | starts with `SellerStatusTemplate` |
| Zalora Stock | starts with `SellerStockTemplate` |
| TikTok | starts with `TikTokSellerCenter` or contains `batchedit` |
| Inventory PH | starts with `Inventory_` |
| Inventory MY | starts with `PUMA_MY_B2C_Channel_Inventory_` |
| Inventory SG | starts with `SG_PUMA SG B2C Inventory Rpt_New_` |
| ZeCom Tracker | contains `zecom` or `tracker` |
| Content Master | master dump: EAN ↔ Article No |
| Special Override | any name — columns: Article_No, Status |

---
## ✅ Active / Inactive Decision Logic

```
Priority 1: Special Override file (ACTIVE / INACTIVE) — overrides everything
Priority 2: ZeCom Tracker
   YES  → ACTIVE  (if also in stock and past launch date)
   NO   → INACTIVE
   OFF  → INACTIVE
Priority 3: Launch Date → future = INACTIVE
Priority 4: Inventory stock → no stock = INACTIVE
```

## 📋 Output Categories (per Marketplace per Region)

| Category | Meaning |
|---|---|
| **Active** | Tracker ACTIVE, listed on MP, in stock |
| **Inactive** | Tracker says INACTIVE, or no stock, or future launch |
| **Not Listed - Article Level** | Tracker ACTIVE, but ZERO variants listed on this MP |
| **Not Listed - Variant Level** | Tracker ACTIVE, article has some listings, but specific EAN/size missing |

## 🔑 Key Identifiers
- **EAN** = Seller SKU = child/variant level (size)
- **Article No / Color No / PIM Article#** = parent level (style)
- PH Lazada inventory buffer: **stock − 1** (floor 0)
    """)

# ── UPLOAD & RUN ─────────────────────────────────────────────────────────────
with tab_upload:
    st.markdown("### Step 1 — Select Region(s)")
    selected_regions=st.multiselect("Regions:",REGIONS,default=REGIONS)

    st.markdown("### Step 2 — Master Files (apply to all regions)")
    c1,c2,c3=st.columns(3)
    with c1:
        zecom_file=st.file_uploader("📋 ZeCom Tracker **(required)**",
            type=["xlsx","xls","csv"],key="zecom",
            help="Article No + Tracker YES/NO/OFF per marketplace + Launch Date")
    with c2:
        content_file=st.file_uploader("📦 Content Master **(required)**",
            type=["xlsx","xls","csv"],key="content",
            help="EAN (child variant) ↔ Article No (parent style)")
    with c3:
        override_file=st.file_uploader("⚡ Special Article Override *(optional)*",
            type=["xlsx","xls","csv"],key="override",
            help="Columns: Article_No | Status (ACTIVE/INACTIVE) — overrides tracker for all regions")

    st.markdown("### Step 3 — Region Files")
    region_files={}
    for region in selected_regions:
        with st.expander(f"📂 {region}",expanded=True):
            region_files[region]={}
            c1,c2=st.columns(2)
            with c1:
                region_files[region]["lazada"]=st.file_uploader(
                    f"Lazada — pricestock*.xlsx",type=["xlsx","xls","csv"],key=f"laz_{region}")
                region_files[region]["shopee"]=st.file_uploader(
                    f"Shopee — Shopee*Masterfile*.xlsx",type=["xlsx","xls","csv"],key=f"sho_{region}")
                region_files[region]["tiktok"]=st.file_uploader(
                    f"TikTok — TikTokSellerCenter*.xlsx",type=["xlsx","xls","csv"],key=f"ttk_{region}")
            with c2:
                region_files[region]["zalora_status"]=st.file_uploader(
                    f"Zalora Status — SellerStatusTemplate*.xlsx",type=["xlsx","xls","csv"],key=f"zst_{region}")
                region_files[region]["zalora_stock"]=st.file_uploader(
                    f"Zalora Stock — SellerStockTemplate*.xlsx",type=["xlsx","xls","csv"],key=f"zsk_{region}")
                region_files[region]["inventory"]=st.file_uploader(
                    f"Inventory ({region})",type=["xlsx","xls","csv"],key=f"inv_{region}",
                    help="PH: Inventory_*  |  MY: PUMA_MY_B2C_Channel_Inventory_*  |  SG: SG_PUMA SG B2C Inventory Rpt_New_*")

    st.markdown("---")
    run_btn=st.button("🚀 Run Listing Audit",type="primary",use_container_width=True)

    if run_btn:
        errors=[]
        if not zecom_file:  errors.append("ZeCom Tracker file is required.")
        if not content_file: errors.append("Content Master file is required.")
        if errors:
            for e in errors: st.error(e)
        else:
            # Load master files
            with st.spinner("Loading ZeCom Tracker…"):
                zecom_df=load_zecom(zecom_file)
            st.success(f"ZeCom Tracker: **{len(zecom_df):,}** articles loaded")

            with st.spinner("Loading Content Master…"):
                content_df=load_content(content_file)
            st.success(f"Content Master: **{len(content_df):,}** EANs "
                       f"/ **{content_df['Article_No'].nunique():,}** articles")

            # Override map
            override_map={}
            if override_file:
                with st.spinner("Loading Special Override…"):
                    ov_df=load_special_override(override_file)
                override_map={r["Article_No"]:r["Status"]
                              for _,r in ov_df.iterrows()
                              if r["Article_No"] and r["Status"] in ("ACTIVE","INACTIVE")}
                st.info(f"Special Override: **{len(override_map):,}** articles overridden")

            all_results={}
            inv_debug_all={}

            for region in selected_regions:
                rf=region_files.get(region,{})

                # Load marketplace files
                mp_dfs={}
                if rf.get("lazada"):
                    with st.spinner(f"[{region}] Lazada…"):
                        mp_dfs["Lazada"]=load_lazada(rf["lazada"])
                    st.info(f"[{region}] Lazada: {len(mp_dfs['Lazada']):,} EANs")
                if rf.get("shopee"):
                    with st.spinner(f"[{region}] Shopee…"):
                        mp_dfs["Shopee"]=load_shopee(rf["shopee"])
                    st.info(f"[{region}] Shopee: {len(mp_dfs['Shopee']):,} EANs")
                if rf.get("zalora_status") and rf.get("zalora_stock"):
                    with st.spinner(f"[{region}] Zalora…"):
                        mp_dfs["Zalora"]=load_zalora(rf["zalora_status"],rf["zalora_stock"])
                    st.info(f"[{region}] Zalora: {len(mp_dfs['Zalora']):,} EANs")
                elif rf.get("zalora_status"):
                    st.warning(f"[{region}] Zalora Stock file missing — Zalora skipped")
                if rf.get("tiktok"):
                    with st.spinner(f"[{region}] TikTok…"):
                        mp_dfs["TikTok"]=load_tiktok(rf["tiktok"])
                    st.info(f"[{region}] TikTok: {len(mp_dfs['TikTok']):,} EANs")

                # Load inventory
                inv_df=pd.DataFrame(columns=["EAN","Inv_Lazada","Inv_Shopee","Inv_Zalora","Inv_TikTok"])
                if rf.get("inventory"):
                    with st.spinner(f"[{region}] Inventory…"):
                        inv_df,inv_dbg=load_inventory(rf["inventory"],region)
                    inv_debug_all[region]=inv_dbg
                    st.info(f"[{region}] Inventory: {len(inv_df):,} EANs "
                            f"— check 🔧 Inventory Debug tab for column mapping")
                else:
                    st.warning(f"[{region}] No Inventory file uploaded — stock = 0 for all EANs")

                if not mp_dfs:
                    st.warning(f"[{region}] No marketplace files — skipping region.")
                    continue

                with st.spinner(f"Running audit for {region}…"):
                    region_results=run_audit(
                        mp_dfs,inv_df,zecom_df,content_df,override_map,region
                    )
                all_results[region]=region_results
                total=sum(
                    sum(len(v) for v in cats.values())
                    for cats in region_results.values()
                )
                st.success(f"✅ [{region}] Audit complete — {total:,} total records across all categories")

            st.session_state.audit_results=all_results
            st.session_state.inv_debug=inv_debug_all
            if all_results:
                st.success("🎉 All regions done! Go to **Results & Download** tab.")


# ── RESULTS ───────────────────────────────────────────────────────────────────
with tab_results:
    results=st.session_state.audit_results
    if not results:
        st.info("Run the audit first from the **Upload & Run** tab.")
    else:
        cats=["Active","Inactive","Not Listed - Article Level","Not Listed - Variant Level"]
        cat_icons={"Active":"🟢","Inactive":"🔴",
                   "Not Listed - Article Level":"🟡","Not Listed - Variant Level":"🟣"}

        # ── KPI row ──────────────────────────────────────────────────────────
        st.markdown("### 📊 Overview")
        totals={cat:0 for cat in cats}
        for mp_res in results.values():
            for mp_cats in mp_res.values():
                for cat in cats:
                    totals[cat]+=len(mp_cats.get(cat,pd.DataFrame()))

        k=st.columns(4)
        colors={"Active":"cg","Inactive":"cr",
                "Not Listed - Article Level":"co","Not Listed - Variant Level":"cp"}
        for i,cat in enumerate(cats):
            k[i].markdown(
                f"<div class='metric-box'>"
                f"<div class='metric-value {colors[cat]}'>{totals[cat]:,}</div>"
                f"<div class='metric-label'>{cat_icons[cat]} {cat}</div>"
                f"</div>",unsafe_allow_html=True)

        # ── Summary table ─────────────────────────────────────────────────
        st.markdown("### 📋 Region × Marketplace × Category")
        summary_rows=[]
        for region,mp_res in results.items():
            for mp in MARKETPLACES:
                if mp not in mp_res: continue
                row={"Region":region,"Marketplace":mp}
                for cat in cats:
                    row[cat]=len(mp_res[mp].get(cat,pd.DataFrame()))
                summary_rows.append(row)
        st.dataframe(pd.DataFrame(summary_rows),use_container_width=True)

        # ── Drilldown ─────────────────────────────────────────────────────
        st.markdown("### 🔍 Drilldown View")
        dr1,dr2,dr3=st.columns(3)
        dr_region=dr1.selectbox("Region",options=list(results.keys()))
        dr_mp    =dr2.selectbox("Marketplace",options=[mp for mp in MARKETPLACES
                                                        if mp in results.get(dr_region,{})])
        dr_cat   =dr3.selectbox("Category",options=cats)

        if dr_region and dr_mp and dr_cat:
            df_view=results[dr_region].get(dr_mp,{}).get(dr_cat,pd.DataFrame())
            st.markdown(f"**{cat_icons[dr_cat]} {dr_mp} [{dr_region}] — {dr_cat}: {len(df_view):,} records**")
            if not df_view.empty:
                st.dataframe(df_view,use_container_width=True,height=420)
            else:
                st.success("No records in this category.")

        # ── Download ─────────────────────────────────────────────────────
        st.markdown("---")
        st.markdown("### 💾 Download Full Report")
        with st.spinner("Building Excel report…"):
            excel_bytes=make_excel(results)
        fname=f"PUMA_Listing_Audit_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        st.download_button(
            label="📥 Download Audit Report (.xlsx)",
            data=excel_bytes, file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, type="primary"
        )
        st.caption("Sheets: Summary + Lazada/Shopee/Zalora/TikTok × Active/Inactive/"
                   "Not Listed Article/Not Listed Variant per region")


# ── INVENTORY DEBUG ───────────────────────────────────────────────────────────
with tab_debug:
    st.markdown("### 🔧 Inventory Column Mapping Debug")
    st.markdown("Use this to verify the app correctly identified stock columns in your inventory files.")
    inv_debug=st.session_state.inv_debug
    if not inv_debug:
        st.info("Run the audit first. This tab shows how inventory columns were detected.")
    else:
        for region,dbg in inv_debug.items():
            st.markdown(f"**Region: {region}**")
            rows=[{"Field":k,"Detected":str(v)} for k,v in dbg.items()]
            st.dataframe(pd.DataFrame(rows),use_container_width=True,hide_index=True)
            st.markdown("---")
        st.markdown("""
**How to fix inventory mapping issues:**
- If a channel shows `not found → using 'total_col'` — the app used the total stock column for that channel
- If a channel shows `not found → using 'None'` — no stock was mapped (stock = 0)
- Make sure your inventory file has columns containing the words: **lazada**, **shopee**, **zalora**, **tiktok**
- Or a single column with: **available**, **qty**, **stock**, **soh**, **on hand**
        """)
