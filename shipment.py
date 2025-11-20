import streamlit as st
import pandas as pd
import plotly.express as px

# -------------------- 页面配置 --------------------
st.set_page_config(page_title="Seagate Backlog Dashboard", layout="wide")
st.title("📊 Seagate SKU ETA")

# -------------------- 文件路径配置 --------------------
FILE_PATH = r"C:\Users\vnjeyu\Desktop\ASI_Daily_Backlog.xlsx"
NEW_FILE_PATH = r"C:\Users\vnjeyu\Desktop\Planning.xlsx"
NEW_LINK_FILE_PATH = r"C:\Users\vnjeyu\Desktop\Lead_Time.xlsx"

BACKORDER_SHEET = 0
SHIPMENT_SHEET = 1

SHIPMENT_MAP = {
    "Cust PO Num": "PO#",
    "Dlv Act GI Date": "Date Ship",
    "ETA (Destination Arrival Date)": "ETA",
    "Ship To City": "Ship To City",
    "Ship To Country": "Ship To Country",
    "ST Model": "ST Model",
    "Delivery Shipped Qty": "Shipped Qty",
    "House Airway Bill Num": "Tracking Number"
}

BACKORDER_MAP = {
    "Cust PO Num": "PO#",
    "Reqt Dlv Item Date": "Req Date",
    "Ship To City": "Ship To City",
    "Ship To Country": "Ship To Country",
    "ST Model": "ST Model",
    "Order Qty": "Order Qty",
    "Total Backlog Qty": "Backlog Qty"
}

# -------------------- 缓存加载函数 --------------------
@st.cache_data
def load_excel(path, sheet=None):
    """加载 Excel 文件并返回 DataFrame"""
    try:
        df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
        if isinstance(df, dict):
            df = list(df.values())[0]
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception as e:
        st.error(f"❌ 文件加载失败: {e}")
        return pd.DataFrame()

@st.cache_data
def load_filtered_stmodel(path):
    """加载并过滤 ST Model 数据"""
    df = load_excel(path)
    keep_cols = ["Product ST Model Num", "Key Figure"] + [c for c in df.columns if c.startswith(("OCT", "NOV", "DEC", "JAN")) or "Q" in c]
    df = df[keep_cols]
    valid_figures = ["Backlog", "Shipments", "SI UCD Final", "Supply Commit (Channel)"]
    df = df[df["Key Figure"].str.strip().isin(valid_figures)]
    df["Product ST Model Num"] = df["Product ST Model Num"].astype(str).str.strip()
    return df

def load_and_prepare(sheet, rename_map):
    """加载并重命名列，清理数据"""
    df = load_excel(FILE_PATH, sheet)
    if df.empty:
        return df
    cols = [c for c in rename_map.keys() if c in df.columns]
    df = df[cols].copy()
    df.rename(columns=rename_map, inplace=True)

    # 日期列转换
    for col in ["Date Ship", "ETA", "Req Date"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # 字符列清理
    for col in ["PO#", "Ship To City", "Ship To Country", "ST Model"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    return df

# -------------------- 加载数据 --------------------
stmodel_df = load_filtered_stmodel(NEW_FILE_PATH)
shipment_df = load_and_prepare(SHIPMENT_SHEET, SHIPMENT_MAP)
backorder_df = load_and_prepare(BACKORDER_SHEET, BACKORDER_MAP)
link_df = load_excel(NEW_LINK_FILE_PATH)

# ✅ 统一 ST Model 格式
for df in [link_df, shipment_df, backorder_df]:
    if "ST MODEL" in df.columns:
        df["ST MODEL"] = df["ST MODEL"].astype(str).str.strip()
    if "ST Model" in df.columns:
        df["ST Model"] = df["ST Model"].astype(str).str.strip()

# -------------------- Sidebar Filters --------------------
with st.sidebar:
    st.header("🔍 Shared Filters")
    search_query = st.text_input("Search by ST Model or PO#", "", key="search_st_model")
    sku_query = st.text_input("Search by SKU", "", key="search_sku")
    countries = sorted(set(shipment_df["Ship To Country"].dropna()) | set(backorder_df["Ship To Country"].dropna()))
    cities = sorted(set(shipment_df["Ship To City"].dropna()) | set(backorder_df["Ship To City"].dropna()))
    country_sel = st.multiselect("Filter by Country", countries, key="filter_country")
    city_sel = st.multiselect("Filter by City", cities, key="filter_city")

# -------------------- SKU 映射 --------------------
sku_models = []
if sku_query.strip():
    matched_rows = link_df[link_df["SKU"].astype(str).str.lower().str.contains(sku_query.lower(), na=False)]
    if not matched_rows.empty:
        sku_models = matched_rows["ST MODEL"].dropna().astype(str).str.strip().tolist()

# -------------------- Filter Function --------------------
def apply_filters(df, date_col):
    """根据搜索条件和筛选项过滤数据"""
    filtered = df.copy()
    mask = pd.Series([True] * len(filtered))

    if search_query.strip():
        q = search_query.lower()
        mask = mask & (
            filtered["PO#"].str.lower().str.contains(q, na=False) |
            filtered["ST Model"].str.lower().str.contains(q, na=False)
        )
    if sku_models:
        mask = mask & filtered["ST Model"].isin(sku_models)

    filtered = filtered[mask]

    if country_sel:
        filtered = filtered[filtered["Ship To Country"].isin(country_sel)]
    if city_sel:
        filtered = filtered[filtered["Ship To City"].isin(city_sel)]

    if date_col in filtered.columns:
        sort_key = filtered[date_col].fillna(pd.Timestamp.min)
        filtered = filtered.assign(_sort_key=sort_key).sort_values("_sort_key", ascending=False).drop(columns="_sort_key")
    return filtered

shipment_filtered = apply_filters(shipment_df, "Date Ship")
backorder_filtered = apply_filters(backorder_df, "Req Date")

# -------------------- Highlight Function --------------------
def highlight_columns(df):
    """高亮特定月份和季度列"""
    yellow_cols = [c for c in df.columns if ("NOV" in c or "DEC" in c or "Q2" in c)]
    green_cols = [c for c in df.columns if ("JAN" in c or "Q3" in c)]

    def highlight(val, col):
        if col in yellow_cols:
            return 'background-color: #fff2cc'
        elif col in green_cols:
            return 'background-color: #e2f0d9'
        return ''

    styled = df.style.apply(lambda row: [highlight(v, col) for col, v in zip(df.columns, row)], axis=1)
    num_cols = df.select_dtypes(include=['number']).columns
    styled = styled.format({col: "{:.0f}" for col in num_cols})
    return styled

# -------------------- 判断是否有有效匹配 --------------------
has_valid_match = False
if search_query.strip() or sku_query.strip():
    if search_query.strip():
        has_valid_match = not stmodel_df[stmodel_df["Product ST Model Num"].str.lower().str.contains(search_query.lower(), na=False)].empty
    if sku_query.strip():
        has_valid_match = has_valid_match or bool(sku_models)

# -------------------- 页面内容 --------------------
st.markdown("---")
if has_valid_match:
    # 📅 Timeline
    st.subheader("📅 Timeline")
    filtered_stmodel = stmodel_df.copy()
    if search_query.strip():
        filtered_stmodel = filtered_stmodel[
            filtered_stmodel["Product ST Model Num"].str.lower().str.contains(search_query.lower(), na=False)
        ]
    if sku_models:
        filtered_stmodel = filtered_stmodel[
            filtered_stmodel["Product ST Model Num"].isin(sku_models)
        ]

    # 显示高亮表格
    st.dataframe(highlight_columns(filtered_stmodel), use_container_width=True)

    # ---- Bar Chart for selected quarters ----
    import plotly.express as px

    wanted_cols = ["Q2 2026", "Q3 2026", "Q4 2026", "Q1 2027", "Q2 2027"]
    available_cols = [c for c in wanted_cols if c in filtered_stmodel.columns]

    if len(available_cols) == 0:
        st.info("No matching quarter columns found.")
    else:
        long_df = filtered_stmodel.melt(
            id_vars=["Product ST Model Num", "Key Figure"],
            value_vars=available_cols,
            var_name="Quarter",
            value_name="Value"
        )

        long_df["Value"] = pd.to_numeric(long_df["Value"], errors="coerce").fillna(0)
        long_df = long_df[long_df["Value"] != 0]

        if long_df.empty:
            st.warning("Selected columns have no non-zero values for current filters.")
        else:
            fig = px.bar(
                long_df,
                x="Value",
                y="Key Figure",
                color="Quarter",
                orientation="h",
                title="📊 ST Model vs Quarters",
                hover_data=["Key Figure"]
            )
            fig.update_layout(
                height=600,
                xaxis_title="Value",
                yaxis_title="Key Figure",
                legend_title_text="Quarter"
            )
            st.plotly_chart(fig, use_container_width=True)

    # 🚚 Shipment Details
    st.subheader("🚚 Shipment Details")
    st.dataframe(shipment_filtered, use_container_width=True)

    st.markdown("---")
    # 📦 Backorder Details
    st.subheader("📦 Backorder Details")
    st.dataframe(backorder_filtered, use_container_width=True)

    # 📌 ETA & Notes
    st.markdown("---")
    st.subheader("📌 Make NEW PO")

    filtered_link = link_df.copy()
    if search_query.strip():
        filtered_link = filtered_link[
            filtered_link["ST MODEL"].str.lower().str.contains(search_query.lower(), na=False)
        ]
    if sku_query.strip():
        filtered_link = filtered_link[
            filtered_link["SKU"].astype(str).str.lower().str.contains(sku_query.lower(), na=False)
        ]

    if not filtered_link.empty:
        eta_value = str(filtered_link.iloc[0].get("ETA", "N/A"))
        note_value = str(filtered_link.iloc[0].get("Note", "No notes available"))

        st.markdown(
            f"""
            <div style="display:flex; justify-content:center; align-items:center; margin-top:20px;">
                <div style="
                    background-color:#f0f8ff;
                    padding:40px;
                    border-radius:15px;
                    width:60%;
                    text-align:center;
                    box-shadow: 0 4px 8px rgba(0,0,0,0.2);
                ">
                    <h2 style="color:#333; font-size:36px; margin-bottom:10px;">ETA</h2>
                    <p style="font-size:48px; font-weight:bold; color:#0078D7;">{eta_value}</p>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        st.markdown(
            f"""
            <div style="text-align:center; margin-top:20px;">
                <h3 style="color:#555;">Note</h3>
                <p style="font-size:20px; color:#666;">{note_value}</p>
            </div>
            """,
            unsafe_allow_html=True
        )
    else:
        st.warning("No matching SKU or ST Model found in ETA/Notes file.")
else:
    st.warning("⚠️ No matching ST Model or SKU found. Please check your input or try different filters.")
