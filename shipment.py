
import re
import pandas as pd
import plotly.express as px
import streamlit as st

# -------------------- 页面配置 --------------------
st.set_page_config(page_title="Seagate Backlog Dashboard", layout="wide")
st.title("📊 Seagate SKU ETA")

# -------------------- 文件路径（相对路径） --------------------
# 如果文件在仓库根目录：
FILE_PATH = "ASI_Daily_Backlog.xlsx"
NEW_FILE_PATH = "Planning.xlsx"
NEW_LINK_FILE_PATH = "Lead_Time.xlsx"

# -------------------- Sheet & 列映射 --------------------
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
    "House Airway Bill Num": "Tracking Number",
}

BACKORDER_MAP = {
    "Cust PO Num": "PO#",
    "Reqt Dlv Item Date": "Req Date",
    "Ship To City": "Ship To City",
    "Ship To Country": "Ship To Country",
    "ST Model": "ST Model",
    "Order Qty": "Order Qty",
    "Total Backlog Qty": "Backlog Qty",
}

# -------------------- 缓存加载函数 --------------------
@st.cache_data
def load_excel(path, sheet=None):
    try:
        df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
        # 如果 sheet_name=None 或为索引，pandas 可能返回 dict；此处取第一个表
        if isinstance(df, dict):
            df = list(df.values())[0]
        # 规范列名
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception as e:
        st.error(f"❌ 文件加载失败: {e}")
        return pd.DataFrame()

@st.cache_data
def load_filtered_stmodel(path):
    """
    加载并过滤 ST Model 数据：
    - 保留所有月份列（JAN~DEC，含类似 'JAN-24 W31-26' 前缀的列名）
    - 保留季度列（任意包含 'Q' 的列，如 'Q3 2026'）
    - 保留 'Product ST Model Num' 与 'Key Figure'
    - 仅保留指定 Key Figure 的行
    """
    df = load_excel(path)
    if df.empty:
        return df

    # 统一清理列名
    df.columns = [str(c).strip() for c in df.columns]

    # 基准列
    base_cols = ["Product ST Model Num", "Key Figure"]

    # 月份缩写（英文）
    month_abbr = ("MAR", "APR", "MAY", "JUN",
                  "JUL", "AUG", "SEP", "OCT", "NOV", "DEC")

    # 动态识别：以月份缩写开头的列（支持 'JAN-17 W30-26' 这类）
    month_cols = [
        c for c in df.columns
        if isinstance(c, str) and c.strip().upper().startswith(month_abbr)
    ]

    # 动态识别：季度列（名字里含 'Q'；若需要更严格可用正则 '^Q[1-4]\\s\\d{4}$'）
    quarter_cols = [
        c for c in df.columns
        if isinstance(c, str) and ("Q" in c.upper())
    ]

    # 合并并去重，保证列存在
    keep_cols = []
    for col in base_cols + month_cols + quarter_cols:
        if col in df.columns and col not in keep_cols:
            keep_cols.append(col)

    df = df[keep_cols].copy()

    # 过滤 Key Figure
    valid_figures = ["Backlog", "Shipments", "SI UCD Final", "Supply Commit (Channel)"]
    if "Key Figure" in df.columns:
        df = df[df["Key Figure"].astype(str).str.strip().isin(valid_figures)]

    # 规范 ST 型号
    if "Product ST Model Num" in df.columns:
        df["Product ST Model Num"] = df["Product ST Model Num"].astype(str).str.strip()

    return df

def load_and_prepare(sheet, rename_map):
    """加载指定 sheet 并做列重命名与基础清理"""
    df = load_excel(FILE_PATH, sheet)
    if df.empty:
        return df

    # 仅保留映射中存在的列
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

# -------------------- SKU -> ST MODEL 映射 --------------------
sku_models = []
if sku_query.strip():
    matched_rows = link_df[link_df["SKU"].astype(str).str.lower().str.contains(sku_query.lower(), na=False)]
    if not matched_rows.empty:
        sku_models = matched_rows["ST MODEL"].dropna().astype(str).str.strip().tolist()

# -------------------- 通用过滤函数 --------------------
def apply_filters(df, date_col):
    """根据搜索和筛选条件过滤数据"""
    if df.empty:
        return df

    filtered = df.copy()
    mask = pd.Series([True] * len(filtered), index=filtered.index)

    if search_query.strip():
        q = search_query.lower()
        # 注意：PO# / ST Model 两列都必须存在才会参与匹配
        conds = []
        if "PO#" in filtered.columns:
            conds.append(filtered["PO#"].astype(str).str.lower().str.contains(q, na=False))
        if "ST Model" in filtered.columns:
            conds.append(filtered["ST Model"].astype(str).str.lower().str.contains(q, na=False))
        if conds:
            mask = mask & conds[0]
            for c in conds[1:]:
                mask = mask | c  # PO# 或 ST Model 之一匹配即可

    if sku_models and "ST Model" in filtered.columns:
        mask = mask & filtered["ST Model"].isin(sku_models)

    filtered = filtered[mask]

    if country_sel and "Ship To Country" in filtered.columns:
        filtered = filtered[filtered["Ship To Country"].isin(country_sel)]
    if city_sel and "Ship To City" in filtered.columns:
        filtered = filtered[filtered["Ship To City"].isin(city_sel)]

    if date_col in filtered.columns:
        sort_key = filtered[date_col].fillna(pd.Timestamp.min)
        filtered = filtered.assign(_sort_key=sort_key).sort_values("_sort_key", ascending=False).drop(columns="_sort_key")
    return filtered

shipment_filtered = apply_filters(shipment_df, "Date Ship")
backorder_filtered = apply_filters(backorder_df, "Req Date")

# -------------------- 判断是否有有效匹配 --------------------
has_valid_match = False
if search_query.strip() or sku_query.strip():
    if search_query.strip() and "Product ST Model Num" in stmodel_df.columns:
        has_valid_match = not stmodel_df[stmodel_df["Product ST Model Num"].astype(str).str.lower().str.contains(search_query.lower(), na=False)].empty
    if sku_query.strip():
        has_valid_match = has_valid_match or bool(sku_models)

# -------------------- 页面内容 --------------------
    
st.markdown("---")
if has_valid_match:
    # 📅 Timeline
    st.subheader("📅 Timeline")

    filtered_stmodel = stmodel_df.copy()

    # 1) 先按现有条件过滤
    if search_query.strip() and "Product ST Model Num" in filtered_stmodel.columns:
        filtered_stmodel = filtered_stmodel[
            filtered_stmodel["Product ST Model Num"].astype(str).str.lower().str.contains(search_query.lower(), na=False)
        ]
    if sku_models and "Product ST Model Num" in filtered_stmodel.columns and "ST MODEL" in link_df.columns:
        filtered_stmodel = filtered_stmodel[
            filtered_stmodel["Product ST Model Num"].isin(sku_models)  # 按需调整
        ]

    # 2) ✅ 在 Timeline 数据中合并 SKU（最小改动）
    if "Product ST Model Num" in filtered_stmodel.columns and "ST MODEL" in link_df.columns and "SKU" in link_df.columns:
        # 规范字符串，避免空格/大小写导致匹配失败
        filtered_stmodel["Product ST Model Num"] = filtered_stmodel["Product ST Model Num"].astype(str).str.strip()
        link_df["ST MODEL"] = link_df["ST MODEL"].astype(str).str.strip()
        link_df["SKU"] = link_df["SKU"].astype(str).str.strip()  # 强制 SKU 为文本（General）

        filtered_stmodel = filtered_stmodel.merge(
            link_df[["ST MODEL", "SKU"]].rename(columns={"ST MODEL": "Product ST Model Num"}),
            on="Product ST Model Num",
            how="left"
        )
    else:
        st.warning("Timeline 缺少关键列：'Product ST Model Num' 或 link_df 缺少 'ST MODEL'/'SKU'，无法合并 SKU。")

    if "SKU" in filtered_stmodel.columns:
        cols = ["SKU"] + [c for c in filtered_stmodel.columns if c != "SKU"]
        filtered_stmodel = filtered_stmodel[cols]

    # ✅ 显示所有列（支持水平滚动）——此时表格里已包含 SKU 列
    st.dataframe(filtered_stmodel, use_container_width=True, hide_index=False)

    # ---- Bar Chart：动态识别所有季度列 ----
    quarter_cols = [
        c for c in filtered_stmodel.columns
        if isinstance(c, str) and re.match(r'^Q[1-4]\s\d{4}$', c.strip().upper())
    ]

    if len(quarter_cols) == 0:
        st.info("No matching quarter columns found (expected like 'Qx YYYY').")
    else:
        def q_sort_key(c: str):
            c = c.strip().upper()  # 'Q2 2026'
            q, y = c.split()
            return (int(y), int(q[1]))  # (年份, 季度)

        quarter_cols_sorted = sorted(quarter_cols, key=q_sort_key)

        # 选择 id_vars（存在才加入）
        id_vars = []
        if "Product ST Model Num" in filtered_stmodel.columns:
            id_vars.append("Product ST Model Num")
        if "Key Figure" in filtered_stmodel.columns:
            id_vars.append("Key Figure")
        # （可选）把 SKU 放进 hover 信息
        if "SKU" in filtered_stmodel.columns:
            id_vars.append("SKU")

        if len(id_vars) == 0:
            st.warning("Missing required id columns for chart (e.g., 'Product ST Model Num', 'Key Figure').")
        else:
            long_df = filtered_stmodel.melt(
                id_vars=id_vars,
                value_vars=quarter_cols_sorted,
                var_name="Quarter",
                value_name="Value"
            )
            long_df["Value"] = pd.to_numeric(long_df["Value"], errors="coerce").fillna(0)
            long_df = long_df[long_df["Value"] != 0]

            if long_df.empty:
                st.warning("Selected quarter columns have no non-zero values for current filters.")
            else:
                # 如果没有 Key Figure，则以第一个 id_vars 为 Y 轴
                y_axis = "Key Figure" if "Key Figure" in long_df.columns else id_vars[0]
                fig = px.bar(
                    long_df,
                    x="Value",
                    y=y_axis,
                    color="Quarter",
                    orientation="h",
                    title="📊 ST Model vs Quarters",
                    hover_data=id_vars,  # 悬浮信息包含 SKU（若存在）
                    category_orders={"Quarter": quarter_cols_sorted}
                )
                fig.update_layout(
                    height=600,
                    xaxis_title="Value",
                    yaxis_title=y_axis,
                    legend_title_text="Quarter"
                )
                st.plotly_chart(fig, use_container_width=True)

    # 🚚 Shipment Details
    st.subheader("🚚 Shipment Details")
    st.dataframe(shipment_filtered, use_container_width=True)


 
    st.markdown("---")
    st.subheader("📦 Backorder Details")
    if "Req Date" in backorder_filtered.columns:
        backorder_filtered = backorder_filtered.sort_values("Req Date", ascending=True)
    st.dataframe(backorder_filtered, use_container_width=True)


    # 📌 ETA & Notes
    st.markdown("---")
    st.subheader("📌 Make NEW PO")

    filtered_link = link_df.copy()
    if search_query.strip() and "ST MODEL" in filtered_link.columns:
        filtered_link = filtered_link[
            filtered_link["ST MODEL"].astype(str).str.lower().str.contains(search_query.lower(), na=False)
        ]
    if sku_query.strip() and "SKU" in filtered_link.columns:
        filtered_link = filtered_link[
            filtered_link["SKU"].astype(str).str.lower().str.contains(sku_query.lower(), na=False)
        ]

    if not filtered_link.empty:
        eta_value = str(filtered_link.iloc[0].get("ETA", "N/A"))
        note_value = str(filtered_link.iloc[0].get("Note", "No notes available"))

        # ✅ 使用真实 HTML 标签（不再转义）
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

st.markdown("---")
st.subheader("🆕Today’s Shipment")
new_ship_df = shipment_df.copy()
cols_to_show = ["PO#", "Date Ship", "ETA", "Ship To City", "Ship To Country","ST Model", "Shipped Qty", "Tracking Number"]
cols_to_show = [c for c in cols_to_show if c in new_ship_df.columns]
if not cols_to_show:
    st.warning("The shipment data is missing expected columns for New’s Shipment.")
else:
    new_ship_df = new_ship_df[cols_to_show]

if "Date Ship" in new_ship_df.columns:
    sort_key = new_ship_df["Date Ship"].fillna(pd.Timestamp.min)
    new_ship_df = (
        new_ship_df
        .assign(_sort_key=sort_key)
        .sort_values("_sort_key", ascending=False)
        .drop(columns="_sort_key")
    )

TOP_N = 10
new_ship_top = new_ship_df.head(TOP_N)
st.dataframe(new_ship_top, use_container_width=True)

















