import streamlit as st
import pandas as pd
import gspread
from google.oauth2.credentials import Credentials
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import datetime
from datetime import timedelta

# =========================================================
# 1. 頁面基本設定 & CSS (最終定案版 V5.3 - 修正 Radio 按鈕樣式)
# =========================================================
st.set_page_config(page_title="全校電力管理平台", page_icon="⚡", layout="wide")

# 定義進階莫蘭迪色系 & 樣式變數
COLORS = {
    "bg_light": "#F7F9F9",
    "primary_blue": "#5D6D7E",
    "btn_bg": "#D6EAF8",       # 淺藍底
    "btn_text": "#154360",     # 深藍字
    "text_main": "#2C3E50",    # 深灰字
    "text_axis": "#424949",    # 座標軸深灰
    "num_orange": "#E59866",
    "kpi_bg_orange": "#F5CBA7", # 莫蘭迪橘
    "chart_usage_blue": "#5DADE2", # 莫蘭迪藍
    "chart_cost_green": "#88B04B", # 莫蘭迪綠
    "text_black": "#000000",       # 純黑
    "line_prev_year": "#9E9E9E",   # 前一年
    "campus": {
        "蘭潭校區": "#A9CCE3",
        "民雄校區": "#F9E79F",
        "新民校區": "#D2B4DE",
        "林森校區": "#A3E4D7",
        "全校": "#B0BEC5"
    }
}

# 地址排序邏輯
ADDRESS_ORDER = {
    "蘭潭校區": [
        "蘭潭教學行政區", "蘭潭畜牧場區", "蘭潭宿舍區", 
        "八掌段803-1地號農地(實習農場)", "短竹段999-37地號溫網室(實習農場)", 
        "中埔鄉社口10-1號表燈", "中埔鄉社口10號表燈", 
        "水上鄉新三界埔段1213地號", "水上鄉三界埔14地號"
    ],
    "民雄校區": ["民雄主校區", "民雄-大學館創意樓", "民雄學生宿舍"],
    "新民校區": ["新民主校區", "新民動物醫院及泳池區", "新民賢德樓宿舍"],
    "林森校區": ["林森主校區", "林森校區-樂育堂", "林森學生宿舍", "進德樓學生宿舍公共", "首長宿舍"]
}

st.markdown(f"""
    <style>
    /* 修復頁首色差 */
    [data-testid="stHeader"] { background-color: transparent !important; }
    .stApp { font-family: "Microsoft JhengHei", sans-serif; background-color: #F4F6F7; } /* 換成帶冷色調的灰白，完美襯托藍色UI */
    
    /* 1. 分頁標籤按鈕 (統一深色質感與橘色點綴) */
    div[data-testid="stTabs"] button[data-baseweb="tab"] {{
        background-color: #384959 !important; 
        border-radius: 8px 8px 0 0 !important;
        padding: 12px 25px !important;
        border: none !important;
        margin-right: 4px !important;
    }}
    div[data-testid="stTabs"] button[data-baseweb="tab"] > div {{
        font-size: 20px !important; 
        color: #FFFFFF !important;
        font-weight: 600 !important;
    }}
    div[data-testid="stTabs"] button[data-baseweb="tab"][aria-selected="true"] {{
        background-color: #1D2631 !important; 
        border-top: 4px solid #F39C12 !important;
        border-bottom: none !important; 
    }}
    div[data-testid="stTabs"] button[data-baseweb="tab"][aria-selected="true"] > div {{
        color: #F39C12 !important;
    }}

    /* 2. 資訊卡片 (KPI) */
    .kpi-card {{
        background-color: white;
        border-radius: 12px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        border: 1px solid #E5E7E9;
        margin-bottom: 20px;
        overflow: hidden;
    }}
    .kpi-header {{
        color: {COLORS['text_main']};
        padding: 10px 15px;
        font-weight: 800;
        font-size: 1.3rem;
        text-align: center;
        border-bottom: 2px solid rgba(0,0,0,0.05);
    }}
    .kpi-content {{
        display: flex;
        flex-direction: row;
        padding: 12px;
        justify-content: space-around;
        align-items: center;
    }}
    .kpi-section {{
        text-align: center;
        border-right: 1px solid #F2F3F4;
    }}
    .kpi-section:last-child {{ border-right: none; }}
    
    .kpi-label {{
        font-size: 1.0rem;
        color: {COLORS['text_main']};
        font-weight: 600;
        margin-bottom: 5px;
    }}
    
    /* 數字大小樣式類別 */
    .num-sm {{ font-size: 1.4rem; font-weight: 800; color: {COLORS['num_orange']}; }}
    .unit-sm {{ font-size: 0.9rem; color: {COLORS['text_main']}; font-weight: 600; margin-left: 3px; }}
    .num-md {{ font-size: 1.8rem; font-weight: 800; color: {COLORS['num_orange']}; }}
    .unit-md {{ font-size: 1.1rem; color: {COLORS['text_main']}; font-weight: 600; margin-left: 3px; }}
    .num-lg {{ font-size: 2.4rem; font-weight: 800; color: {COLORS['num_orange']}; }} 
    .unit-lg {{ font-size: 1.4rem; color: {COLORS['text_main']}; font-weight: 600; margin-left: 3px; }}

    /* 3. 按鈕與下拉選單樣式 - 🔥 修正 Radio 選擇器，強制套用淺藍底深藍字 */
    div.stButton > button {{
        background-color: {COLORS['btn_bg']} !important;
        color: {COLORS['btn_text']} !important;
        font-weight: 800 !important;
        border: 1px solid #AED6F1 !important;
        border-radius: 8px !important;
    }}
    div.stButton > button:hover {{
        background-color: #AED6F1 !important;
        color: {COLORS['btn_text']} !important;
    }}
    
    .stRadio div[role="radiogroup"] label {{
        background-color: {COLORS['btn_bg']} !important;
        border: 1px solid #AED6F1 !important;
        border-radius: 8px !important;
        padding: 8px 15px !important;
        margin-right: 10px !important;
    }}
    .stRadio div[role="radiogroup"] label p {{
        font-size: 1.1rem !important;
        font-weight: 800 !important;
        color: {COLORS['btn_text']} !important;
    }}
    .stRadio div[role="radiogroup"] label[data-checked="true"] {{
        background-color: #1A5276 !important;
        border-color: #1A5276 !important;
    }}
    .stRadio div[role="radiogroup"] label[data-checked="true"] p {{
        color: #FFFFFF !important;
    }}
    
    div[data-baseweb="select"] span {{
        color: {COLORS['text_main']} !important;
        font-weight: 600;
    }}
    
    /* 4. 圖表區塊框線 */
    .chart-box {{
        border: 1px solid #BDC3C7;
        border-radius: 8px;
        padding: 10px;
        background-color: white;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        margin-bottom: 15px;
    }}

    .js-plotly-plot .plotly .main-svg {{
        background: rgba(0,0,0,0) !important;
    }}

    /* 5. 置中標題樣式 (+底色與圓角) */
    .section-title {{
        text-align: center;
        font-size: 2.2rem;
        font-weight: 900;
        color: #1A5276;
        background-color: #EBF5FB;
        padding: 15px 20px;
        border-radius: 12px;
        margin-bottom: 25px;
        margin-top: 20px;
        letter-spacing: 1px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }}
    .section-subtitle {{
        text-align: center;
        font-size: 1.6rem;
        font-weight: 900;
        color: #2C3E50;
        background-color: #F2F4F4;
        padding: 12px 20px;
        border-radius: 10px;
        margin-bottom: 20px;
        margin-top: 10px;
    }}

    /* 6. 展開面板 (Expander) 樣式優化 */
    [data-testid="stExpander"] {{
        background-color: #FFFFFF;
        border: 1px solid #BDC3C7;
        border-radius: 12px;
        box-shadow: 0 3px 6px rgba(0,0,0,0.08);
        margin-bottom: 15px;
    }}
    [data-testid="stExpander"] summary {{ padding: 10px 15px; }}
    [data-testid="stExpander"] summary p {{
        font-size: 1.4rem !important;
        font-weight: 900 !important;
        color: #000000 !important;
    }}
    [data-testid="stExpander"] summary:hover {{ background-color: #F8F9F9; }}
    [data-testid="stExpanderDetails"] {{
        padding: 20px;
        background-color: #EAECEE;
        border-top: 1px solid #BDC3C7;
        border-radius: 0 0 12px 12px;
    }}

    /* 左側欄登出按鈕專屬樣式 (含深橘紅邊框設計) */
    [data-testid="stSidebar"] div.stButton > button {{
        background-color: #E67E22 !important; 
        color: #FFFFFF !important; 
        border: 2px solid #D35400 !important; 
        font-weight: 600 !important; 
        font-size: 18px !important; 
        padding: 10px 32px !important; 
        border-radius: 8px !important;
        width: 100% !important; 
        margin-top: 15px !important;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1) !important;
    }}
    [data-testid="stSidebar"] div.stButton > button:hover {{
        background-color: #D35400 !important; 
        border: 2px solid #BA4A00 !important; 
        transform: translateY(-2px);
    }}
    </style>
""", unsafe_allow_html=True)

# =========================================================
# 2. 資料處理函數 (核心邏輯：使用期間按日攤算)
# =========================================================
SHEET_ID = "17BTvriIxK1ibFaPlmDUaTYwGrbmq556XNbzythA3rVo"

@st.cache_resource
def init_google_sheet():
    try:
        creds_dict = dict(st.secrets["gcp_oauth"])
        creds = Credentials.from_authorized_user_info(creds_dict, scopes=["https://www.googleapis.com/auth/spreadsheets"])
        gc = gspread.authorize(creds)
        return gc
    except Exception as e:
        st.error(f"連線失敗: {e}")
        return None

def update_google_sheet_data(updated_df):
    try:
        gc = init_google_sheet()
        sh = gc.open_by_key(SHEET_ID)
        ws = sh.worksheet("用電填報紀錄")
        data_to_upload = [updated_df.columns.values.tolist()] + updated_df.astype(str).values.tolist()
        ws.clear()
        ws.update(data_to_upload)
        st.cache_data.clear() 
        return True
    except Exception as e:
        st.error(f"儲存失敗: {e}")
        return False

def calculate_daily_proration(df, df_coef):
    daily_records = []
    coef_map = {}
    if not df_coef.empty:
        for _, row in df_coef.iterrows():
            try: coef_map[int(row['年度'])] = float(row['排碳係數'])
            except: pass
    
    for _, row in df.iterrows():
        try:
            start_date = pd.to_datetime(row['使用期間(起)']).date()
            end_date = pd.to_datetime(row['使用期間(訖)']).date()
            
            if pd.isnull(start_date) or pd.isnull(end_date) or start_date > end_date: continue
            
            days_count = (end_date - start_date).days + 1
            if days_count <= 0: continue

            total_kwh = float(row['用電量(度數)'])
            total_cost = float(row['電費金額(元)'])
            
            daily_kwh = total_kwh / days_count
            daily_cost = total_cost / days_count
            
            current_date = start_date
            while current_date <= end_date:
                y = current_date.year
                coef = coef_map.get(y, 0.495)
                daily_co2 = daily_kwh * coef / 1000.0
                
                daily_records.append({
                    "date": current_date, 
                    "year": y, 
                    "month": current_date.month,
                    "校區": row['校區'], 
                    "用電地址": row['用電地址'], 
                    "電號": str(row['電號']),
                    "daily_kwh": daily_kwh, 
                    "daily_cost": daily_cost, 
                    "daily_co2": daily_co2
                })
                current_date += timedelta(days=1)
        except: continue

    if not daily_records: return pd.DataFrame()
    
    df_daily = pd.DataFrame(daily_records)
    df_grouped = df_daily.groupby(['year', 'month', '校區', '用電地址', '電號'], as_index=False).agg({
        'daily_kwh': 'sum', 'daily_cost': 'sum', 'daily_co2': 'sum'
    })
    
    df_grouped.rename(columns={
        'year': '統計年度', 
        'month': '統計月份',
        'daily_kwh': '用電量(度數)', 
        'daily_cost': '電費金額(元)', 
        'daily_co2': '碳排放量'
    }, inplace=True)
    
    return df_grouped

@st.cache_data(ttl=60)
def load_and_process_data(_gc):
    sh = _gc.open_by_key(SHEET_ID)
    try:
        ws_records = sh.worksheet("用電填報紀錄")
        df_raw = pd.DataFrame(ws_records.get_all_records())
    except:
        return pd.DataFrame(), pd.DataFrame()
        
    try:
        ws_coef = sh.worksheet("碳排係數管理")
        df_coef = pd.DataFrame(ws_coef.get_all_records())
        df_coef.columns = [str(c).strip() for c in df_coef.columns]
    except:
        df_coef = pd.DataFrame(columns=["年度", "排碳係數"])

    if df_raw.empty: return pd.DataFrame(), pd.DataFrame()
    
    df_raw['用電量(度數)'] = pd.to_numeric(df_raw['用電量(度數)'], errors='coerce').fillna(0)
    df_raw['電費金額(元)'] = pd.to_numeric(df_raw['電費金額(元)'], errors='coerce').fillna(0)
    df_raw['帳單年度'] = pd.to_numeric(df_raw['帳單年度'], errors='coerce')
    
    df_processed = calculate_daily_proration(df_raw, df_coef)
    return df_processed, df_raw

def get_year_range_string(df, year):
    df_y = df[df['統計年度'] == year]
    if df_y.empty: return f"{year}年"
    months = sorted(df_y['統計月份'].unique())
    if not months: return f"{year}年"
    return f"{year}年 {months[0]}~{months[-1]}月"

def sort_addresses(df_subset, campus_name):
    order = ADDRESS_ORDER.get(campus_name, [])
    if not order: return df_subset
    df_subset['sort_idx'] = df_subset['用電地址'].apply(lambda x: order.index(x) if x in order else 999)
    return df_subset.sort_values('sort_idx').drop(columns=['sort_idx'])

# =========================================================
# 3. 繪圖與渲染函數
# =========================================================

def render_kpi_card_custom(title, col1, col2, col3, color_code, ratios=[1,1,1], font_class="md"):
    html = f"""
    <div class="kpi-card">
        <div class="kpi-header" style="background-color: {color_code};">{title}</div>
        <div class="kpi-content">
            <div class="kpi-section" style="flex: {ratios[0]};">
                <div class="kpi-label">{col1[0]}</div>
                <div class="num-{font_class}">{col1[1]}<span class="unit-{font_class}">{col1[2]}</span></div>
            </div>
            <div class="kpi-section" style="flex: {ratios[1]};">
                <div class="kpi-label">{col2[0]}</div>
                <div class="num-{font_class}">{col2[1]}<span class="unit-{font_class}">{col2[2]}</span></div>
            </div>
            <div class="kpi-section" style="flex: {ratios[2]};">
                <div class="kpi-label">{col3[0]}</div>
                <div class="num-{font_class}">{col3[1]}<span class="unit-{font_class}">{col3[2]}</span></div>
            </div>
        </div>
    </div>
    """
    st.markdown(html, unsafe_allow_html=True)

def plot_single_bar_chart(df_chart, x_col, y_col, title, color, unit):
    all_months = pd.DataFrame({x_col: range(1, 13)})
    grouped = df_chart.groupby(x_col)[y_col].sum().reset_index()
    merged = pd.merge(all_months, grouped, on=x_col, how='left').fillna(0)
    
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=merged[x_col], y=merged[y_col],
        marker_color=color,
        text=merged[y_col].apply(lambda x: f"{int(x):,}" if x>0 else ""),
        textposition='outside', 
        textfont=dict(color='black', size=18) 
    ))
    
    axis_font = dict(size=18, color=COLORS['text_axis'])
    
    fig.update_layout(
        title=dict(text=title, font=dict(size=20, color=COLORS['text_main'], family="Microsoft JhengHei")),
        plot_bgcolor='rgba(0,0,0,0)',
        margin=dict(l=10, r=10, t=50, b=10),
        height=400,
        hovermode="x unified"
    )
    fig.update_xaxes(tickmode='linear', dtick=1, title_text="月份 (Month)", tickfont=axis_font, title_font=axis_font, showgrid=False)
    fig.update_yaxes(title_text=unit, tickfont=axis_font, title_font=axis_font, showgrid=True, gridcolor='#EAEDED')
    
    if not merged.empty:
        max_val = merged[y_col].max()
        fig.update_yaxes(range=[0, max_val * 1.25])
        
    return fig

def plot_horizontal_ranking(df, category_col, value_col, total_value, title, height=400):
    grouped = df.groupby(category_col)[value_col].sum().reset_index()
    grouped = grouped.sort_values(value_col, ascending=True) 
    
    grouped['label'] = grouped[value_col].apply(
        lambda x: f"{int(x):,} ({x/total_value*100:.1f}%)" if total_value > 0 else f"{int(x):,}"
    )
    
    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=grouped[category_col], x=grouped[value_col],
        orientation='h',
        marker_color=COLORS['chart_usage_blue'],
        text=grouped['label'],
        textposition='outside',
        textfont=dict(color='black', size=16)
    ))
    
    axis_font = dict(size=18, color=COLORS['text_axis'])
    
    fig.update_layout(
        title=dict(text=title, font=dict(size=20, color=COLORS['text_main'])),
        plot_bgcolor='rgba(0,0,0,0)',
        margin=dict(l=10, r=10, t=40, b=10),
        height=height,
        hovermode="y unified"
    )
    fig.update_xaxes(showgrid=True, gridcolor='#EAEDED', title_text="用電量(度) (Usage in kWh)", tickfont=axis_font, title_font=axis_font)
    fig.update_yaxes(tickfont=axis_font, title_font=axis_font)
    
    if not grouped.empty:
        max_val = grouped[value_col].max()
        fig.update_xaxes(range=[0, max_val * 1.35])
        
    return fig

def plot_yoy_line_custom(df, year_curr, year_prev, group_col=None, group_val=None, metric="usage"):
    if metric == "cost":
        col_name = '電費金額(元)'; unit_label = "元 (NTD)"
        color_curr = COLORS['chart_cost_green']; title_suffix = "電費 (Cost)"
    else:
        col_name = '用電量(度數)'; unit_label = "度 (kWh)"
        color_curr = COLORS['chart_usage_blue']; title_suffix = "用電量 (Usage)"

    df_c = df[df['統計年度'] == year_curr]
    df_p = df[df['統計年度'] == year_prev]
    
    if group_col and group_val and group_val != "全部":
        df_c = df_c[df_c[group_col] == group_val]
        df_p = df_p[df_p[group_col] == group_val]
        
    df_c_g = df_c.groupby('統計月份')[col_name].sum().reindex(range(1,13), fill_value=0)
    df_p_g = df_p.groupby('統計月份')[col_name].sum().reindex(range(1,13), fill_value=0)
    
    fig = go.Figure()
    
    # 🔥 加入 <b> 標籤使圖例文字加粗
    fig.add_trace(go.Scatter(
        x=list(range(1,13)), y=df_p_g, name=f"<b>{year_prev}年</b>",
        mode='lines+markers', 
        line=dict(color=COLORS['line_prev_year'], width=3, dash='solid'),
        marker=dict(size=8, opacity=0.8),
    ))
    
    # 🔥 加入 <b> 標籤使圖例文字加粗
    fig.add_trace(go.Scatter(
        x=list(range(1,13)), y=df_c_g, name=f"<b>{year_curr}年</b>",
        mode='lines+markers+text',
        line=dict(color=color_curr, width=5),
        marker=dict(size=12),
        text=df_c_g.apply(lambda x: f"{int(x):,}" if x>0 else ""),
        textposition='top center', 
        textfont=dict(color='black', size=16)
    ))
    
    axis_font = dict(size=18, color=COLORS['text_axis'])
    
    fig.update_layout(
        title=dict(text=f"{title_suffix}", font=dict(size=20, color=COLORS['text_main'])),
        plot_bgcolor='rgba(0,0,0,0)',
        # 🔥 修改 legend 設定：移除背景色、放大字體
        legend=dict(
            orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5, 
            font=dict(size=20), 
            bgcolor="rgba(0,0,0,0)" 
        ),
        margin=dict(t=120, b=20, l=20, r=20),
        height=500,
        hovermode="x unified"
    )
    fig.update_xaxes(tickmode='linear', dtick=1, title_text="月份 (Month)", tickfont=axis_font, title_font=axis_font)
    fig.update_yaxes(title_text=f"{title_suffix} ({unit_label})", gridcolor='#EAEDED', tickfont=axis_font, title_font=axis_font)
    
    if not df_c_g.empty and not df_p_g.empty:
        max_v = max(df_c_g.max(), df_p_g.max())
        fig.update_yaxes(range=[0, max_v * 1.3])

    return fig

# =========================================================
# 4. 分頁 Fragment: 電力設備管理 (Tab 1)
# =========================================================
@st.fragment
def render_equipment_management_fragment(df, years, default_idx):
    st.markdown("<br>", unsafe_allow_html=True)
    
    # 動態計算符合全域狀態的 index
    current_idx = years.index(st.session_state.global_selected_year) if st.session_state.global_selected_year in years else default_idx
    
    selected_year = st.selectbox(
        "請選擇統計年度", 
        years, 
        index=current_idx, 
        key="tab1_year_select",
        on_change=lambda: st.session_state.update({"global_selected_year": st.session_state.tab1_year_select})
    )
    
    df_curr = df[df['統計年度'] == selected_year]
    df_prev = df[df['統計年度'] == (selected_year - 1)]
    time_str = get_year_range_string(df, selected_year)
    
    # 1. 資訊統計
    st.markdown(f"<div class='section-title'>📌 {time_str} 各電力設備資訊統計 (Equipment Information Statistics)</div>", unsafe_allow_html=True)
    
    campuses = ["蘭潭校區", "民雄校區", "新民校區", "林森校區"]
    
    for campus in campuses:
        df_c = df_curr[df_curr['校區'] == campus]
        if df_c.empty: continue
        
        with st.expander(f"🏫 {campus}"):
            grid = st.columns(2)
            card_idx = 0
            df_c_sorted = sort_addresses(df_c[['用電地址', '電號']].drop_duplicates(), campus)
            
            for _, row in df_c_sorted.iterrows():
                addr = row['用電地址']
                mid = row['電號']
                m_data = df_c[df_c['電號'] == mid]
                
                val_kwh = m_data['用電量(度數)'].sum()
                val_cost = m_data['電費金額(元)'].sum()
                val_co2 = m_data['碳排放量'].sum()
                
                with grid[card_idx % 2]:
                    render_kpi_card_custom(
                        title=f"{addr} ({mid})",
                        col1=("⚡ 用電 (Usage)", f"{val_kwh:,.0f}", "度 (kWh)"),
                        col2=("💰 電費 (Cost)", f"{val_cost:,.0f}", "元 (NTD)"),
                        col3=("🌍 碳排放量 (Carbon Emissions)", f"{val_co2:,.4f}", "公噸CO<sub>2</sub>e (tCO<sub>2</sub>e)"),
                        color_code=COLORS['campus'].get(campus, "#EAEDED"),
                        ratios=[3,3,4],
                        font_class="sm"
                    )
                card_idx += 1
            
    st.markdown("---")
    
    # 2. 逐月統計
    st.markdown(f"<div class='section-title'>📊 {selected_year}年 電力設備逐月用電統計 (Monthly Electricity Usage Statistics)</div>", unsafe_allow_html=True)
    
    sel_campus = st.radio("選擇校區", campuses, horizontal=True, label_visibility="collapsed", key="tab1_campus_radio")
    
    df_target_c = df_curr[df_curr['校區'] == sel_campus]
    sorted_addrs = sort_addresses(pd.DataFrame({'用電地址': df_target_c['用電地址'].unique()}), sel_campus)['用電地址'].tolist()
    addr_list = [f"{sel_campus}總計"] + sorted_addrs
    
    sel_addr = st.selectbox("選擇統計範圍", addr_list, label_visibility="collapsed", key="tab1_addr_select")
    
    if "總計" in sel_addr:
        chart_data = df_target_c
        chart_title_prefix = f"{sel_campus}總計"
    else:
        chart_data = df_target_c[df_target_c['用電地址'] == sel_addr]
        chart_title_prefix = f"{sel_addr}"
        
    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="chart-box">', unsafe_allow_html=True)
        fig_u = plot_single_bar_chart(chart_data, '統計月份', '用電量(度數)', f"{chart_title_prefix} 用電量 (Electricity Usage)", COLORS['chart_usage_blue'], "度 (kWh)")
        st.plotly_chart(fig_u, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    with c2:
        st.markdown('<div class="chart-box">', unsafe_allow_html=True)
        fig_c = plot_single_bar_chart(chart_data, '統計月份', '電費金額(元)', f"{chart_title_prefix} 電費金額 (Electricity Cost)", COLORS['chart_cost_green'], "元 (NTD)")
        st.plotly_chart(fig_c, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown("---")
    
    # 3. 近2年比較
    months_curr = sorted(df_curr['統計月份'].unique())
    month_range = f"{months_curr[0]}~{months_curr[-1]}月" if months_curr else ""
    compare_year_str = f"{selected_year-1} vs {selected_year}"
    
    st.markdown(f"<div class='section-title'>📉 近2年({compare_year_str}) {month_range} 電力使用比較 (2-Year Electricity Usage Comparison)</div>", unsafe_allow_html=True)
    
    for campus in campuses:
        meters_c = df[df['校區'] == campus][['用電地址', '電號']].drop_duplicates()
        if meters_c.empty: continue
        
        with st.expander(f"🏫 {campus}"):
            grid_yoy = st.columns(2)
            yoy_idx = 0
            meters_c_sorted = sort_addresses(meters_c, campus)
            
            for _, row in meters_c_sorted.iterrows():
                addr = row['用電地址']
                mid = row['電號']
                
                curr_months = df_curr[(df_curr['電號'] == mid)]['統計月份'].unique()
                kwh_curr = df_curr[(df_curr['電號'] == mid)]['用電量(度數)'].sum()
                co2_curr = df_curr[(df_curr['電號'] == mid)]['碳排放量'].sum()
                
                df_prev_match = df_prev[(df_prev['電號'] == mid) & (df_prev['統計月份'].isin(curr_months))]
                kwh_prev_period = df_prev_match['用電量(度數)'].sum()
                co2_prev_period = df_prev_match['碳排放量'].sum()
                
                if kwh_prev_period > 0:
                    rate = (kwh_prev_period - kwh_curr) / kwh_prev_period * 100
                    diff_kwh = kwh_curr - kwh_prev_period
                    diff_co2 = co2_curr - co2_prev_period
                    txt_rate = f"{rate:.1f}%"
                    txt_kwh = f"{diff_kwh:+,.0f}"
                    txt_co2 = f"{diff_co2:+.4f}"
                else:
                    txt_rate = "-"
                    txt_kwh = "-"
                    txt_co2 = "-"
                
                with grid_yoy[yoy_idx % 2]:
                    render_kpi_card_custom(
                        title=f"{addr} ({mid})",
                        col1=("節能率 (Saving Rate)", txt_rate, ""),
                        col2=("用電增減 (Usage Diff)", txt_kwh, "度 (kWh)"),
                        col3=("碳排增減 (CO<sub>2</sub>e Diff)", txt_co2, "公噸 (tCO<sub>2</sub>e)"),
                        color_code=COLORS['campus'].get(campus),
                        ratios=[3,3,4],
                        font_class="sm"
                    )
                yoy_idx += 1
            
    st.markdown(f"<div class='section-subtitle'>📅 近2年({compare_year_str}) {month_range} 用電地址逐月電力使用資訊比較 (2-Year Monthly Comparison by Address)</div>", unsafe_allow_html=True)
    
    sel_campus_yoy = st.radio("校區", campuses, horizontal=True, label_visibility="collapsed", key="tab1_yoy_campus_radio")
    sorted_addrs_yoy = sort_addresses(pd.DataFrame({'用電地址': df[df['校區'] == sel_campus_yoy]['用電地址'].unique()}), sel_campus_yoy)['用電地址'].tolist()
    addr_opts = [f"{sel_campus_yoy}總計"] + sorted_addrs_yoy
    sel_addr_yoy = st.selectbox("範圍", addr_opts, label_visibility="collapsed", key="tab1_yoy_addr_select")
        
    target_df = df.copy()
    if "總計" in sel_addr_yoy:
        target_df = target_df[target_df['校區'] == sel_campus_yoy]
    else:
        target_df = target_df[target_df['用電地址'] == sel_addr_yoy]
    
    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="chart-box">', unsafe_allow_html=True)
        fig_yoy_usage = plot_yoy_line_custom(target_df, selected_year, selected_year-1, metric="usage")
        st.plotly_chart(fig_yoy_usage, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    with c2:
        st.markdown('<div class="chart-box">', unsafe_allow_html=True)
        fig_yoy_cost = plot_yoy_line_custom(target_df, selected_year, selected_year-1, metric="cost")
        st.plotly_chart(fig_yoy_cost, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)


# =========================================================
# 5. 分頁 Fragment: 電力數據儀表板 (Tab 2)
# =========================================================
@st.fragment
def render_dashboard_fragment(df, years, default_idx):
    st.markdown("<br>", unsafe_allow_html=True)
    
    # 動態計算符合全域狀態的 index
    current_idx = years.index(st.session_state.global_selected_year) if st.session_state.global_selected_year in years else default_idx
    
    selected_year = st.selectbox(
        "請選擇統計年度", 
        years, 
        index=current_idx, 
        key="tab2_year_select",
        on_change=lambda: st.session_state.update({"global_selected_year": st.session_state.tab2_year_select})
    )
    
    df_curr = df[df['統計年度'] == selected_year]
    df_prev = df[df['統計年度'] == (selected_year - 1)]
    time_str = get_year_range_string(df, selected_year)

    # 1. 全校指標
    st.markdown(f"<div class='section-title'>📌 {time_str} 全校電力使用重要指標 (Key Electricity Usage Indicators)</div>", unsafe_allow_html=True)
    
    targets = ["全校", "蘭潭校區", "民雄校區", "新民校區", "林森校區"]
    
    for target in targets:
        if target == "全校":
            d = df_curr
            bg_color = COLORS['kpi_bg_orange']
        else:
            d = df_curr[df_curr['校區'] == target]
            bg_color = COLORS['campus'].get(target, "#CCD1D1")
            
        val_kwh = d['用電量(度數)'].sum()
        val_cost = d['電費金額(元)'].sum()
        val_co2 = d['碳排放量'].sum()
        
        render_kpi_card_custom(
            title=target,
            col1=("⚡ 用電 (Usage)", f"{val_kwh:,.0f}", "度 (kWh)"),
            col2=("💰 電費 (Cost)", f"{val_cost:,.0f}", "元 (NTD)"),
            col3=("🌍 碳排放量 (Carbon Emissions)", f"{val_co2:,.4f}", "公噸CO<sub>2</sub>e (tCO<sub>2</sub>e)"),
            color_code=bg_color,
            ratios=[1,1,1],
            font_class="lg"
        )
        
    st.markdown("---")
    
    # 2. 全校用電設備統計
    st.markdown(f"<div class='section-title'>📊 {time_str} 全校用電設備統計 (Electricity Usage by Equipment)</div>", unsafe_allow_html=True)
    
    total_usage_school = df_curr['用電量(度數)'].sum()
    
    col_rank_1, col_rank_2 = st.columns(2)
    
    with col_rank_1:
        st.markdown('<div class="chart-box">', unsafe_allow_html=True)
        st.markdown(f"<div class='section-subtitle'>🏫 依校區統計 (By Campus)</div>", unsafe_allow_html=True) 
        fig_rank_campus = plot_horizontal_ranking(
            df_curr, '校區', '用電量(度數)', total_usage_school, 
            "", height=500
        )
        st.plotly_chart(fig_rank_campus, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
        
    with col_rank_2:
        st.markdown('<div class="chart-box">', unsafe_allow_html=True)
        st.markdown(f"<div class='section-subtitle'>🔌 依用電設備統計 (By Equipment)</div>", unsafe_allow_html=True)
        addr_count = df_curr['用電地址'].nunique()
        dynamic_height = max(500, addr_count * 45) 
        
        fig_rank_addr = plot_horizontal_ranking(
            df_curr, '用電地址', '用電量(度數)', total_usage_school, 
            "", height=dynamic_height
        )
        st.plotly_chart(fig_rank_addr, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown("---")
    
    # 3. 年度校區用電趨勢
    st.markdown(f"<div class='section-title'>📈 {time_str} 年度校區用電趨勢 (Annual Electricity Usage Trend)</div>", unsafe_allow_html=True)
    
    sel_scope = st.radio("選擇統計範圍", targets, horizontal=True, key="tab2_scope_radio")
    
    if sel_scope == "全校":
        d_scope = df_curr
        chart_title_prefix = "全校"
    else:
        d_scope = df_curr[df_curr['校區'] == sel_scope]
        chart_title_prefix = sel_scope
    
    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="chart-box">', unsafe_allow_html=True)
        fig_u = plot_single_bar_chart(d_scope, '統計月份', '用電量(度數)', f"{chart_title_prefix} 用電量 (Electricity Usage)", COLORS['chart_usage_blue'], "度 (kWh)")
        st.plotly_chart(fig_u, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    with c2:
        st.markdown('<div class="chart-box">', unsafe_allow_html=True)
        fig_c = plot_single_bar_chart(d_scope, '統計月份', '電費金額(元)', f"{chart_title_prefix} 電費金額 (Electricity Cost)", COLORS['chart_cost_green'], "元 (NTD)")
        st.plotly_chart(fig_c, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown("---")
    
    # 4. 節能成效
    months_curr = sorted(df_curr['統計月份'].unique())
    month_range = f"{months_curr[0]}~{months_curr[-1]}月" if months_curr else ""
    compare_year_str = f"{selected_year-1} vs {selected_year}"
    
    st.markdown(f"<div class='section-title'>📉 近2年({compare_year_str}) {month_range} 節能成效總覽<br><span style='font-size: 1.5rem; color: #5D6D7E; font-weight: 600;'>(2-Year Energy Saving Performance Overview)</span></div>", unsafe_allow_html=True)
    
    # =========================================================
    # 新增：全校節能成效資訊卡 (置頂綜整)
    # =========================================================
    curr_months_all = df_curr['統計月份'].unique()
    kwh_curr_all = df_curr['用電量(度數)'].sum()
    co2_curr_all = df_curr['碳排放量'].sum()
    
    df_prev_match_all = df_prev[df_prev['統計月份'].isin(curr_months_all)]
    kwh_prev_all = df_prev_match_all['用電量(度數)'].sum()
    co2_prev_all = df_prev_match_all['碳排放量'].sum()
    
    if kwh_prev_all > 0:
        rate_all = (kwh_prev_all - kwh_curr_all) / kwh_prev_all * 100
        diff_kwh_all = kwh_curr_all - kwh_prev_all
        diff_co2_all = co2_curr_all - co2_prev_all
        txt_rate_all = f"{rate_all:.1f}%"
        txt_kwh_all = f"{diff_kwh_all:+,.0f}"
        txt_co2_all = f"{diff_co2_all:+.4f}"
    else:
        txt_rate_all = "-"
        txt_kwh_all = "-"
        txt_co2_all = "-"
        
    render_kpi_card_custom(
        title="全校",
        col1=("節能率 (Saving Rate)", txt_rate_all, ""),
        col2=("用電增減 (Usage Diff)", txt_kwh_all, "度 (kWh)"),
        col3=("碳排增減 (CO<sub>2</sub>e Diff)", txt_co2_all, "公噸 (tCO<sub>2</sub>e)"),
        color_code=COLORS['kpi_bg_orange'],
        ratios=[2,4,4],
        font_class="md"
    )
    # =========================================================

    col_kpi_yoy = st.columns(2)
    campuses = ["蘭潭校區", "民雄校區", "新民校區", "林森校區"]
    
    for i, c in enumerate(campuses):
        curr_months = df_curr[df_curr['校區'] == c]['統計月份'].unique()
        kwh_curr = df_curr[df_curr['校區'] == c]['用電量(度數)'].sum()
        co2_curr = df_curr[df_curr['校區'] == c]['碳排放量'].sum()
        
        df_prev_match = df_prev[(df_prev['校區'] == c) & (df_prev['統計月份'].isin(curr_months))]
        kwh_prev_p = df_prev_match['用電量(度數)'].sum()
        co2_prev_p = df_prev_match['碳排放量'].sum()
        
        if kwh_prev_p > 0:
            rate = (kwh_prev_p - kwh_curr) / kwh_prev_p * 100
            diff_kwh = kwh_curr - kwh_prev_p
            diff_co2 = co2_curr - co2_prev_p
            txt_rate = f"{rate:.1f}%"
            txt_kwh = f"{diff_kwh:+,.0f}"
            txt_co2 = f"{diff_co2:+.4f}"
        else:
            txt_rate = "-"
            txt_kwh = "-"
            txt_co2 = "-"
            
        with col_kpi_yoy[i % 2]:
            render_kpi_card_custom(
                title=f"{c}",
                col1=("節能率 (Saving Rate)", txt_rate, ""),
                col2=("用電增減 (Usage Diff)", txt_kwh, "度 (kWh)"),
                col3=("碳排增減 (CO<sub>2</sub>e Diff)", txt_co2, "公噸 (tCO<sub>2</sub>e)"),
                color_code=COLORS['campus'].get(c),
                ratios=[2,4,4],
                font_class="md"
            )

    st.markdown(f"<div class='section-subtitle'>📅 近2年({compare_year_str}) {month_range} 校區電力使用逐月比較 (2-Year Monthly Comparison)</div>", unsafe_allow_html=True)
    
    sel_yoy_scope = st.radio("範圍", targets, horizontal=True, label_visibility="collapsed", key="tab2_yoy_scope_radio")
    
    target_d = df.copy()
    if sel_yoy_scope != "全校":
        target_d = target_d[target_d['校區'] == sel_yoy_scope]
    
    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="chart-box">', unsafe_allow_html=True)
        fig_dash_yoy_u = plot_yoy_line_custom(target_d, selected_year, selected_year-1, metric="usage")
        st.plotly_chart(fig_dash_yoy_u, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    with c2:
        st.markdown('<div class="chart-box">', unsafe_allow_html=True)
        fig_dash_yoy_c = plot_yoy_line_custom(target_d, selected_year, selected_year-1, metric="cost")
        st.plotly_chart(fig_dash_yoy_c, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

# =========================================================
# 6. 分頁 Fragment: 資料異動及下載 (Tab 3)
# =========================================================
@st.fragment
def render_data_management_fragment(df_raw, df_processed, years, default_idx):
    st.markdown("<br>", unsafe_allow_html=True)
    
    # 動態計算符合全域狀態的 index
    current_idx = years.index(st.session_state.global_selected_year) if st.session_state.global_selected_year in years else default_idx
    
    selected_year = st.selectbox(
        "請選擇資料年度", 
        years, 
        index=current_idx, 
        key="tab3_year_select",
        on_change=lambda: st.session_state.update({"global_selected_year": st.session_state.tab3_year_select})
    )
    st.markdown("---")
    
    # 1. 下載區塊
    st.markdown("### 📥 年度資料下載")
    st.write(f"下載 {selected_year} 年度各用電地址統計資料 (統計基準：使用期間)")
    
    df_download_src = df_processed[df_processed['統計年度'] == selected_year]
    
    if not df_download_src.empty:
        df_download = df_download_src.groupby(['統計年度', '校區', '用電地址', '電號'], as_index=False).agg({
            '用電量(度數)': 'sum',
            '電費金額(元)': 'sum'
        })
        
        df_download.rename(columns={
            '統計年度': '資料統計年度',
            '用電量(度數)': '年度用電量(度)',
            '電費金額(元)': '年度電費金額(元)'
        }, inplace=True)
        
        cols_order = ['資料統計年度', '校區', '用電地址', '電號', '年度用電量(度)', '年度電費金額(元)']
        df_download = df_download[cols_order]
        
        csv = df_download.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label=f"📥 下載 {selected_year} 年統計報表 (CSV)",
            data=csv,
            file_name=f"{selected_year}_electricity_stats.csv",
            mime='text/csv',
        )
    else:
        st.info("該年度尚無統計資料")
        
    st.markdown("---")
    
    # 2. 編輯區塊
    st.markdown("### ✏️ 填報資料異動")
    st.info("💡 說明：此處針對「原始帳單資料」進行修改。系統將保留其他年度及本年度未修改的資料。")
    
    df_edit_src = df_raw[df_raw['帳單年度'] == selected_year]
    
    if not df_edit_src.empty:
        edited_df = st.data_editor(
            df_edit_src, 
            key="data_editor_tab3",
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True
        )
        
        if st.button("💾 儲存異動", type="primary"):
            with st.spinner("正在寫入資料庫，請稍候..."):
                df_other_years = df_raw[df_raw['帳單年度'] != selected_year]
                df_final_to_save = pd.concat([df_other_years, edited_df], ignore_index=True)
                df_final_to_save.sort_values(by=['帳單年度', '帳單月份', '校區'], inplace=True, ascending=[False, False, True])
                
                if update_google_sheet_data(df_final_to_save):
                    st.success("✅ 資料異動已成功儲存！")
                    st.rerun()
                else:
                    st.error("❌ 儲存失敗，請檢查網路連線。")
    else:
        st.info("該年度尚無填報紀錄可供編輯。")

# =========================================================
# 7. 主程式
# =========================================================
def main():
    st.markdown('<div style="font-size: 2.4rem; font-weight: 900; color: #2C3E50; margin-bottom: 20px;">⚡ 全校電力管理平台 (University Electricity Management Platform)</div>', unsafe_allow_html=True)
    
    gc = init_google_sheet()
    if not gc: return
    
    with st.spinner("資料讀取與精密計算中 (按日攤算)..."):
        df_processed, df_raw = load_and_process_data(gc)
    
    if df_processed.empty:
        st.warning("目前尚無資料，請先至「填報系統」輸入數據。")
        return

    tab1, tab2, tab3 = st.tabs(["🏗️ 電力設備管理", "📊 電力數據儀表板", "📝 資料異動及下載"])
    
    years = sorted(df_processed['統計年度'].unique(), reverse=True)
    current_year = datetime.datetime.now().year
    default_idx = years.index(current_year) if current_year in years else 0
    
    # 🟢 新增：初始化全域年度狀態 (st.session_state)
    if 'global_selected_year' not in st.session_state:
        st.session_state.global_selected_year = years[default_idx] if years else None
    
    with tab1:
        render_equipment_management_fragment(df_processed, years, default_idx)
        
    with tab2:
        render_dashboard_fragment(df_processed, years, default_idx)
        
    with tab3:
        render_data_management_fragment(df_raw, df_processed, years, default_idx)

if __name__ == "__main__":
    main()