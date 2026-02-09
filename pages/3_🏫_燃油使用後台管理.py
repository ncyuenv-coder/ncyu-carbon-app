import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import gspread
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import streamlit_authenticator as stauth
import plotly.express as px
import plotly.graph_objects as go
import time
import re
import os

# ==========================================
# 0. 系統設定
# ==========================================
st.set_page_config(page_title="燃油後台管理", page_icon="🏫", layout="wide")

def get_taiwan_time():
    return datetime.utcnow() + timedelta(hours=8)

# ==========================================
# 1. CSS 樣式表 (V164.0 原版)
# ==========================================
st.markdown("""
<style>
    /* --- 全域設定 --- */
    :root {
        color-scheme: light;
        --orange-bg: #E67E22;     
        --orange-dark: #D35400;
        --text-main: #2C3E50;
        --text-sub: #566573;
        --morandi-red: #C0392B; 
        --kpi-gas: #52BE80;
        --kpi-diesel: #F4D03F;
        --kpi-total: #5DADE2;
        --kpi-co2: #AF7AC5;
        --morandi-blue: #34495E;
        --deep-gray: #333333;
    }

    [data-testid="stAppViewContainer"] { background-color: #EAEDED; color: var(--text-main); }
    [data-testid="stHeader"] { background-color: rgba(0,0,0,0); }
    [data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 1px solid #BDC3C7; }

    /* 表格字體放大 */
    [data-testid="stDataFrame"] { font-size: 1.25rem !important; }
    [data-testid="stDataFrame"] div { font-size: 1.25rem !important; }

    /* --- 設備詳細卡片樣式 --- */
    .dev-card-v148 {
        background-color: #FFFFFF; border: 1px solid #BDC3C7; border-radius: 12px; overflow: hidden;
        box-shadow: 0 3px 6px rgba(0,0,0,0.08); margin-bottom: 20px; display: flex; flex-direction: column;
    }
    .dev-header {
        padding: 12px 15px; display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid rgba(0,0,0,0.1); 
    }
    .dev-header-left { display: flex; flex-direction: column; gap: 3px; }
    .dev-id { font-size: 1.15rem; font-weight: 800; color: #000000 !important; opacity: 0.8; } 
    .dev-name-row { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
    .dev-name { font-size: 1.25rem; font-weight: 900; color: #000000 !important; }
    .qty-badge {
        font-size: 0.85rem; background-color: rgba(255,255,255,0.9); padding: 2px 8px;
        border-radius: 12px; color: #2C3E50; font-weight: bold; border: 1px solid rgba(0,0,0,0.1);
    }
    
    .dev-header-right { text-align: right; display: flex; flex-direction: column; align-items: flex-end; justify-content: center; }
    .dev-vol { font-size: 1.8rem; color: #C0392B !important; font-weight: 900; line-height: 1.1; text-shadow: none !important; }
    .dev-unit { font-size: 0.95rem; color: var(--deep-gray); font-weight: bold; margin-left: 2px; }

    .dev-body {
        padding: 15px; font-size: 0.95rem; color: var(--deep-gray);
        display: grid; grid-template-columns: 1fr 1fr; gap: 8px; 
    }
    .dev-item { margin-bottom: 2px; display: flex; align-items: baseline; }
    .dev-label { font-weight: 700; color: var(--deep-gray) !important; font-size: 0.95rem; margin-right: 5px; min-width: 80px; }
    .dev-val { color: var(--deep-gray) !important; font-weight: 600; font-size: 1rem; }
    
    .dev-footer {
        padding: 10px 15px; background-color: #F8F9F9; border-top: 1px solid #E5E7E9;
        display: flex; justify-content: space-between; align-items: center;
    }
    .dev-count { font-weight: 700; color: #34495E; font-size: 0.95rem; }
    .alert-status { color: #C0392B; font-weight: 900; display: flex; align-items: center; gap: 5px; background-color: #FADBD8; padding: 4px 12px; border-radius: 12px; font-size: 0.9rem; }

    /* 後台 - 統計卡片 */
    .stat-card-v119 { background-color: #FFFFFF; border: 1px solid #BDC3C7; border-radius: 12px; overflow: hidden; box-shadow: 0 3px 6px rgba(0,0,0,0.08); margin-bottom: 15px; height: 100%; }
    .stat-header { padding: 15px 25px; display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid rgba(0,0,0,0.1); }
    .stat-title { font-size: 1.4rem; font-weight: 900; color: #2C3E50; } 
    .stat-count { font-size: 2.2rem; font-weight: 900; color: var(--morandi-red); }
    .stat-body-split { padding: 15px 20px; display: flex; }
    .stat-col-left { width: 50%; padding-right: 15px; border-right: 1px dashed #BDC3C7; }
    .stat-col-right { width: 50%; padding-left: 15px; }
    .stat-item { font-size: 1.1rem; color: #566573; margin-bottom: 8px; display: flex; justify-content: space-between; } 
    .stat-item-label { font-weight: bold; color: #2C3E50; }
    .stat-item-val { color: #2C3E50; font-weight: 900; }

    /* 後台 - Top KPI */
    .top-kpi-card { background-color: #FFFFFF; border-radius: 12px; padding: 25px; text-align: center; box-shadow: 0 4px 10px rgba(0,0,0,0.1); border: 1px solid #BDC3C7; margin-bottom: 10px; }
    .top-kpi-title { font-size: 1.15rem; color: #7F8C8D; font-weight: bold; margin-bottom: 5px; }
    .top-kpi-value { font-size: 3.5rem; color: #2C3E50; font-weight: 900; line-height: 1.1; }

    /* Admin 儀表板 KPI */
    .admin-kpi-card { background-color: #FFFFFF; border: 1px solid #BDC3C7; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.1); height: 100%; text-align: center; margin-bottom: 20px; }
    .admin-kpi-header { padding: 10px; font-size: 1.2rem; font-weight: bold; color: #2C3E50; border-bottom: 1px solid rgba(0,0,0,0.1); }
    .admin-kpi-body { padding: 20px; }
    .admin-kpi-value { font-size: 2.8rem; font-weight: 900; color: #2C3E50; margin-bottom: 5px; }
    .admin-kpi-unit { font-size: 1rem; color: #7F8C8D; font-weight: normal; margin-left: 5px; }
    .admin-kpi-sub { font-size: 0.9rem; display: inline-block; padding: 2px 10px; border-radius: 15px; background-color: #F9E79F; color: #7D6608; margin-top: 5px; font-weight: bold; }

    /* 其他 */
    .dashboard-main-title { font-size: 1.8rem; font-weight: 900; text-align: center; color: #2C3E50; margin-bottom: 20px; background-color: #F8F9F9; padding: 10px; border-radius: 10px; border: 1px solid #BDC3C7; }
    .morandi-header { background-color: #EBF5FB; color: #2E4053; padding: 15px; border-radius: 8px; border-left: 8px solid #5499C7; font-size: 1.35rem; font-weight: 700; margin-top: 25px; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. 身份驗證與初始化
# ==========================================
def clean_secrets(obj):
    if isinstance(obj, dict) or "AttrDict" in str(type(obj)): return {k: clean_secrets(v) for k, v in obj.items()}
    elif isinstance(obj, list): return [clean_secrets(i) for i in obj]
    return obj

if st.session_state.get("authentication_status") is not True:
    st.warning("🔒 請先至首頁 (Hello) 登入系統")
    st.stop()

username = st.session_state.get("username")
name = st.session_state.get("name")

# 🔴 管理員權限檢查
if username != 'admin':
    st.error("⛔ 您沒有權限訪問此頁面")
    st.stop()

with st.sidebar:
    st.header(f"👑 {name}")
    st.caption("管理員後台模式")
    st.success("☁️ 雲端連線正常")
    st.markdown("---")
    authenticator.logout('登出系統', 'sidebar')

# ==========================================
# 3. 資料庫連線與設定
# ==========================================
SHEET_ID = "1gqDU21YJeBoBOd8rMYzwwZ45offXWPGEODKTF6B8k-Y" 
DRIVE_FOLDER_ID = "1Uryuk3-9FHJ39w5Uo8FYxuh9VOFndeqD"
DEVICE_ORDER = ["公務車輛(GV-1-)", "乘坐式割草機(GV-2-)", "乘坐式農用機具(GV-3-)", "鍋爐(GS-1-)", "發電機(GS-2-)", "肩背或手持式割草機、吹葉機(GS-3-)", "肩背或手持式農用機具(GS-4-)"]
DEVICE_CODE_MAP = {"GV-1": "公務車輛(GV-1-)", "GV-2": "乘坐式割草機(GV-2-)", "GV-3": "乘坐式農用機具(GV-3-)", "GS-1": "鍋爐(GS-1-)", "GS-2": "發電機(GS-2-)", "GS-3": "肩背或手持式割草機、吹葉機(GS-3-)", "GS-4": "肩背或手持式農用機具(GS-4-)"}
MORANDI_COLORS = { "公務車輛(GV-1-)": "#B0C4DE", "乘坐式割草機(GV-2-)": "#F5CBA7", "乘坐式農用機具(GV-3-)": "#D7BDE2", "鍋爐(GS-1-)": "#E6B0AA", "發電機(GS-2-)": "#A9CCE3", "肩背或手持式割草機、吹葉機(GS-3-)": "#A3E4D7", "肩背或手持式農用機具(GS-4-)": "#F9E79F" }
DASH_PALETTE = ['#B0C4DE', '#F5CBA7', '#A9CCE3', '#E6B0AA', '#D7BDE2', '#A3E4D7', '#F9E79F', '#95A5A6', '#85C1E9', '#D2B4DE', '#F1948A', '#76D7C4']
UNREPORTED_COLORS = ["#D5DBDB", "#FAD7A0", "#D2B4DE", "#AED6F1", "#A3E4D7", "#F5B7B1"]

@st.cache_resource
def init_google_fuel():
    oauth = st.secrets["gcp_oauth"]
    creds = Credentials(token=None, refresh_token=oauth["refresh_token"], token_uri="https://oauth2.googleapis.com/token", client_id=oauth["client_id"], client_secret=oauth["client_secret"], scopes=["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/spreadsheets"])
    gc = gspread.authorize(creds); drive = build('drive', 'v3', credentials=creds)
    return gc, drive

try:
    gc, drive_service = init_google_fuel()
    sh = gc.open_by_key(SHEET_ID)
    try: ws_equip = sh.worksheet("設備清單") 
    except: ws_equip = sh.sheet1 
    try: ws_record = sh.worksheet("油料填報紀錄")
    except:
        try: ws_record = sh.worksheet("填報紀錄")
        except: ws_record = sh.add_worksheet(title="油料填報紀錄", rows="1000", cols="13")
except Exception as e: st.error(f"燃油資料庫連線失敗: {e}"); st.stop()

@st.cache_data(ttl=600)
def load_fuel_data():
    max_retries = 3; delay = 2; df_e = pd.DataFrame(); df_r = pd.DataFrame()
    for attempt in range(max_retries):
        try: df_e = pd.DataFrame(ws_equip.get_all_records()).astype(str); break
        except Exception as e:
            if "429" in str(e) and attempt < max_retries - 1: time.sleep(delay); delay *= 2
            else: raise e
    if '設備編號' in df_e.columns: df_e['統計類別'] = df_e['設備編號'].apply(lambda c: next((v for k, v in DEVICE_CODE_MAP.items() if str(c).startswith(k)), "其他/未分類"))
    else: df_e['統計類別'] = "未設定I欄"
    df_e['設備數量_num'] = pd.to_numeric(df_e['設備數量'], errors='coerce').fillna(1)
    
    for attempt in range(max_retries):
        try: data = ws_record.get_all_values(); df_r = pd.DataFrame(data[1:], columns=data[0]) if len(data) > 1 else pd.DataFrame(columns=data[0]); break
        except Exception as e:
            if "429" in str(e) and attempt < max_retries - 1: time.sleep(delay); delay *= 2
            else: raise e
    return df_e, df_r

df_equip, df_records = load_fuel_data()

# ==========================================
# 4. 後台功能函式 (原汁原味移植)
# ==========================================

def render_admin_overview():
    st.markdown('<div class="morandi-header">🏢 全校燃油設備總覽</div>', unsafe_allow_html=True)
    
    # 1. 總量 KPI
    total_cnt = int(pd.to_numeric(df_equip['設備數量'], errors='coerce').fillna(1).sum())
    gas_cnt = int(pd.to_numeric(df_equip[df_equip['原燃物料名稱'].str.contains('汽油')]['設備數量'], errors='coerce').fillna(1).sum())
    diesel_cnt = int(pd.to_numeric(df_equip[df_equip['原燃物料名稱'].str.contains('柴油')]['設備數量'], errors='coerce').fillna(1).sum())
    
    k1, k2, k3 = st.columns(3)
    k1.markdown(f'<div class="admin-kpi-card"><div class="admin-kpi-header">全校列管設備總數</div><div class="admin-kpi-body"><div class="admin-kpi-value">{total_cnt}</div><span class="admin-kpi-unit">台/部</span></div></div>', unsafe_allow_html=True)
    k2.markdown(f'<div class="admin-kpi-card"><div class="admin-kpi-header">汽油設備</div><div class="admin-kpi-body"><div class="admin-kpi-value" style="color:#27AE60;">{gas_cnt}</div><span class="admin-kpi-unit">台/部</span></div></div>', unsafe_allow_html=True)
    k3.markdown(f'<div class="admin-kpi-card"><div class="admin-kpi-header">柴油設備</div><div class="admin-kpi-body"><div class="admin-kpi-value" style="color:#D35400;">{diesel_cnt}</div><span class="admin-kpi-unit">台/部</span></div></div>', unsafe_allow_html=True)

    # 2. 統計卡片
    c_stat1, c_stat2 = st.columns(2)
    with c_stat1:
        st.subheader("📊 設備類別統計")
        df_cat = df_equip.groupby('統計類別')['設備數量_num'].sum().reset_index()
        fig_cat = px.bar(df_cat, x='統計類別', y='設備數量_num', color='統計類別', color_discrete_map=MORANDI_COLORS, text='設備數量_num')
        st.plotly_chart(fig_cat, use_container_width=True)

    with c_stat2:
        st.subheader("🚙 油品設備用油量佔比分析")
        df_fuel = df_equip.groupby('原燃物料名稱')['設備數量_num'].sum().reset_index()
        fig_fuel = px.pie(df_fuel, values='設備數量_num', names='原燃物料名稱', hole=0.4)
        st.plotly_chart(fig_fuel, use_container_width=True)

    st.markdown("---")
    st.subheader("📋 設備詳細清單")
    st.dataframe(df_equip, use_container_width=True)

def render_admin_dashboard():
    st.markdown('<div class="morandi-header">📈 全校油料使用儀表板</div>', unsafe_allow_html=True)
    if df_records.empty: st.warning("無資料"); return

    # 1. Cleaning
    df = df_records.copy()
    df['加油量'] = pd.to_numeric(df['加油量'], errors='coerce').fillna(0)
    df['加油日期'] = pd.to_datetime(df['加油日期'], errors='coerce')
    df = df.dropna(subset=['加油日期'])
    df['年'] = df['加油日期'].dt.year
    df['月'] = df['加油日期'].dt.month
    
    years = sorted(df['年'].unique(), reverse=True)
    sel_year = st.selectbox("📅 選擇年度", years, key="admin_dash_year")
    df_year = df[df['年'] == sel_year]

    # 2. KPI
    total_vol = df_year['加油量'].sum()
    gas_vol = df_year[df_year['原燃物料名稱'].str.contains("汽油")]['加油量'].sum()
    diesel_vol = df_year[df_year['原燃物料名稱'].str.contains("柴油")]['加油量'].sum()
    
    k1, k2, k3 = st.columns(3)
    k1.markdown(f'<div class="top-kpi-card"><div class="top-kpi-title">年度總加油量</div><div class="top-kpi-value">{total_vol:.0f}</div></div>', unsafe_allow_html=True)
    k2.markdown(f'<div class="top-kpi-card"><div class="top-kpi-title">汽油總量</div><div class="top-kpi-value" style="color:#27AE60;">{gas_vol:.0f}</div></div>', unsafe_allow_html=True)
    k3.markdown(f'<div class="top-kpi-card"><div class="top-kpi-title">柴油總量</div><div class="top-kpi-value" style="color:#D35400;">{diesel_vol:.0f}</div></div>', unsafe_allow_html=True)

    # 3. Charts
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("📊 逐月加油趨勢")
        df_trend = df_year.groupby(['月', '原燃物料名稱'])['加油量'].sum().reset_index()
        fig_trend = px.line(df_trend, x='月', y='加油量', color='原燃物料名稱', markers=True, text='加油量')
        fig_trend.update_layout(xaxis=dict(tickmode='linear', dtick=1))
        st.plotly_chart(fig_trend, use_container_width=True)
    
    with c2:
        st.subheader("🏢 全校加油量單位佔比")
        if not df_year.empty:
            df_unit = df_year.groupby('填報單位')['加油量'].sum().reset_index().sort_values('加油量', ascending=True)
            fig_unit = px.pie(df_unit, values='加油量', names='填報單位', hole=0.4, color_discrete_sequence=DASH_PALETTE)
            st.plotly_chart(fig_unit, use_container_width=True)

    st.markdown("---")
    
    c3, c4 = st.columns(2)
    with c3:
        st.subheader("⛽ 汽油/柴油 使用比例")
        df_type = df_year.groupby('原燃物料名稱')['加油量'].sum().reset_index()
        fig_pie = px.pie(df_type, values='加油量', names='原燃物料名稱', color='原燃物料名稱', 
                         color_discrete_map={'92無鉛汽油':'#ABEBC6','95無鉛汽油':'#58D68D','98無鉛汽油':'#28B463','超級柴油':'#F4D03F'})
        st.plotly_chart(fig_pie, use_container_width=True)

    with c4:
        st.subheader("🌍 碳排放量 Treemap")
        df_year['CO2e'] = df_year.apply(lambda r: r['加油量']*2.263 if '汽油' in str(r['原燃物料名稱']) else r['加油量']*2.606, axis=1) # kg
        if not df_year.empty:
            fig_tree = px.treemap(df_year, path=['填報單位', '設備名稱備註'], values='CO2e', color='填報單位', color_discrete_sequence=DASH_PALETTE)
            st.plotly_chart(fig_tree, use_container_width=True)

# ==========================================
# 5. 主程式入口
# ==========================================
menu = st.sidebar.radio("後台管理選單", ["後台設備總覽", "後台數據儀表板"])
if menu == "後台設備總覽": render_admin_overview()
elif menu == "後台數據儀表板": render_admin_dashboard()