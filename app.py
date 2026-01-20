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

# ==========================================
# 0. 系統設定
# ==========================================
st.set_page_config(page_title="國立嘉義大學碳盤查平台", page_icon="🌍", layout="wide")

# 時區校正 (GMT+8)
def get_taiwan_time():
    return datetime.utcnow() + timedelta(hours=8)

# ==========================================
# 1. CSS 樣式表 (V122: 針對新需求微調)
# ==========================================
st.markdown("""
<style>
    /* 全域設定 */
    :root {
        color-scheme: light;
        --btn-bg: #B0BEC5;        
        --btn-border: #2C3E50;    
        --btn-text: #17202A;      
        
        --orange-bg: #E67E22;     
        --orange-dark: #D35400;
        --orange-text: #FFFFFF;

        --bg-color: #EAEDED;
        --card-bg: #FFFFFF;
        --text-main: #2C3E50;
        --text-sub: #566573;
        --border-color: #BDC3C7;
        
        --morandi-red: #A93226; 
        
        /* KPI 配色 */
        --kpi-gas: #52BE80;
        --kpi-diesel: #F4D03F;
        --kpi-total: #5DADE2;
        --kpi-co2: #AF7AC5;
    }

    [data-testid="stAppViewContainer"] { background-color: var(--bg-color); color: var(--text-main); }
    [data-testid="stHeader"] { background-color: rgba(0,0,0,0); }
    [data-testid="stSidebar"] { background-color: var(--card-bg); border-right: 1px solid var(--border-color); }
    h1, h2, h3, h4, h5, h6, p, span, label, div { color: var(--text-main); }

    /* 輸入元件優化 */
    div[data-baseweb="input"] > div, div[data-baseweb="base-input"] > input, textarea, input {
        background-color: #FFFFFF !important; border-color: #BDC3C7 !important; color: #000000 !important; font-size: 1.15rem !important;
    }
    div[data-baseweb="select"] > div { border-color: #BDC3C7 !important; background-color: #FFFFFF !important; }
    ul[data-baseweb="menu"] { background-color: #FFFFFF !important; }

    /* 按鈕樣式 (橘底白字 + 置中) */
    div.stButton > button, 
    button[kind="primary"], 
    [data-testid="stFormSubmitButton"] > button {
        background-color: var(--orange-bg) !important; 
        color: #FFFFFF !important; /* 白字 */
        border: 2px solid var(--orange-dark) !important; 
        border-radius: 12px !important;
        font-size: 1.3rem !important; 
        font-weight: 800 !important; 
        padding: 0.7rem 1.5rem !important;
        box-shadow: 0 4px 6px rgba(230, 126, 34, 0.3) !important;
        display: flex; justify-content: center; align-items: center; 
        width: 100%; 
    }
    div.stButton > button p { color: #FFFFFF !important; } 
    div.stButton > button:hover, [data-testid="stFormSubmitButton"] > button:hover { 
        background-color: var(--orange-dark) !important; transform: translateY(-2px) !important; color: #FFFFFF !important;
    }

    /* V122: Tab 字體縮小 */
    button[data-baseweb="tab"] div p { font-size: 1.3rem !important; font-weight: 900 !important; color: var(--text-sub); }
    button[data-baseweb="tab"][aria-selected="true"] div p { color: #E67E22 !important; border-bottom: 3px solid #E67E22; }

    /* V122: 申報類型選單樣式優化 (針對 Radio Options) */
    /* 這裡透過 CSS 選擇器針對特定的 Radio 群組樣式 */
    .stRadio div[role="radiogroup"] label {
        background-color: #E8DAEF !important; /* 莫蘭迪紫底 */
        border: 1px solid #C39BD3 !important;
        border-radius: 8px !important;
        padding: 10px 15px !important;
        margin-right: 10px !important;
    }
    .stRadio div[role="radiogroup"] label p {
        font-size: 1.2rem !important; /* 大字 */
        font-weight: 800 !important;
        color: #5B2C6F !important;
    }

    /* 檔案上傳區 */
    [data-testid="stFileUploaderDropzone"] {
        background-color: #EBF5FB !important; border: 2px dashed #AED6F1 !important; border-radius: 12px; padding: 20px;
    }
    [data-testid="stFileUploaderDropzone"] div, [data-testid="stFileUploaderDropzone"] span, [data-testid="stFileUploaderDropzone"] small {
        color: #2E86C1 !important; 
    }

    /* V122: 深灰色說明文字 */
    .note-text-darkgray { color: #566573 !important; font-weight: bold; font-size: 1rem; margin-top: 5px; margin-bottom: 15px; }

    /* V119: 勾選框優化 */
    div[data-testid="stCheckbox"] label p {
        font-size: 1.2rem !important; color: #1F618D !important; font-weight: 900 !important;
    }

    /* 批次申報卡片 */
    .batch-card-final {
        background-color: #FFFFFF; border: 1px solid #BDC3C7; border-radius: 10px; overflow: hidden;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1); height: 100%; display: flex; flex-direction: column;
        border-left: 5px solid #E67E22;
        margin-bottom: 25px; 
    }
    .batch-header-final {
        padding: 14px 15px; font-weight: 800; color: #2C3E50; display: flex; justify-content: space-between; align-items: center;
        border-bottom: 1px solid rgba(0,0,0,0.1); font-size: 1.15rem; background-color: #F4F6F6;
    }
    .batch-qty-badge {
        font-size: 0.95rem; background-color: rgba(255,255,255,0.7); padding: 2px 10px;
        border-radius: 12px; border: 1px solid rgba(0,0,0,0.1); color: #2C3E50; font-weight: bold;
    }
    .batch-body-final {
        background-color: #FFFFFF; padding: 15px; font-size: 1rem; color: #566573; line-height: 1.6; flex-grow: 1;
    }
    .batch-row { display: flex; justify-content: space-between; margin-bottom: 5px; }
    .batch-item { flex: 1; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; margin-right: 5px; }

    /* 後台 - 統計模式專用卡片 */
    .stat-card-v119 {
        background-color: #FFFFFF; border: 1px solid #BDC3C7; border-radius: 12px; overflow: hidden;
        box-shadow: 0 3px 6px rgba(0,0,0,0.08); margin-bottom: 15px; height: 100%;
    }
    .stat-header {
        padding: 15px 25px; display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid rgba(0,0,0,0.1);
    }
    .stat-title { font-size: 1.25rem; font-weight: 800; color: #2C3E50; }
    .stat-count { font-size: 2rem; font-weight: 900; color: var(--morandi-red); }
    .stat-body-split { padding: 15px 20px; display: flex; }
    .stat-col-left { width: 50%; padding-right: 15px; border-right: 1px dashed #BDC3C7; }
    .stat-col-right { width: 50%; padding-left: 15px; }
    .stat-item { font-size: 0.95rem; color: #566573; margin-bottom: 8px; display: flex; justify-content: space-between; }
    .stat-item-label { font-weight: bold; color: #2C3E50; }
    .stat-item-val { color: #2C3E50; font-weight: 900; }

    /* 後台 - Top KPI */
    .top-kpi-card {
        background-color: #FFFFFF; border-radius: 12px; padding: 25px; text-align: center;
        box-shadow: 0 4px 10px rgba(0,0,0,0.1); border: 1px solid #BDC3C7; margin-bottom: 10px;
    }
    .top-kpi-title { font-size: 1.15rem; color: #7F8C8D; font-weight: bold; margin-bottom: 5px; }
    .top-kpi-value { font-size: 3.5rem; color: #2C3E50; font-weight: 900; line-height: 1.1; }

    /* 後台 - 儀表板 KPI (莫蘭迪黃底標) */
    .admin-kpi-card {
        background-color: #FFFFFF; border: 1px solid #BDC3C7; border-radius: 12px; overflow: hidden;
        box-shadow: 0 4px 10px rgba(0,0,0,0.1); height: 100%; text-align: center; margin-bottom: 20px;
    }
    .admin-kpi-header { padding: 10px; font-size: 1.2rem; font-weight: bold; color: #2C3E50; border-bottom: 1px solid rgba(0,0,0,0.1); }
    .admin-kpi-body { padding: 20px; }
    .admin-kpi-value { font-size: 2.8rem; font-weight: 900; color: #2C3E50; margin-bottom: 5px; }
    .admin-kpi-unit { font-size: 1rem; color: #7F8C8D; font-weight: normal; margin-left: 5px; }
    .admin-kpi-sub {
        font-size: 0.9rem; display: inline-block; padding: 2px 10px; border-radius: 15px;
        background-color: #F9E79F; color: #7D6608; margin-top: 5px; font-weight: bold;
    }

    /* 外部看板 KPI */
    .kpi-card {
        padding: 20px; border-radius: 15px; text-align: center; background-color: #FFFFFF;
        box-shadow: 0 4px 10px rgba(0,0,0,0.1); border: 1px solid var(--border-color); height: 100%; transition: transform 0.2s;
    }
    .kpi-card:hover { transform: translateY(-5px); }
    .kpi-gas { border-top: 8px solid var(--kpi-gas); } .kpi-diesel { border-top: 8px solid var(--kpi-diesel); }
    .kpi-total { border-top: 8px solid var(--kpi-total); } .kpi-co2 { border-top: 8px solid var(--kpi-co2); }
    .kpi-title { font-size: 1.2rem; font-weight: bold; opacity: 0.8; color: var(--text-sub) !important; margin-bottom: 5px; }
    .kpi-value { font-size: 2.8rem; font-weight: 800; color: var(--text-main) !important; margin: 0; }
    .kpi-unit { font-size: 1rem; font-weight: normal; color: var(--text-sub) !important; margin-left: 5px; }
    .kpi-sub { font-size: 0.9rem; color: #C0392B !important; font-weight: 700; background-color: rgba(192, 57, 43, 0.1); padding: 2px 10px; border-radius: 20px; display: inline-block; margin-top: 5px;}

    /* 未申報名單 */
    .unreported-block {
        padding: 15px 20px; border-radius: 12px; margin-bottom: 20px; color: #2C3E50;
        box-shadow: 0 2px 6px rgba(0,0,0,0.08); border: 1px solid rgba(0,0,0,0.05);
    }
    .unreported-title {
        font-size: 1.6rem; font-weight: 900; margin-bottom: 12px; border-bottom: 2px solid rgba(0,0,0,0.1); padding-bottom: 8px;
    }

    /* 其他 */
    .device-info-box { background-color: var(--card-bg); border: 2px solid #5DADE2; border-radius: 10px; padding: 20px; margin-bottom: 20px; }
    .alert-box { background-color: #FCF3CF; border: 2px solid #F1C40F; padding: 15px; border-radius: 10px; margin-bottom: 20px; color: #9A7D0A !important; font-weight: bold; text-align: center; }
    .login-header {
        font-size: 2.2rem; font-weight: 800; color: var(--text-main) !important; text-align: center;
        margin-bottom: 20px; padding: 25px; background-color: var(--card-bg); border: 2px solid var(--border-color); border-radius: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .privacy-box { background-color: #F8F9F9; border: 1px solid #BDC3C7; padding: 15px; border-radius: 10px; font-size: 0.9rem; color: #566573; margin-bottom: 10px; }
    .privacy-title { font-weight: bold; color: #2C3E50; margin-bottom: 5px; font-size: 1rem; }
    
    input[aria-label="搜尋框"] { height: 50px !important; font-size: 1.2rem !important; }
    .pie-chart-box { background-color: var(--card-bg); border: 2px solid var(--border-color); border-radius: 15px; padding: 10px; }
    .dashboard-main-title {
        font-size: 1.8rem; font-weight: 900; text-align: center; color: #2C3E50; margin-bottom: 20px;
        background-color: #F8F9F9; padding: 10px; border-radius: 10px; border: 1px solid #BDC3C7;
    }
    
    /* 一般設備卡片 */
    .equip-card { background-color: #FFFFFF; border: 1px solid #BDC3C7; border-radius: 10px; overflow: hidden; height: 100%; margin-bottom: 10px; box-shadow: 0 2px 5px rgba(0,0,0,0.05); }
    .equip-header { padding: 12px 18px; display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #BDC3C7; }
    .equip-code { font-size: 1.25rem; font-weight: 800; color: #2C3E50; }
    .equip-name { font-size: 1.1rem; color: #455A64; font-weight: 600; }
    .equip-vol { font-size: 1.6rem; font-weight: 900; color: var(--morandi-red); line-height: 1.2; }
    .equip-fuel-type { font-size: 0.85rem; color: #566573; font-weight: bold; background: rgba(255,255,255,0.6); padding: 2px 6px; border-radius: 4px; margin-left: 5px; }
    .equip-body { padding: 15px; display: flex; flex-direction: column; gap: 8px; }
    .equip-info { font-size: 0.95rem; line-height: 1.5; color: #34495E; }
    .equip-footer { padding: 10px 15px; border-top: 1px dashed #D7DBDD; display: flex; justify-content: space-between; align-items: center; }
    .status-badge { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 0.8rem; font-weight: bold; }
    .status-warn { background-color: #FADBD8; color: #943126; border: 1px solid #F1948A; }
    .last-date { font-size: 0.85rem; color: #7F8C8D; font-style: italic; }
    .count-text { font-size: 0.85rem; color: #2C3E50; font-weight: 800; margin-right: 8px; }
</style>
""", unsafe_allow_html=True)

# ☁️ 設定區
SHEET_ID = "1gqDU21YJeBoBOd8rMYzwwZ45offXWPGEODKTF6B8k-Y" 
DRIVE_FOLDER_ID = "1DCmR0dXOdFBdTrgnvCYFPtNq_bGzSJeB" 
VIP_UNITS = ["總務處事務組", "民雄總務", "新民聯辦", "產推處產學營運組"]
FLEET_CARDS = {"總務處事務組-柴油": "TZI510508", "總務處事務組-汽油": "TZI510509", "民雄總務": "TZI510594", "新民聯辦": "TZI510410", "產推處產學營運組": "TZI510244"}
DEVICE_ORDER = ["公務車輛(GV-1-)", "乘坐式割草機(GV-2-)", "乘坐式農用機具(GV-3-)", "鍋爐(GS-1-)", "發電機(GS-2-)", "肩背或手持式割草機、吹葉機(GS-3-)", "肩背或手持式農用機具(GS-4-)"]
DEVICE_CODE_MAP = {"GV-1": "公務車輛(GV-1-)", "GV-2": "乘坐式割草機(GV-2-)", "GV-3": "乘坐式農用機具(GV-3-)", "GS-1": "鍋爐(GS-1-)", "GS-2": "發電機(GS-2-)", "GS-3": "肩背或手持式割草機、吹葉機(GS-3-)", "GS-4": "肩背或手持式農用機具(GS-4-)"}

MORANDI_COLORS = {
    "公務車輛(GV-1-)": "#B0C4DE", "乘坐式割草機(GV-2-)": "#F5CBA7", "乘坐式農用機具(GV-3-)": "#D7BDE2",
    "鍋爐(GS-1-)": "#E6B0AA", "發電機(GS-2-)": "#A9CCE3",
    "肩背或手持式割草機、吹葉機(GS-3-)": "#A3E4D7", "肩背或手持式農用機具(GS-4-)": "#F9E79F"
}
UNREPORTED_COLORS = ["#D5DBDB", "#FAD7A0", "#D2B4DE", "#AED6F1", "#A3E4D7", "#F5B7B1"]
DASH_PALETTE = ['#B0C4DE', '#F5CBA7', '#A9CCE3', '#E6B0AA', '#D7BDE2', '#A3E4D7', '#F9E79F', '#95A5A6']

def clean_secrets(obj):
    if isinstance(obj, dict) or "AttrDict" in str(type(obj)): return {k: clean_secrets(v) for k, v in obj.items()}
    elif isinstance(obj, list): return [clean_secrets(i) for i in obj]
    return obj

if 'current_page' not in st.session_state: st.session_state['current_page'] = 'home'
if 'reset_counter' not in st.session_state: st.session_state['reset_counter'] = 0
if 'multi_row_count' not in st.session_state: st.session_state['multi_row_count'] = 1

try:
    _raw_creds = st.secrets["credentials"]
    credentials_login = clean_secrets(_raw_creds)
    cookie_cfg = st.secrets["cookie"]
    authenticator = stauth.Authenticate(credentials_login, cookie_cfg["name"], cookie_cfg["key"], cookie_cfg["expiry_days"])
    
    if st.session_state["authentication_status"] is not True:
        st.markdown('<div class="login-header">🏫 國立嘉義大學碳盤查<br>油料使用及冷媒填充回報平台</div>', unsafe_allow_html=True)
        st.markdown("<h3 style='text-align: center;'>登入系統 (Login)</h3>", unsafe_allow_html=True)
        st.markdown("---")
    
    authenticator.login('main')
    
    if st.session_state["authentication_status"] is False: st.error('❌ 帳號或密碼錯誤'); st.stop()
    elif st.session_state["authentication_status"] is None: st.info('🔒 請輸入帳號密碼登入'); st.stop()
        
    name, username = st.session_state["name"], st.session_state["username"]
    with st.sidebar:
        st.header(f"👤 {name}"); st.success("☁️ 雲端連線正常")
        if st.button("🏠 返回主選單"): st.session_state['current_page'] = 'home'; st.rerun()
        st.markdown("---"); authenticator.logout('登出系統 (Logout)', 'sidebar')
except Exception as e: st.error(f"登入錯誤: {e}"); st.stop()

@st.cache_resource
def init_google():
    oauth = st.secrets["gcp_oauth"]
    creds = Credentials(token=None, refresh_token=oauth["refresh_token"], token_uri="https://oauth2.googleapis.com/token", client_id=oauth["client_id"], client_secret=oauth["client_secret"], scopes=["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/spreadsheets"])
    gc = gspread.authorize(creds); drive = build('drive', 'v3', credentials=creds)
    return gc, drive

try:
    gc, drive_service = init_google(); sh = gc.open_by_key(SHEET_ID)
    try: ws_equip = sh.worksheet("設備清單") 
    except: ws_equip = sh.sheet1 
    try: ws_record = sh.worksheet("填報紀錄")
    except: ws_record = sh.add_worksheet(title="填報紀錄", rows="1000", cols="13")
    if len(ws_record.get_all_values()) == 0: ws_record.append_row(["填報時間", "填報單位", "填報人", "填報人分機", "設備名稱備註", "校內財產編號", "原燃物料名稱", "油卡編號", "加油日期", "加油量", "與其他設備共用加油單", "備註", "佐證資料"])
except Exception as e: st.error(f"連線失敗: {e}"); st.stop()

@st.cache_data(ttl=600)
def load_data():
    df_e = pd.DataFrame(ws_equip.get_all_records()).astype(str)
    if '設備編號' in df_e.columns:
        df_e['統計類別'] = df_e['設備編號'].apply(lambda c: next((v for k, v in DEVICE_CODE_MAP.items() if str(c).startswith(k)), "其他/未分類"))
    else: df_e['統計類別'] = "未設定I欄"
    # V119: 預處理數字
    df_e['設備數量_num'] = pd.to_numeric(df_e['設備數量'], errors='coerce').fillna(1)
    
    data = ws_record.get_all_values()
    df_r = pd.DataFrame(data[1:], columns=data[0]) if len(data) > 1 else pd.DataFrame(columns=data[0])
    return df_e, df_r

df_equip, df_records = load_data()

# ==========================================
# 3. 頁面邏輯
# ==========================================
if st.session_state['current_page'] == 'home':
    st.title("🏫 國立嘉義大學碳盤查回報平台")
    st.markdown("### 請選擇填報項目：")
    col1, col2 = st.columns(2)
    with col1:
        st.info("⛽ 車輛/機具用油")
        if st.button("前往「燃油設備填報區」", use_container_width=True, type="primary"): st.session_state['current_page'] = 'fuel'; st.rerun()
    with col2:
        st.info("❄️ 冷氣/冰水主機")
        st.button("前往「冷媒類設備填報區」 (建置中)", use_container_width=True, disabled=True)
    if username == 'admin':
        st.markdown("---"); st.markdown("### 👑 超級管理員專區")
        if st.button("進入「管理員後台」", use_container_width=True): st.session_state['current_page'] = 'admin_dashboard'; st.rerun()
    st.markdown('<div class="contact-footer">如有填報疑問，請電洽環安中心林小姐(分機 7137)，謝謝</div>', unsafe_allow_html=True)

# ------------------------------------------
# ⛽ 外部填報區 (V122.0: 一般申報UI/文字修正)
# ------------------------------------------
elif st.session_state['current_page'] == 'fuel':
    st.title("⛽ 燃油設備填報專區")
    tabs = st.tabs(["📝 新增填報", "📊 動態查詢看板"])

    # === Tab 1: 填報 ===
    with tabs[0]:
        st.markdown('<div class="alert-box">📢 請「誠實申報」，以保障單位及自身權益！</div>', unsafe_allow_html=True)
        if not df_equip.empty:
            st.markdown("#### 步驟 1：選擇設備或單位")
            c1, c2 = st.columns(2)
            units = sorted([x for x in df_equip['填報單位'].unique() if x != '-' and x != '填報單位'])
            selected_dept = c1.selectbox("填報單位", units, index=None, placeholder="請選擇單位...", key="dept_selector")
            
            # V119: 完整個資聲明文字
            privacy_html = """
            <div class="privacy-box">
                <div class="privacy-title">📜 個人資料蒐集、處理及利用告知聲明</div>
                1. <strong>蒐集機關</strong>：國立嘉義大學。<br>
                2. <strong>蒐集目的</strong>：進行本校公務車輛/機具之加油紀錄管理、校園溫室氣體（碳）盤查統計、稽核佐證資料蒐集及後續能源使用分析。<br>
                3. <strong>個資類別</strong>：填報人姓名。<br>
                4. <strong>利用期間</strong>：姓名保留至填報年度後第二年1月1日，期滿即進行「去識別化」刪除，其餘數據永久保存。<br>
                5. <strong>利用對象</strong>：本校教師、行政人員及碳盤查查驗人員。<br>
                6. <strong>您有權依個資法請求查詢、更正或刪除您的個資。如不提供，將無法完成填報。</strong><br>
            </div>
            """
            
            # V122: 深灰色備註文字 (通用)
            typo_note = '<div class="note-text-darkgray">如有資料誤繕情形，請重新登錄1次資訊，並於備註欄填寫：「前筆資料誤繕，請刪除。」，管理單位將協助刪除誤打資訊</div>'

            # --- 批次申報 ---
            if selected_dept in VIP_UNITS:
                st.info(f"💡 您選擇了 **{selected_dept}**，系統已自動切換為「油卡批次申報模式」。")
                sub_categories = []
                if selected_dept == "總務處事務組": sub_categories = ["具車牌的汽油公務車", "具車牌的柴油公務車", "無車牌的汽油機具", "無車牌的柴油機具"]
                elif selected_dept in ["民雄總務", "新民聯辦"]: sub_categories = ["無車牌的汽油機具", "無車牌的柴油機具"]
                elif selected_dept == "產推處產學營運組": sub_categories = ["無車牌的汽油機具"]
                
                target_sub_cat = c2.selectbox("請選擇細部類別", sub_categories, index=None, placeholder="請選擇...")
                
                if target_sub_cat:
                    def has_plate(name): return bool(re.search(r'\([A-Za-z0-9\-]+\)', name))
                    filtered_equip = df_equip[df_equip['填報單位'] == selected_dept].copy()
                    if "具車牌" in target_sub_cat: filtered_equip = filtered_equip[filtered_equip['設備名稱備註'].apply(has_plate)]
                    elif "無車牌" in target_sub_cat: filtered_equip = filtered_equip[~filtered_equip['設備名稱備註'].apply(has_plate)]
                    if "汽油" in target_sub_cat: filtered_equip = filtered_equip[filtered_equip['原燃物料名稱'].str.contains("汽油")]
                    elif "柴油" in target_sub_cat: filtered_equip = filtered_equip[filtered_equip['原燃物料名稱'].str.contains("柴油")]
                    
                    st.markdown("#### 步驟 2：批次填寫與上傳")
                    with st.form("batch_form", clear_on_submit=True):
                        col_p1, col_p2, col_p3 = st.columns(3)
                        p_name = col_p1.text_input("👤 填報人姓名 (必填)")
                        p_ext = col_p2.text_input("📞 聯絡分機 (必填)")
                        batch_date = col_p3.date_input("📅 加油月份 (日期統一選擇該月份最終日)", datetime.today())
                        
                        st.markdown("⛽ **請填入各設備該月份之加油總量(公升)，若該月份無使用請填0：**")
                        batch_inputs = {}
                        for idx, row in filtered_equip.iterrows():
                            # V120: 7:3 佈局, 間距加大
                            c_card, c_val = st.columns([7, 3]) 
                            with c_card:
                                header_color = MORANDI_COLORS.get(row.get('統計類別'), '#D5DBDB')
                                st.markdown(f"""
                                <div class="batch-card-final">
                                    <div class="batch-header-final" style="background-color: {header_color};">
                                        <span class="batch-title-text">⛽ {row['設備名稱備註']}</span>
                                        <span class="batch-qty-badge">數量: {row.get('設備數量','-')}</span>
                                    </div>
                                    <div class="batch-body-final">
                                        <div class="batch-row">
                                            <div class="batch-item">🏢 部門: {row.get('設備所屬單位/部門','-')}</div>
                                            <div class="batch-item">👤 保管人: {row.get('保管人','-')}</div>
                                        </div>
                                        <div class="batch-row">
                                            <div class="batch-item">⛽ 燃料: {row.get('原燃物料名稱')}</div>
                                            <div class="batch-item">🔢 財產編號: {row.get('校內財產編號','-')}</div>
                                        </div>
                                    </div>
                                </div>
                                """, unsafe_allow_html=True)
                            with c_val:
                                st.write("") 
                                st.write("") 
                                vol = st.number_input(f"加油量", min_value=0.0, step=0.1, key=f"b_v_{row['校內財產編號']}_{idx}", label_visibility="collapsed")
                                batch_inputs[idx] = vol
                                
                        st.markdown("---")
                        st.markdown("**📂 上傳中油加油明細 (只需一份)**")
                        f_file = st.file_uploader("支援 PDF/JPG/PNG", type=['pdf', 'jpg', 'png', 'jpeg'])
                        st.markdown("---")
                        
                        # V122: 備註與說明文字
                        st.text_input("備註 (選填)", key="batch_note")
                        st.markdown(typo_note, unsafe_allow_html=True)
                        st.write("")

                        st.markdown(privacy_html, unsafe_allow_html=True)
                        agree_privacy = st.checkbox("我已閱讀並同意個資聲明，且確認所填資料無誤。", value=False)
                        submitted = st.form_submit_button("🚀 批次確認送出", use_container_width=True)
                        
                        if submitted:
                            total_vol = sum(batch_inputs.values())
                            if not agree_privacy: st.error("❌ 請勾選同意聲明")
                            elif not p_name or not p_ext: st.warning("⚠️ 姓名與分機為必填")
                            elif not f_file: st.error("⚠️ 請上傳加油明細佐證")
                            else:
                                try:
                                    f_file.seek(0); file_ext = f_file.name.split('.')[-1]
                                    clean_name = f"BATCH_{selected_dept}_{batch_date}_{int(time.time())}.{file_ext}"
                                    file_meta = {'name': clean_name, 'parents': [DRIVE_FOLDER_ID]}
                                    media = MediaIoBaseUpload(f_file, mimetype=f_file.type, resumable=True)
                                    file = drive_service.files().create(body=file_meta, media_body=media, fields='webViewLink').execute()
                                    file_link = file.get('webViewLink')
                                    fleet_id = "-"
                                    if selected_dept == "總務處事務組": fleet_id = FLEET_CARDS.get(f"總務處事務組-{'汽油' if '汽油' in target_sub_cat else '柴油'}", "-")
                                    else: fleet_id = FLEET_CARDS.get(selected_dept, "-")
                                    
                                    rows_to_append = []
                                    current_time = get_taiwan_time().strftime("%Y-%m-%d %H:%M:%S")
                                    note_val = st.session_state.get("batch_note", "")
                                    
                                    for idx, vol in batch_inputs.items():
                                        row = filtered_equip.loc[idx]
                                        rows_to_append.append([current_time, selected_dept, p_name, p_ext, row['設備名稱備註'], str(row.get('校內財產編號','-')), row['原燃物料名稱'], fleet_id, str(batch_date), vol, "是", f"批次申報-{target_sub_cat} | {note_val}", file_link])
                                    if rows_to_append:
                                        ws_record.append_rows(rows_to_append); st.success(f"✅ 批次申報成功！已寫入 {len(rows_to_append)} 筆紀錄。")
                                        st.balloons(); st.session_state['reset_counter'] += 1; time.sleep(1.5); st.rerun()
                                    else: st.warning("系統錯誤：無法產生寫入資料。")
                                except Exception as e: st.error(f"失敗: {e}")

            # --- 一般申報模式 (V122: 標題、選單、欄位文字更新) ---
            else:
                filtered = df_equip[df_equip['填報單位'] == selected_dept]
                devices = sorted([x for x in filtered['設備名稱備註'].unique()])
                dynamic_key = f"vehicle_selector_{st.session_state['reset_counter']}"
                selected_device = c2.selectbox("車輛/機具名稱", devices, index=None, placeholder="請選擇車輛...", key=dynamic_key)
                if selected_device:
                    row = filtered[filtered['設備名稱備註'] == selected_device].iloc[0]
                    info_html = f"""<div class="device-info-box"><div style="border-bottom: 1px solid #BDC3C7; padding-bottom: 10px; margin-bottom: 10px; font-weight: bold; font-size: 1.2rem; color: #5DADE2;">📋 設備詳細資料</div><div><strong>🏢 部門：</strong>{row.get('設備所屬單位/部門', '-')}</div><div><strong>👤 保管人：</strong>{row.get('保管人', '-')}</div><div><strong>🔢 財產編號：</strong>{row.get('校內財產編號', '-')}</div><div><strong>📍 位置：</strong>{row.get('設備詳細位置/樓層', '-')}</div><div><strong>⛽ 燃料：</strong>{row.get('原燃物料名稱', '-')}</div><div><strong>📊 數量：</strong>{row.get('設備數量', '-')}</div></div>"""
                    st.markdown(info_html, unsafe_allow_html=True)
                    
                    st.markdown("#### 步驟2：填報設備加油資訊")
                    
                    # V122: 申報類型標題小字橘色
                    st.markdown('<p style="color:#E67E22; font-size:1rem; font-weight:bold; margin-bottom:-10px;">請選擇申報類型：</p>', unsafe_allow_html=True)
                    # V122: 選項樣式由 CSS .stRadio div[role="radiogroup"] label 控制
                    report_mode = st.radio("類型選擇", ["用油量申報 (含單筆/多筆/油卡)", "無使用"], horizontal=True, label_visibility="collapsed")
                    
                    if report_mode == "用油量申報 (含單筆/多筆/油卡)":
                        c_btn1, c_btn2, _ = st.columns([1, 1, 3])
                        with c_btn1: 
                            if st.button("➕ 增加一列"): st.session_state['multi_row_count'] += 1
                        with c_btn2: 
                            if st.button("➖ 減少一列") and st.session_state['multi_row_count'] > 1: st.session_state['multi_row_count'] -= 1
                        # V122: 增減列說明
                        st.caption("填報前請先設定申報筆數，至多10筆")

                    with st.form("entry_form", clear_on_submit=True):
                        col_p1, col_p2 = st.columns(2)
                        p_name = col_p1.text_input("👤 填報人姓名 (必填)")
                        p_ext = col_p2.text_input("📞 聯絡分機 (必填)")
                        fuel_card_id = ""; data_entries = []; f_files = None; note_input = ""
                        
                        if report_mode == "用油量申報 (含單筆/多筆/油卡)":
                            fuel_card_id = st.text_input("💳 油卡編號 (選填)")
                            for i in range(st.session_state['multi_row_count']):
                                c_d, c_v = st.columns(2)
                                # V122: 欄位文字更新
                                _date = c_d.date_input(f"📅 加油日期 填報序號 {i+1}", datetime.today(), key=f"d_{i}")
                                _vol = c_v.number_input(f"💧 加油量(公升) 填報序號 {i+1}", min_value=0.0, step=0.1, key=f"v_{i}")
                                data_entries.append({"date": _date, "vol": _vol})
                            
                            st.markdown("---")
                            is_shared = st.checkbox("與其他設備共用加油單") 
                            
                            note_input = st.text_input("備註", placeholder="")
                            st.markdown(typo_note, unsafe_allow_html=True)
                            
                            # V122: 檔案上傳文字更新
                            f_files = st.file_uploader("上傳佐證資料 (必填) - 支援 png, jpg, jpeg, pdf (單檔最多3MB，最多 可上傳10 個檔案)", type=['png', 'jpg', 'jpeg', 'pdf'], accept_multiple_files=True)
                        else:
                            st.info("ℹ️ 您選擇了「無使用」，請選擇無使用的期間。")
                            c_s, c_e = st.columns(2)
                            d_start = c_s.date_input("開始日期", datetime(datetime.now().year, 1, 1))
                            d_end = c_e.date_input("結束日期", datetime.now())
                            data_entries.append({"date": d_end, "vol": 0.0})
                            note_input = f"無使用 (期間: {d_start} ~ {d_end})"
                            st.markdown(typo_note, unsafe_allow_html=True)
                            is_shared = False

                        st.markdown("---"); st.markdown(privacy_html, unsafe_allow_html=True)
                        agree = st.checkbox("我已閱讀並同意個資聲明，且確認所填資料無誤。", value=False)
                        submitted = st.form_submit_button("🚀 確認送出", use_container_width=True)
                        
                        if submitted:
                            if not agree: st.error("❌ 請務必勾選同意聲明！")
                            elif not p_name or not p_ext: st.warning("⚠️ 姓名與分機為必填！")
                            elif report_mode == "用油量申報 (含單筆/多筆/油卡)":
                                if not f_files: st.error("⚠️ 請上傳佐證資料！")
                                elif len(f_files) > 10: st.error("⚠️ 最多只能上傳 10 個檔案！")
                                elif data_entries[0]['vol'] <= 0: st.warning("⚠️ 第一筆加油量不能為 0。")
                                else:
                                    # Process submission
                                    valid_logic = True; links=[]
                                    if f_files:
                                        for idx, f in enumerate(f_files):
                                            try:
                                                f.seek(0); clean_name = f"{selected_dept}_{selected_device}_{data_entries[0]['date']}_{idx+1}.{f.name.split('.')[-1]}"
                                                meta = {'name': clean_name, 'parents': [DRIVE_FOLDER_ID]}
                                                media = MediaIoBaseUpload(f, mimetype=f.type, resumable=True)
                                                file = drive_service.files().create(body=meta, media_body=media, fields='webViewLink').execute()
                                                links.append(file.get('webViewLink'))
                                            except: valid_logic=False; st.error("上傳失敗"); break
                                    
                                    if valid_logic:
                                        rows = []; now_str = get_taiwan_time().strftime("%Y-%m-%d %H:%M:%S")
                                        final_link = "\n".join(links) if links else "無"
                                        shared_str = "是" if is_shared else "-"; card_str = fuel_card_id if fuel_card_id else "-"
                                        for e in data_entries:
                                            rows.append([now_str, selected_dept, p_name, p_ext, selected_device, str(row.get('校內財產編號','-')), str(row.get('原燃物料名稱','-')), card_str, str(e['date']), e['vol'], shared_str, note_input, final_link])
                                        if rows: ws_record.append_rows(rows); st.success("✅ 申報成功！"); st.balloons(); st.session_state['reset_counter'] += 1; st.cache_data.clear(); st.rerun()
                            elif report_mode == "無使用":
                                # 無使用直接送出
                                rows = [[get_taiwan_time().strftime("%Y-%m-%d %H:%M:%S"), selected_dept, p_name, p_ext, selected_device, str(row.get('校內財產編號','-')), str(row.get('原燃物料名稱','-')), "-", str(data_entries[0]['date']), 0.0, "-", note_input, "無"]]
                                ws_record.append_rows(rows); st.success("✅ 申報成功！"); st.balloons(); st.session_state['reset_counter'] += 1; st.cache_data.clear(); st.rerun()

    # === Tab 2: 看板 (V122: 環形圖標籤 inside) ===
    with tabs[1]:
        st.markdown("### 📊 動態查詢看板 (年度檢視)")
        st.info("請選擇「單位」與「年份」，檢視該年度的用油統計與碳排放分析。")
        
        col_r1, col_r2 = st.columns([4, 1])
        with col_r2:
            if st.button("🔄 刷新數據", use_container_width=True, key="refresh_all"): 
                st.cache_data.clear()
                st.rerun()
        
        available_years = []
        if not df_records.empty:
            df_records['加油量'] = pd.to_numeric(df_records['加油量'], errors='coerce').fillna(0)
            df_records['日期格式'] = pd.to_datetime(df_records['加油日期'], errors='coerce')
            available_years = sorted(df_records['日期格式'].dt.year.dropna().astype(int).unique(), reverse=True)
            if not available_years: available_years = [datetime.now().year]
            record_units = sorted([str(x) for x in df_records['填報單位'].unique() if str(x) != 'nan'])
            
            c_dept, c_year = st.columns([2, 1])
            query_dept = c_dept.selectbox("🏢 選擇查詢單位", record_units, index=None, placeholder="請選擇...")
            query_year = c_year.selectbox("📅 選擇統計年度", available_years, index=0) 
            
            if query_dept and query_year:
                df_dept = df_records[df_records['填報單位'] == query_dept].copy()
                df_final = df_dept[df_dept['日期格式'].dt.year == query_year]
                
                if not df_final.empty:
                    # HTML KPI
                    if '原燃物料名稱' in df_final.columns:
                        gas_sum = df_final[df_final['原燃物料名稱'].str.contains('汽油', na=False)]['加油量'].sum()
                        diesel_sum = df_final[df_final['原燃物料名稱'].str.contains('柴油', na=False)]['加油量'].sum()
                        total_co2 = (gas_sum * 0.0022) + (diesel_sum * 0.0027)
                    else: gas_sum = 0; diesel_sum = 0; total_co2 = 0
                    total_sum = df_final['加油量'].sum()
                    gas_pct = (gas_sum / total_sum * 100) if total_sum > 0 else 0
                    diesel_pct = (diesel_sum / total_sum * 100) if total_sum > 0 else 0
                    
                    st.markdown(f"<div class='dashboard-main-title'>{query_dept} - {query_year}年度 能源使用與碳排統計</div>", unsafe_allow_html=True)
                    r1c1, r1c2 = st.columns(2)
                    with r1c1: st.markdown(f"""<div class="kpi-card kpi-gas"><div class="kpi-title">⛽ 汽油使用量</div><div class="kpi-value">{gas_sum:,.2f}<span class="kpi-unit"> 公升</span></div><div class="kpi-sub">佔比 {gas_pct:.2f}%</div></div>""", unsafe_allow_html=True)
                    with r1c2: st.markdown(f"""<div class="kpi-card kpi-diesel"><div class="kpi-title">🚛 柴油使用量</div><div class="kpi-value">{diesel_sum:,.2f}<span class="kpi-unit"> 公升</span></div><div class="kpi-sub">佔比 {diesel_pct:.2f}%</div></div>""", unsafe_allow_html=True)
                    st.markdown("<div style='margin-bottom: 20px;'></div>", unsafe_allow_html=True)
                    r2c1, r2c2 = st.columns(2)
                    with r2c1: st.markdown(f"""<div class="kpi-card kpi-total"><div class="kpi-title">💧 總用油量</div><div class="kpi-value">{total_sum:,.2f}<span class="kpi-unit"> 公升</span></div><div class="kpi-sub">100%</div></div>""", unsafe_allow_html=True)
                    with r2c2: st.markdown(f"""<div class="kpi-card kpi-co2"><div class="kpi-title">☁️ 碳排放量</div><div class="kpi-value">{total_co2:,.4f}<span class="kpi-unit"> 公噸CO<sub>2</sub>e</span></div><div class="kpi-sub" style="background-color: #F4ECF7; color: #AF7AC5 !important;">ESG 指標</div></div>""", unsafe_allow_html=True)
                    st.markdown("---")
                    
                    st.subheader(f"📊 {query_year}年度 逐月油料統計", anchor=False)
                    filter_mode = st.radio("顯示類別", ["全部顯示", "只看汽油", "只看柴油"], horizontal=True)
                    df_final['月份'] = df_final['日期格式'].dt.month
                    df_final['油品類別'] = df_final['原燃物料名稱'].apply(lambda x: '汽油' if '汽油' in x else ('柴油' if '柴油' in x else '其他'))
                    months = list(range(1, 13))
                    if filter_mode == "全部顯示": target_fuels = ['汽油', '柴油']
                    elif filter_mode == "只看汽油": target_fuels = ['汽油']
                    else: target_fuels = ['柴油']
                    base_x = pd.MultiIndex.from_product([months, target_fuels], names=['月份', '油品類別']).to_frame(index=False)
                    unique_devices = df_final['設備名稱備註'].unique()
                    
                    fig = go.Figure()
                    morandi_colors = ['#88B04B', '#92A8D1', '#F7CAC9', '#B565A7', '#009B77', '#DD4124', '#D65076', '#45B8AC', '#EFC050', '#5B5EA6']
                    device_color_map = {dev: morandi_colors[i % len(morandi_colors)] for i, dev in enumerate(unique_devices)}
                    for dev in unique_devices:
                        dev_data = df_final[df_final['設備名稱備註'] == dev]
                        dev_grouped = dev_data.groupby(['月份', '油品類別'])['加油量'].sum().reset_index()
                        merged_dev = pd.merge(base_x, dev_grouped, on=['月份', '油品類別'], how='left').fillna(0)
                        fig.add_trace(go.Bar(x=[merged_dev['月份'], merged_dev['油品類別']], y=merged_dev['加油量'], name=dev, marker_color=device_color_map[dev], text=merged_dev['加油量'].apply(lambda x: f"{x:.1f}" if x > 0 else ""), texttemplate='%{text}', textposition='inside'))
                    
                    total_grouped = df_final.groupby(['月份', '油品類別'])['加油量'].sum().reset_index()
                    merged_total = pd.merge(base_x, total_grouped, on=['月份', '油品類別'], how='left').fillna(0)
                    label_data = merged_total[merged_total['加油量'] > 0]
                    fig.add_trace(go.Scatter(x=[label_data['月份'], label_data['油品類別']], y=label_data['加油量'], text=label_data['加油量'].apply(lambda x: f"{x:.1f}"), mode='text', textposition='top center', textfont=dict(size=14, color='black'), showlegend=False))
                    fig.update_layout(barmode='stack', font=dict(size=14), xaxis=dict(title="月份 / 油品"), yaxis=dict(title="加油量 (公升)"), height=550, margin=dict(t=50, b=120))
                    st.plotly_chart(fig, use_container_width=True)
                    
                    st.markdown("---")
                    st.subheader(f"🌞 單位油料使用碳排放量(公噸二氧化碳當量)結構", anchor=False)
                    df_final['CO2e'] = df_final.apply(lambda r: r['加油量']*0.0022 if '汽油' in r['原燃物料名稱'] else r['加油量']*0.0027, axis=1)
                    treemap_data = df_final.groupby(['設備名稱備註'])['CO2e'].sum().reset_index()
                    fig_tree = px.treemap(treemap_data, path=['設備名稱備註'], values='CO2e', title=f"{query_dept} - 設備碳排放量權重分析", color='CO2e', color_continuous_scale='Teal')
                    fig_tree.update_traces(textinfo="label+value+percent entry", textfont=dict(size=24))
                    fig_tree.update_coloraxes(showscale=False)
                    st.plotly_chart(fig_tree, use_container_width=True)

                    st.subheader("🍩 油品設備用油量佔比分析", anchor=False)
                    c_pie1, c_pie2 = st.columns(2)
                    with c_pie1:
                        st.markdown('<div class="pie-chart-box">', unsafe_allow_html=True) 
                        gas_df = df_final[df_final['原燃物料名稱'].str.contains('汽油', na=False)]
                        if not gas_df.empty:
                            fig_gas = px.pie(gas_df, values='加油量', names='設備名稱備註', title='⛽ 汽油設備用油量分析', color_discrete_sequence=px.colors.sequential.Teal, hole=0.5)
                            # V122: 標籤 inside
                            fig_gas.update_traces(textinfo='percent+label', textfont_size=16, textposition='inside', insidetextorientation='horizontal')
                            fig_gas.update_layout(legend=dict(orientation="h", yanchor="top", y=-0.1, xanchor="center", x=0.5), margin=dict(l=40, r=40, t=40, b=40))
                            st.plotly_chart(fig_gas, use_container_width=True)
                        else: st.info("無汽油使用紀錄")
                        st.markdown('</div>', unsafe_allow_html=True) 
                    with c_pie2:
                        st.markdown('<div class="pie-chart-box">', unsafe_allow_html=True) 
                        diesel_df = df_final[df_final['原燃物料名稱'].str.contains('柴油', na=False)]
                        if not diesel_df.empty:
                            fig_diesel = px.pie(diesel_df, values='加油量', names='設備名稱備註', title='🚛 柴油設備用油量分析', color_discrete_sequence=px.colors.sequential.Oranges, hole=0.5)
                            # V122: 標籤 inside
                            fig_diesel.update_traces(textinfo='percent+label', textfont_size=16, textposition='inside', insidetextorientation='horizontal')
                            fig_diesel.update_layout(legend=dict(orientation="h", yanchor="top", y=-0.1, xanchor="center", x=0.5), margin=dict(l=40, r=40, t=40, b=40))
                            st.plotly_chart(fig_diesel, use_container_width=True)
                        else: st.info("無柴油使用紀錄")
                        st.markdown('</div>', unsafe_allow_html=True) 
                    
                    st.markdown("---")
                    st.subheader(f"📋 {query_year}年度 填報明細")
                    df_display = df_final[["加油日期", "設備名稱備註", "原燃物料名稱", "油卡編號", "加油量", "填報人", "備註"]].sort_values(by='加油日期', ascending=False).rename(columns={'加油量': '加油量(公升)'})
                    st.dataframe(df_display.style.format({"加油量(公升)": "{:.2f}"}), use_container_width=True)
                else: st.warning(f"⚠️ {query_dept} 在 {query_year} 年度尚無填報紀錄。")
        else: st.warning("📭 目前資料庫尚無有效資料，請先至「新增填報」分頁填寫。")
        st.markdown('<div class="contact-footer">如有填報疑問，請電洽環安中心林小姐(分機 7137)，謝謝</div>', unsafe_allow_html=True)

# ------------------------------------------
# 👑 超級管理員專區 (V122.0: 完工版)
# ------------------------------------------
elif st.session_state['current_page'] == 'admin_dashboard' and username == 'admin':
    st.title("👑 超級管理員後台")
    
    # 1. 核心資料預處理
    df_clean = df_records.copy()
    if not df_clean.empty:
        df_clean['加油量'] = pd.to_numeric(df_clean['加油量'], errors='coerce').fillna(0)
        df_clean['日期格式'] = pd.to_datetime(df_clean['加油日期'], errors='coerce')
        df_clean['年份'] = df_clean['日期格式'].dt.year.fillna(0).astype(int)
        df_clean['月份'] = df_clean['日期格式'].dt.month.fillna(0).astype(int)
        df_clean['油品大類'] = df_clean['原燃物料名稱'].apply(lambda x: '汽油' if '汽油' in str(x) else ('柴油' if '柴油' in str(x) else '其他'))
        
        if not df_equip.empty:
            device_map = pd.Series(df_equip['統計類別'].values, index=df_equip['設備名稱備註']).to_dict()
            df_clean['統計類別'] = df_clean['設備名稱備註'].map(device_map).fillna("其他/未分類")

    all_years = sorted(df_clean['年份'][df_clean['年份']>0].unique(), reverse=True) if not df_clean.empty else [datetime.now().year]
    c_year, _ = st.columns([1, 3])
    selected_admin_year = c_year.selectbox("📅 請選擇檢視年度", all_years, index=0)
    
    df_year = df_clean[df_clean['年份'] == selected_admin_year] if not df_clean.empty else pd.DataFrame()

    # V119/V122: 架構重組 - 4大分頁 (異動 -> 未申報 -> 總覽 -> 儀表板)
    admin_tabs = st.tabs(["🔍 申報資料異動", "⚠️ 篩選未申報名單", "📝 全校燃油設備總覽", "📊 全校油料使用儀表板"])

    # === Tab 1: 申報資料異動 ===
    with admin_tabs[0]:
        st.subheader("🔍 申報資料異動")
        if not df_year.empty:
            df_year['加油日期'] = pd.to_datetime(df_year['加油日期']).dt.date
            edited = st.data_editor(df_year, column_config={"佐證資料": st.column_config.LinkColumn("佐證", display_text="🔗"), "加油日期": st.column_config.DateColumn("日期", format="YYYY-MM-DD"), "加油量": st.column_config.NumberColumn("油量", format="%.2f"), "填報時間": st.column_config.TextColumn("填報時間", disabled=True)}, num_rows="dynamic", use_container_width=True, key="editor_v122")
            if st.button("💾 儲存變更", type="primary"):
                try:
                    ws_record.clear()
                    exp = edited.copy(); exp['加油日期'] = exp['加油日期'].astype(str)
                    ws_record.update([exp.columns.tolist()] + exp.astype(str).values.tolist())
                    st.success("✅ 更新成功！"); st.cache_data.clear(); time.sleep(1); st.rerun()
                except Exception as e: st.error(f"更新失敗: {e}")
        else: st.info(f"{selected_admin_year} 年度尚無資料。")

    # === Tab 2: 篩選未申報 ===
    with admin_tabs[1]:
        st.subheader("⚠️ 篩選未申報名單")
        c_f1, c_f2 = st.columns(2)
        d_start = c_f1.date_input("查詢起始日", date(selected_admin_year, 1, 1))
        d_end = c_f2.date_input("查詢結束日", date.today())
        
        if st.button("開始篩選"):
            if not df_clean.empty:
                mask = (df_clean['日期格式'].dt.date >= d_start) & (df_clean['日期格式'].dt.date <= d_end)
                reported = set(df_clean[mask]['設備名稱備註'].unique())
                df_eq_copy = df_equip.copy()
                df_eq_copy['已申報'] = df_eq_copy['設備名稱備註'].apply(lambda x: x in reported)
                unreported = df_eq_copy[~df_eq_copy['已申報']]
                
                if not unreported.empty:
                    st.error(f"🚩 期間 [{d_start} ~ {d_end}] 共有 {len(unreported)} 台設備未申報！")
                    for idx, (unit, group) in enumerate(unreported.groupby('填報單位')):
                        bg_color = UNREPORTED_COLORS[idx % len(UNREPORTED_COLORS)]
                        st.markdown(f"""<div class="unreported-block" style="background-color: {bg_color};"><div class="unreported-title">🏢 {unit} (未申報數: {len(group)})</div></div>""", unsafe_allow_html=True)
                        st.dataframe(group[['設備名稱備註', '保管人', '校內財產編號']], use_container_width=True)
                else: st.success("🎉 太棒了！全數已申報。")
            else: st.warning("無資料可供篩選。")

    # === Tab 3: 全校總覽 (各類設備數量及用油統計) ===
    with admin_tabs[2]:
        if not df_year.empty and not df_equip.empty:
            # 1. 關鍵數字 KPI
            total_eq = int(df_equip['設備數量_num'].sum())
            gas_eq = int(df_equip[df_equip['原燃物料名稱'].str.contains('汽油', na=False)]['設備數量_num'].sum())
            diesel_eq = int(df_equip[df_equip['原燃物料名稱'].str.contains('柴油', na=False)]['設備數量_num'].sum())
            
            k1, k2, k3 = st.columns(3)
            k1.markdown(f"""<div class="top-kpi-card"><div class="top-kpi-title">🚜 全校燃油設備總數</div><div class="top-kpi-value">{total_eq}</div></div>""", unsafe_allow_html=True)
            k2.markdown(f"""<div class="top-kpi-card"><div class="top-kpi-title">⛽ 全校汽油設備數</div><div class="top-kpi-value">{gas_eq}</div></div>""", unsafe_allow_html=True)
            k3.markdown(f"""<div class="top-kpi-card"><div class="top-kpi-title">🚛 全校柴油設備數</div><div class="top-kpi-value">{diesel_eq}</div></div>""", unsafe_allow_html=True)
            st.markdown("---")

            # 2. 各類設備用油統計 (V122: 標題更新)
            st.subheader("📂 各類設備數量及用油統計")
            eq_sums = df_equip.groupby('統計類別')['設備數量_num'].sum()
            eq_gas_sums = df_equip[df_equip['原燃物料名稱'].str.contains('汽油', na=False)].groupby('統計類別')['設備數量_num'].sum()
            eq_dsl_sums = df_equip[df_equip['原燃物料名稱'].str.contains('柴油', na=False)].groupby('統計類別')['設備數量_num'].sum()
            
            fuel_sums = df_year.groupby(['統計類別', '油品大類'])['加油量'].sum().unstack(fill_value=0)
            
            for i in range(0, len(DEVICE_ORDER), 2):
                cols = st.columns(2)
                for j in range(2):
                    if i + j < len(DEVICE_ORDER):
                        category = DEVICE_ORDER[i + j]
                        with cols[j]:
                            # 計算設備數
                            count_tot = int(eq_sums.get(category, 0))
                            count_gas = int(eq_gas_sums.get(category, 0))
                            count_dsl = int(eq_dsl_sums.get(category, 0))
                            
                            # 計算油量
                            gas_vol = fuel_sums.loc[category, '汽油'] if category in fuel_sums.index and '汽油' in fuel_sums.columns else 0
                            diesel_vol = fuel_sums.loc[category, '柴油'] if category in fuel_sums.index and '柴油' in fuel_sums.columns else 0
                            total_vol = gas_vol + diesel_vol
                            header_color = MORANDI_COLORS.get(category, "#CFD8DC")
                            
                            # V122: 卡片欄位名稱更新
                            st.markdown(f"""
                            <div class="stat-card-v119">
                                <div class="stat-header" style="background-color: {header_color};">
                                    <span class="stat-title">{category}</span>
                                    <span class="stat-count">{count_tot}</span>
                                </div>
                                <div class="stat-body-split">
                                    <div class="stat-col-left">
                                        <div class="stat-item"><span class="stat-item-label">⛽ 汽油設備數</span><span class="stat-item-val">{count_gas}</span></div>
                                        <div class="stat-item"><span class="stat-item-label">🚛 柴油設備數</span><span class="stat-item-val">{count_dsl}</span></div>
                                        <div class="stat-item"><span class="stat-item-label">🔥 燃油設備數</span><span class="stat-item-val">{count_tot}</span></div>
                                    </div>
                                    <div class="stat-col-right">
                                        <div class="stat-item"><span class="stat-item-label">汽油加油量(公升)</span><span class="stat-item-val">{gas_vol:,.1f}</span></div>
                                        <div class="stat-item"><span class="stat-item-label">柴油加油量(公升)</span><span class="stat-item-val">{diesel_vol:,.1f}</span></div>
                                        <div class="stat-item"><span class="stat-item-label">總計加油量(公升)</span><span class="stat-item-val">{total_vol:,.1f}</span></div>
                                    </div>
                                </div>
                            </div>
                            """, unsafe_allow_html=True)
            
            st.markdown("---")
            # 3. 環形圖
            st.subheader("🍩 油品設備用油量佔比分析")
            c_pie1, c_pie2 = st.columns(2)
            color_map = {
                "公務車輛(GV-1-)": "#B0C4DE", "乘坐式割草機(GV-2-)": "#F5CBA7", "乘坐式農用機具(GV-3-)": "#D7BDE2",
                "鍋爐(GS-1-)": "#E6B0AA", "發電機(GS-2-)": "#A9CCE3",
                "肩背或手持式割草機、吹葉機(GS-3-)": "#A3E4D7", "肩背或手持式農用機具(GS-4-)": "#F9E79F"
            }
            gas_data = df_year[(df_year['油品大類'] == '汽油') & (df_year['統計類別'].isin(DEVICE_ORDER))].groupby('統計類別')['加油量'].sum().reset_index()
            if not gas_data.empty:
                fig_g = px.pie(gas_data, values='加油量', names='統計類別', title='⛽ 汽油用量佔比', hole=0.4, color='統計類別', color_discrete_map=color_map)
                fig_g.update_traces(textinfo='percent+label', textfont_size=18, insidetextorientation='horizontal')
                c_pie1.plotly_chart(fig_g, use_container_width=True)
            else: c_pie1.info("無汽油數據")
            
            dsl_data = df_year[(df_year['油品大類'] == '柴油') & (df_year['統計類別'].isin(DEVICE_ORDER))].groupby('統計類別')['加油量'].sum().reset_index()
            if not dsl_data.empty:
                fig_d = px.pie(dsl_data, values='加油量', names='統計類別', title='🚛 柴油用量佔比', hole=0.4, color='統計類別', color_discrete_map=color_map)
                fig_d.update_traces(textinfo='percent+label', textfont_size=18, insidetextorientation='horizontal')
                c_pie2.plotly_chart(fig_d, use_container_width=True)
            else: c_pie2.info("無柴油數據")
        else: st.warning("尚無資料可供統計。")

    # === Tab 4: 儀表板 ===
    with admin_tabs[3]:
        if not df_year.empty:
            st.markdown(f"<div class='dashboard-main-title'>{selected_admin_year}年度 能源使用與碳排統計</div>", unsafe_allow_html=True)
            
            gas_sum = df_year[df_year['油品大類'] == '汽油']['加油量'].sum()
            diesel_sum = df_year[df_year['油品大類'] == '柴油']['加油量'].sum()
            total_sum = df_year['加油量'].sum()
            total_co2 = (gas_sum * 0.0022) + (diesel_sum * 0.0027)
            gas_pct = (gas_sum / total_sum * 100) if total_sum > 0 else 0
            diesel_pct = (diesel_sum / total_sum * 100) if total_sum > 0 else 0
            
            c1, c2 = st.columns(2)
            with c1: st.markdown(f"""<div class="admin-kpi-card"><div class="admin-kpi-header" style="background-color: #B0C4DE;">⛽ 汽油使用量</div><div class="admin-kpi-body"><div class="admin-kpi-value">{gas_sum:,.2f}<span class="admin-kpi-unit">公升</span></div><div class="admin-kpi-sub">佔比 {gas_pct:.1f}%</div></div></div>""", unsafe_allow_html=True)
            with c2: st.markdown(f"""<div class="admin-kpi-card"><div class="admin-kpi-header" style="background-color: #F5CBA7;">🚛 柴油使用量</div><div class="admin-kpi-body"><div class="admin-kpi-value">{diesel_sum:,.2f}<span class="admin-kpi-unit">公升</span></div><div class="admin-kpi-sub">佔比 {diesel_pct:.1f}%</div></div></div>""", unsafe_allow_html=True)
            st.write("") 
            c3, c4 = st.columns(2)
            with c3: st.markdown(f"""<div class="admin-kpi-card"><div class="admin-kpi-header" style="background-color: #A9CCE3;">💧 總用油量</div><div class="admin-kpi-body"><div class="admin-kpi-value">{total_sum:,.2f}<span class="admin-kpi-unit">公升</span></div><div class="admin-kpi-sub">100%</div></div></div>""", unsafe_allow_html=True)
            with c4: st.markdown(f"""<div class="admin-kpi-card"><div class="admin-kpi-header" style="background-color: #E6B0AA;">☁️ 碳排放量</div><div class="admin-kpi-body"><div class="admin-kpi-value">{total_co2:,.4f}<span class="admin-kpi-unit">公噸CO<sub>2</sub>e</span></div><div class="admin-kpi-sub">ESG 指標</div></div></div>""", unsafe_allow_html=True)
            st.markdown("---")

            st.subheader("📈 全校逐月加油量統計")
            monthly = df_year.groupby(['月份', '油品大類'])['加油量'].sum().reset_index()
            full_months = pd.DataFrame({'月份': range(1, 13)})
            monthly = full_months.merge(monthly, on='月份', how='left').fillna({'加油量':0, '油品大類':'汽油'})
            fig_month = px.bar(monthly, x='月份', y='加油量', color='油品大類', barmode='group', text_auto='.1f', color_discrete_sequence=DASH_PALETTE)
            fig_month.update_layout(xaxis=dict(tickmode='linear', tick0=1, dtick=1), yaxis_title="加油量(公升)", font=dict(size=18), showlegend=True)
            st.plotly_chart(fig_month, use_container_width=True)

            st.subheader("🏆 全校前十大加油量單位")
            top_fuel = st.radio("選擇油品類型", ["汽油", "柴油"], horizontal=True)
            df_top = df_year[df_year['油品大類'] == top_fuel]
            if not df_top.empty:
                top10_data = df_top.groupby('填報單位')['加油量'].sum().nlargest(10).reset_index()
                fig_top = px.bar(top10_data, x='填報單位', y='加油量', text_auto='.1f', title=f"{top_fuel}用量前十大單位", color_discrete_sequence=DASH_PALETTE)
                fig_top.update_layout(xaxis={'categoryorder':'total descending'}, yaxis_title="加油量(公升)", font=dict(size=16))
                st.plotly_chart(fig_top, use_container_width=True)
            else: st.info("無此油品數據。")

            st.markdown("---")
            st.subheader("🍩 全校加油量單位佔比")
            c_d1, c_d2 = st.columns(2)
            with c_d1:
                df_gas = df_year[df_year['油品大類'] == '汽油']
                if not df_gas.empty:
                    fig_dg = px.pie(df_gas, values='加油量', names='填報單位', title='⛽ 汽油用量分佈', hole=0.4, color_discrete_sequence=DASH_PALETTE)
                    fig_dg.update_traces(textposition='inside', textinfo='label+percent', hovertemplate='%{label}<br>加油量: %{value:.2f} L<br>佔比: %{percent}', textfont_size=18, insidetextorientation='horizontal')
                    st.plotly_chart(fig_dg, use_container_width=True)
                else: st.info("無汽油數據")
            with c_d2:
                df_dsl = df_year[df_year['油品大類'] == '柴油']
                if not df_dsl.empty:
                    fig_dd = px.pie(df_dsl, values='加油量', names='填報單位', title='🚛 柴油用量分佈', hole=0.4, color_discrete_sequence=DASH_PALETTE)
                    fig_dd.update_traces(textposition='inside', textinfo='label+percent', hovertemplate='%{label}<br>加油量: %{value:.2f} L<br>佔比: %{percent}', textfont_size=18, insidetextorientation='horizontal')
                    st.plotly_chart(fig_dd, use_container_width=True)
                else: st.info("無柴油數據")

            st.markdown("---")
            st.subheader("🌍 全校油料使用碳排放量(公噸二氧化碳當量)結構")
            df_year['CO2e'] = df_year.apply(lambda r: r['加油量']*0.0022 if '汽油' in str(r['原燃物料名稱']) else r['加油量']*0.0027, axis=1)
            if not df_year.empty:
                fig_tree = px.treemap(df_year, path=['填報單位', '設備名稱備註'], values='CO2e', color='填報單位', color_discrete_sequence=DASH_PALETTE)
                fig_tree.update_traces(textinfo="label+value+percent root", textfont=dict(size=18))
                st.plotly_chart(fig_tree, use_container_width=True)
            else: st.info("無數據")
        else: st.info("尚無該年度資料，無法顯示儀表板。")

    st.markdown('<div class="contact-footer">管理員系統版本 V122.0 (Verified Correction)</div>', unsafe_allow_html=True)