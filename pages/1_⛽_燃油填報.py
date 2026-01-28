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
st.set_page_config(page_title="燃油設備填報", page_icon="⛽", layout="wide")

def get_taiwan_time():
    return datetime.utcnow() + timedelta(hours=8)

# ==========================================
# 1. CSS 樣式表 (V163.0 定案修正)
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

    /* 輸入元件優化 */
    div[data-baseweb="input"] > div, div[data-baseweb="base-input"] > input, textarea, input {
        background-color: #FFFFFF !important; border-color: #BDC3C7 !important; color: #000000 !important; font-size: 1.15rem !important;
    }
    div[data-baseweb="select"] > div { border-color: #BDC3C7 !important; background-color: #FFFFFF !important; }
    ul[data-baseweb="menu"] { background-color: #FFFFFF !important; }

    /* 按鈕樣式 */
    div.stButton > button, button[kind="primary"], [data-testid="stFormSubmitButton"] > button {
        background-color: var(--orange-bg) !important; 
        color: #FFFFFF !important; border: 2px solid var(--orange-dark) !important; border-radius: 12px !important;
        font-size: 1.3rem !important; font-weight: 800 !important; padding: 0.7rem 1.5rem !important;
        box-shadow: 0 4px 6px rgba(230, 126, 34, 0.3) !important; width: 100%; 
    }
    div.stButton > button p { color: #FFFFFF !important; } 
    div.stButton > button:hover, [data-testid="stFormSubmitButton"] > button:hover { 
        background-color: var(--orange-dark) !important; transform: translateY(-2px) !important; color: #FFFFFF !important;
    }

    /* Tab 分頁字體 */
    button[data-baseweb="tab"] div p { font-size: 1.3rem !important; font-weight: 900 !important; color: var(--text-sub); }
    button[data-baseweb="tab"][aria-selected="true"] div p { color: #E67E22 !important; border-bottom: 3px solid #E67E22; }

    /* Checkbox & Upload */
    div[data-testid="stCheckbox"] label p { font-size: 1.2rem !important; color: #1F618D !important; font-weight: 900 !important; }
    [data-testid="stFileUploaderDropzone"] { background-color: #D6EAF8 !important; border: 2px dashed #2E86C1 !important; padding: 20px; border-radius: 12px; }
    [data-testid="stFileUploaderDropzone"] div, span, small { color: #154360 !important; font-weight: bold !important; }

    /* 選項標籤設計 (淺藍底、黑字、字體放大) */
    .stRadio div[role="radiogroup"] label {
        background-color: #D6EAF8 !important; 
        border: 1px solid #AED6F1 !important;
        border-radius: 8px !important; 
        padding: 8px 15px !important; 
        margin-right: 10px !important;
    }
    .stRadio div[role="radiogroup"] label p { 
        font-size: 1.25rem !important; 
        font-weight: 800 !important; 
        color: #000000 !important; 
    }

    /* 表格字體放大 */
    [data-testid="stDataFrame"] { font-size: 1.25rem !important; }
    [data-testid="stDataFrame"] div { font-size: 1.25rem !important; }

    /* --- 設備詳細卡片樣式 (V163.0 修正回 Grid) --- */
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

    /* V163 修正: 回復 Grid 2欄, 確保 User/Admin 顯示一致且同一列 */
    .dev-body {
        padding: 15px; font-size: 0.95rem; color: var(--deep-gray);
        display: grid; grid-template-columns: 1fr 1fr; gap: 8px; 
    }
    /* V163 修正: 標籤與數值同一列 (inline) */
    .dev-item { margin-bottom: 2px; display: flex; align-items: baseline; }
    .dev-label { font-weight: 700; color: var(--deep-gray) !important; font-size: 0.95rem; margin-right: 5px; min-width: 80px; }
    .dev-val { color: var(--deep-gray) !important; font-weight: 600; font-size: 1rem; }
    
    .dev-footer {
        padding: 10px 15px; background-color: #F8F9F9; border-top: 1px solid #E5E7E9;
        display: flex; justify-content: space-between; align-items: center;
    }
    .dev-count { font-weight: 700; color: #34495E; font-size: 0.95rem; }
    .alert-status { color: #C0392B; font-weight: 900; display: flex; align-items: center; gap: 5px; background-color: #FADBD8; padding: 4px 12px; border-radius: 12px; font-size: 0.9rem; }

    /* 批次申報卡片 */
    .batch-card-final {
        background-color: #FFFFFF; border: 1px solid #BDC3C7; border-radius: 10px; overflow: hidden;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1); height: 100%; display: flex; flex-direction: column;
        border-left: 5px solid #E67E22; margin-bottom: 15px; 
    }
    .batch-header-final { padding: 10px 15px; font-weight: 800; color: #2C3E50; display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid rgba(0,0,0,0.1); font-size: 1.1rem; background-color: #F4F6F6; }
    .batch-qty-badge { font-size: 0.9rem; background-color: rgba(255,255,255,0.7); padding: 2px 8px; border-radius: 12px; border: 1px solid rgba(0,0,0,0.1); color: #2C3E50; font-weight: bold; }
    .batch-body-final { background-color: #FFFFFF; padding: 12px; font-size: 0.95rem; color: #566573; line-height: 1.6; flex-grow: 1; }
    .batch-row { display: flex; justify-content: space-between; margin-bottom: 5px; }
    .batch-item { flex: 1; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; margin-right: 5px; }

    /* 後台 - 統計卡片 (V162 修正: 標題與數值放大) */
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

    /* 看板 KPI */
    .kpi-card { padding: 20px; border-radius: 15px; text-align: center; background-color: #FFFFFF; box-shadow: 0 4px 10px rgba(0,0,0,0.1); border: 1px solid #BDC3C7; height: 100%; transition: transform 0.2s; }
    .kpi-card:hover { transform: translateY(-5px); }
    .kpi-gas { border-top: 8px solid var(--kpi-gas); } .kpi-diesel { border-top: 8px solid var(--kpi-diesel); }
    .kpi-total { border-top: 8px solid var(--kpi-total); } .kpi-co2 { border-top: 8px solid var(--kpi-co2); }
    .kpi-title { font-size: 1.2rem; font-weight: bold; opacity: 0.8; color: var(--text-sub) !important; margin-bottom: 5px; }
    .kpi-value { font-size: 2.8rem; font-weight: 800; color: var(--text-main) !important; margin: 0; }
    .kpi-unit { font-size: 1rem; font-weight: normal; color: var(--text-sub) !important; margin-left: 5px; }
    .kpi-sub { font-size: 0.9rem; color: #C0392B !important; font-weight: 700; background-color: rgba(192, 57, 43, 0.1); padding: 2px 10px; border-radius: 20px; display: inline-block; margin-top: 5px;}

    /* Admin 儀表板 KPI */
    .admin-kpi-card { background-color: #FFFFFF; border: 1px solid #BDC3C7; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.1); height: 100%; text-align: center; margin-bottom: 20px; }
    .admin-kpi-header { padding: 10px; font-size: 1.2rem; font-weight: bold; color: #2C3E50; border-bottom: 1px solid rgba(0,0,0,0.1); }
    .admin-kpi-body { padding: 20px; }
    .admin-kpi-value { font-size: 2.8rem; font-weight: 900; color: #2C3E50; margin-bottom: 5px; }
    .admin-kpi-unit { font-size: 1rem; color: #7F8C8D; font-weight: normal; margin-left: 5px; }
    .admin-kpi-sub { font-size: 0.9rem; display: inline-block; padding: 2px 10px; border-radius: 15px; background-color: #F9E79F; color: #7D6608; margin-top: 5px; font-weight: bold; }

    /* 其他 */
    .unreported-block { padding: 15px 20px; border-radius: 12px; margin-bottom: 20px; color: #2C3E50; box-shadow: 0 2px 6px rgba(0,0,0,0.08); border: 1px solid rgba(0,0,0,0.05); }
    .unreported-title { font-size: 1.6rem; font-weight: 900; margin-bottom: 12px; border-bottom: 2px solid rgba(0,0,0,0.1); padding-bottom: 8px; }
    .device-info-box { background-color: #FFFFFF; border: 2px solid #5DADE2; border-radius: 10px; padding: 20px; margin-bottom: 20px; }
    .alert-box { background-color: #FCF3CF; border: 2px solid #F1C40F; padding: 15px; border-radius: 10px; margin-bottom: 20px; color: #9A7D0A !important; font-weight: bold; text-align: center; }
    .privacy-box { background-color: #F8F9F9; border: 1px solid #BDC3C7; padding: 15px; border-radius: 10px; font-size: 0.9rem; color: #566573; margin-bottom: 10px; }
    .privacy-title { font-weight: bold; color: #2C3E50; margin-bottom: 5px; font-size: 1rem; }
    .dashboard-main-title { font-size: 1.8rem; font-weight: 900; text-align: center; color: #2C3E50; margin-bottom: 20px; background-color: #F8F9F9; padding: 10px; border-radius: 10px; border: 1px solid #BDC3C7; }
    .morandi-header { background-color: #EBF5FB; color: #2E4053; padding: 15px; border-radius: 8px; border-left: 8px solid #5499C7; font-size: 1.35rem; font-weight: 700; margin-top: 25px; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    
    /* 橫式資訊卡 (Info Card) */
    .horizontal-card { display: flex; border: 1px solid #BDC3C7; border-radius: 12px; overflow: hidden; margin-bottom: 25px; box-shadow: 0 4px 8px rgba(0,0,0,0.08); background-color: #FFFFFF; min-height: 250px; }
    .card-left { flex: 3; background-color: var(--morandi-blue); color: #FFFFFF; display: flex; flex-direction: column; justify-content: center; align-items: center; padding: 20px; text-align: center; border-right: 1px solid #2C3E50; }
    .dept-text { font-size: 1.6rem; font-weight: 700; margin-bottom: 8px; line-height: 1.4; }
    .card-right { flex: 7; padding: 20px 30px; display: flex; flex-direction: column; justify-content: center; }
    .info-row { display: flex; align-items: flex-start; padding: 10px 0; font-size: 1rem; color: #566573; border-bottom: 1px dashed #F2F3F4; }
    .info-row:last-child { border-bottom: none; }
    .info-icon { margin-right: 12px; font-size: 1.1rem; width: 25px; text-align: center; margin-top: 2px; }
    .info-label { font-weight: 700; margin-right: 10px; min-width: 150px; color: #2E4053; }
    .info-value { font-weight: 500; color: #17202A; flex: 1; line-height: 1.5; }
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

try:
    _raw_creds = st.secrets["credentials"]
    credentials_login = clean_secrets(_raw_creds)
    cookie_cfg = st.secrets["cookie"]
    authenticator = stauth.Authenticate(credentials_login, cookie_cfg["name"], cookie_cfg["key"], cookie_cfg["expiry_days"])
except:
    pass

with st.sidebar:
    st.header(f"👤 {name}")
    st.caption(f"帳號: {username}")
    st.success("☁️ 雲端連線正常")
    if username == 'admin':
        st.info("👑 管理員權限已啟用")
    st.markdown("---")
    authenticator.logout('登出系統', 'sidebar')

# ==========================================
# 3. 資料庫連線與設定
# ==========================================
SHEET_ID = "1gqDU21YJeBoBOd8rMYzwwZ45offXWPGEODKTF6B8k-Y" 
DRIVE_FOLDER_ID = "1Uryuk3-9FHJ39w5Uo8FYxuh9VOFndeqD"
VIP_UNITS = ["總務處事務組", "民雄總務", "新民聯辦", "產推處產學營運組"]
FLEET_CARDS = {"總務處事務組-柴油": "TZI510508", "總務處事務組-汽油": "TZI510509", "民雄總務": "TZI510594", "新民聯辦": "TZI510410", "產推處產學營運組": "TZI510244"}
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
            
    if len(ws_record.get_all_values()) == 0: ws_record.append_row(["填報時間", "填報單位", "填報人", "填報人分機", "設備名稱備註", "校內財產編號", "原燃物料名稱", "油卡編號", "加油日期", "加油量", "與其他設備共用加油單", "備註", "佐證資料"])
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

# 初始化 Session
if 'multi_row_count' not in st.session_state: st.session_state['multi_row_count'] = 1
if 'reset_counter' not in st.session_state: st.session_state['reset_counter'] = 0

# ==========================================
# 4. 功能函式 (將介面封裝)
# ==========================================

def render_user_interface():
    """ 一般使用者 / 管理員的前台填報介面 """
    st.markdown("### ⛽ 燃油設備填報專區")
    tabs = st.tabs(["📝 新增填報", "📊 動態查詢看板"])
    
    # --- Tab 1: 填報 ---
    with tabs[0]:
        st.markdown('<div class="alert-box">📢 請「誠實申報」，以保障單位及自身權益！</div>', unsafe_allow_html=True)
        if not df_equip.empty:
            st.markdown("#### 步驟 1：請選擇您的單位及設備")
            c1, c2 = st.columns(2)
            units = sorted([x for x in df_equip['填報單位'].unique() if x != '-' and x != '填報單位'])
            selected_dept = c1.selectbox("填報單位", units, index=None, placeholder="請選擇單位...", key="dept_selector")
            
            shared_note_text = "如有與其他設備共用油單，請於備註區備註是與哪個設備共用，謝謝。<br>"
            typo_note = f'<div class="correction-note"><span style="color:#C0392B; font-weight:900;">{shared_note_text}</span>如有資料誤繕情形，請重新登錄1次資訊，並於備註欄填寫：「前筆資料誤繕，請刪除。」，管理單位將協助刪除誤打資訊</div>'
            typo_note_simple = '<div class="correction-note">如有資料誤繕情形，請重新登錄1次資訊，並於備註欄填寫：「前筆資料誤繕，請刪除。」，管理單位將協助刪除誤打資訊</div>'
            privacy_html = """<div class="privacy-box"><div class="privacy-title">📜 個人資料蒐集、處理及利用告知聲明</div>1. <strong>蒐集機關</strong>：國立嘉義大學。<br>2. <strong>蒐集目的</strong>：進行本校公務車輛/機具之加油紀錄管理、校園溫室氣體（碳）盤查統計、稽核佐證資料蒐集及後續能源使用分析。<br>3. <strong>個資類別</strong>：填報人姓名。<br>4. <strong>利用期間</strong>：姓名保留至填報年度後第二年1月1日，期滿即進行「去識別化」刪除，其餘數據永久保存。<br>5. <strong>利用對象</strong>：本校教師、行政人員及碳盤查查驗人員。<br>6. <strong>您有權依個資法請求查詢、更正或刪除您的個資。如不提供，將無法完成填報。</strong><br></div>"""
            
            # 批次申報
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
                            c_card, c_val = st.columns([7, 3]) 
                            with c_card:
                                header_color = MORANDI_COLORS.get(row.get('統計類別'), '#D5DBDB')
                                st.markdown(f"""<div class="batch-card-final"><div class="batch-header-final" style="background-color: {header_color};"><span class="batch-title-text">⛽ {row['設備名稱備註']}</span><span class="batch-qty-badge">數量: {row.get('設備數量','-')}</span></div><div class="batch-body-final"><div class="batch-row"><div class="batch-item">🏢 部門: {row.get('設備所屬單位/部門','-')}</div><div class="batch-item">👤 保管人: {row.get('保管人','-')}</div></div><div class="batch-row"><div class="batch-item">⛽ 燃料: {row.get('原燃物料名稱')}</div><div class="batch-item">🔢 財產編號: {row.get('校內財產編號','-')}</div></div></div></div>""", unsafe_allow_html=True)
                            with c_val:
                                st.write(""); st.write("") 
                                vol = st.number_input(f"加油量", min_value=0.0, step=0.1, key=f"b_v_{row['校內財產編號']}_{idx}", label_visibility="collapsed")
                                batch_inputs[idx] = vol
                        st.markdown("---"); st.markdown("**📂 上傳中油加油明細 (只需一份)**")
                        f_file = st.file_uploader("支援 PDF/JPG/PNG", type=['pdf', 'jpg', 'png', 'jpeg'])
                        
                        st.text_input("備註", key="batch_note", placeholder="備註 (選填)", label_visibility="collapsed")
                        st.write("") 
                        st.markdown(typo_note_simple, unsafe_allow_html=True); st.markdown(privacy_html, unsafe_allow_html=True)
                        
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
                                    fuel_rep = filtered_equip.iloc[0]['原燃物料名稱'] if not filtered_equip.empty else "混合油品"
                                    clean_name = f"{selected_dept}_{target_sub_cat}_{fuel_rep}_{total_vol}.{file_ext}"
                                    file_meta = {'name': clean_name, 'parents': [DRIVE_FOLDER_ID]}
                                    media = MediaIoBaseUpload(f_file, mimetype=f_file.type, resumable=True)
                                    file = drive_service.files().create(body=file_meta, media_body=media, fields='webViewLink').execute()
                                    file_link = file.get('webViewLink')
                                    fleet_id = "-"; 
                                    if selected_dept == "總務處事務組": fleet_id = FLEET_CARDS.get(f"總務處事務組-{'汽油' if '汽油' in target_sub_cat else '柴油'}", "-")
                                    else: fleet_id = FLEET_CARDS.get(selected_dept, "-")
                                    rows_to_append = []
                                    current_time = get_taiwan_time().strftime("%Y-%m-%d %H:%M:%S")
                                    note_val = st.session_state.get("batch_note", "")
                                    for idx, vol in batch_inputs.items():
                                        row = filtered_equip.loc[idx]
                                        rows_to_append.append([current_time, selected_dept, p_name, p_ext, row['設備名稱備註'], str(row.get('校內財產編號','-')), row['原燃物料名稱'], fleet_id, str(batch_date), vol, "是", f"批次申報-{target_sub_cat} | {note_val}", file_link])
                                    if rows_to_append: ws_record.append_rows(rows_to_append); st.success(f"✅ 批次申報成功！已寫入 {len(rows_to_append)} 筆紀錄。"); st.balloons(); st.session_state['reset_counter'] += 1
                                    else: st.warning("系統錯誤：無法產生寫入資料。")
                                except Exception as e: st.error(f"失敗: {e}")
            else:
                # 一般單筆申報
                filtered = df_equip[df_equip['填報單位'] == selected_dept]
                devices = sorted([x for x in filtered['設備名稱備註'].unique()])
                dynamic_key = f"vehicle_selector_{st.session_state['reset_counter']}"
                selected_device = c2.selectbox("車輛/機具名稱", devices, index=None, placeholder="請選擇車輛...", key=dynamic_key)
                if selected_device:
                    row = filtered[filtered['設備名稱備註'] == selected_device].iloc[0]
                    info_html = f"""<div class="device-info-box"><div style="border-bottom: 1px solid #BDC3C7; padding-bottom: 10px; margin-bottom: 10px; font-weight: bold; font-size: 1.2rem; color: #5DADE2;">📋 設備詳細資料</div><div><strong>🏢 部門：</strong>{row.get('設備所屬單位/部門', '-')}</div><div><strong>👤 保管人：</strong>{row.get('保管人', '-')}</div><div><strong>🔢 財產編號：</strong>{row.get('校內財產編號', '-')}</div><div><strong>📍 位置：</strong>{row.get('設備詳細位置/樓層', '-')}</div><div><strong>⛽ 燃料：</strong>{row.get('原燃物料名稱', '-')}</div><div><strong>📊 數量：</strong>{row.get('設備數量', '-')}</div></div>"""
                    st.markdown(info_html, unsafe_allow_html=True)
                    st.markdown("#### 步驟2：填報設備加油資訊")
                    
                    st.write("") 
                    st.markdown('<p style="color:#566573; font-size:1rem; font-weight:bold; margin-bottom:-10px;">請選擇申報類型：</p>', unsafe_allow_html=True)
                    st.write("") 
                    report_mode = st.radio("類型選擇", ["用油量申報 (含單筆/多筆/油卡)", "無使用"], horizontal=True, label_visibility="collapsed")
                    
                    if report_mode == "用油量申報 (含單筆/多筆/油卡)":
                        c_btn1, c_btn2, _ = st.columns([1, 1, 3])
                        with c_btn1: 
                            if st.button("➕ 增加一列"): st.session_state['multi_row_count'] += 1
                        with c_btn2: 
                            if st.button("➖ 減少一列") and st.session_state['multi_row_count'] > 1: st.session_state['multi_row_count'] -= 1
                        st.markdown('<p style="color:#566573; font-size:0.9rem;">填報前請先設定申報筆數，至多10筆</p>', unsafe_allow_html=True)

                    with st.form("entry_form", clear_on_submit=True):
                        col_p1, col_p2 = st.columns(2)
                        p_name = col_p1.text_input("👤 填報人姓名 (必填)")
                        p_ext = col_p2.text_input("📞 聯絡分機 (必填)")
                        fuel_card_id = ""; data_entries = []; f_files = None; note_input = ""
                        
                        if report_mode == "用油量申報 (含單筆/多筆/油卡)":
                            fuel_card_id = st.text_input("💳 油卡編號 (選填)")
                            for i in range(st.session_state['multi_row_count']):
                                c_d, c_v = st.columns(2)
                                _date = c_d.date_input(f"📅 加油日期 填報序號 {i+1}", datetime.today(), key=f"d_{i}")
                                _vol = c_v.number_input(f"💧 加油量(公升) 填報序號 {i+1}", min_value=0.0, step=0.1, key=f"v_{i}")
                                data_entries.append({"date": _date, "vol": _vol})
                            st.markdown("---"); is_shared = st.checkbox("與其他設備共用加油單"); 
                            st.markdown("<h4 style='color: #1A5276;'>📂 上傳佐證資料 (必填)</h4>", unsafe_allow_html=True)
                            st.markdown("""* **A. 請依填報加油日期之順序上傳檔案。**\n* **B. 一次多筆申報時，可採單張油單逐一按時序上傳，或依時序彙整成一個檔案後統一上傳。**\n* **C. 支援 png, jpg, jpeg, pdf (單檔最多3MB，最多可上傳10個檔案)。**""")
                            f_files = st.file_uploader("選擇檔案", type=['png', 'jpg', 'jpeg', 'pdf'], accept_multiple_files=True)
                            
                            note_input = st.text_input("備註", placeholder="備註 (選填)")
                            st.markdown(typo_note, unsafe_allow_html=True)
                            
                        else:
                            st.info("ℹ️ 您選擇了「無使用」，請選擇無使用的期間。")
                            c_s, c_e = st.columns(2)
                            d_start = c_s.date_input("開始日期", datetime(datetime.now().year, 1, 1))
                            d_end = c_e.date_input("結束日期", datetime.now())
                            data_entries.append({"date": d_end, "vol": 0.0})
                            note_input = f"無使用 (期間: {d_start} ~ {d_end})"
                            
                            note_ext = st.text_input("備註", placeholder="備註 (選填)")
                            if note_ext: note_input += f" | {note_ext}"
                            st.markdown(typo_note_simple, unsafe_allow_html=True)
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
                                    valid_logic = True; links=[]
                                    if f_files:
                                        total_report_vol = sum([e['vol'] for e in data_entries])
                                        fuel_type = row.get('原燃物料名稱', '未知油品')
                                        shared_tag = "(共用)" if is_shared else ""
                                        for idx, f in enumerate(f_files):
                                            try:
                                                f.seek(0); file_ext = f.name.split('.')[-1]; clean_name = ""
                                                if len(f_files) == len(data_entries): c_date = data_entries[idx]['date']; c_vol = data_entries[idx]['vol']; clean_name = f"{selected_dept}_{selected_device}_{fuel_type}_{c_date}_{c_vol}{shared_tag}.{file_ext}"
                                                elif len(f_files) == 1 and len(data_entries) > 1: clean_name = f"{selected_dept}_{selected_device}_{fuel_type}_{total_report_vol}{shared_tag}.{file_ext}"
                                                else: clean_name = f"{selected_dept}_{selected_device}_{fuel_type}_{data_entries[0]['date']}_{idx+1}{shared_tag}.{file_ext}"
                                                meta = {'name': clean_name, 'parents': [DRIVE_FOLDER_ID]}
                                                media = MediaIoBaseUpload(f, mimetype=f.type, resumable=True)
                                                file = drive_service.files().create(body=meta, media_body=media, fields='webViewLink').execute()
                                                links.append(file.get('webViewLink'))
                                            except: valid_logic=False; st.error("上傳失敗"); break
                                    if valid_logic:
                                        rows = []; now_str = get_taiwan_time().strftime("%Y-%m-%d %H:%M:%S")
                                        final_link = "\n".join(links) if links else "無"
                                        shared_str = "是" if is_shared else "-"; card_str = fuel_card_id if fuel_card_id else "-"
                                        for e in data_entries: rows.append([now_str, selected_dept, p_name, p_ext, selected_device, str(row.get('校內財產編號','-')), str(row.get('原燃物料名稱','-')), card_str, str(e['date']), e['vol'], shared_str, note_input, final_link])
                                        if rows: ws_record.append_rows(rows); st.success("✅ 申報成功！"); st.balloons(); st.session_state['reset_counter'] += 1; st.cache_data.clear()
                            elif report_mode == "無使用":
                                rows = [[get_taiwan_time().strftime("%Y-%m-%d %H:%M:%S"), selected_dept, p_name, p_ext, selected_device, str(row.get('校內財產編號','-')), str(row.get('原燃物料名稱','-')), "-", str(data_entries[0]['date']), 0.0, "-", note_input, "無"]]
                                ws_record.append_rows(rows); st.success("✅ 申報成功！"); st.balloons(); st.session_state['reset_counter'] += 1; st.cache_data.clear()
        else: st.warning("📭 目前資料庫尚無有效資料，請先至「新增填報」分頁填寫。")

    # === Tab 2: 看板 (V163.0) ===
    with tabs[1]:
        st.markdown("### 📊 動態查詢看板 (年度檢視)")
        st.info("請選擇「單位」與「年份」，檢視該年度的用油統計與碳排放分析。")
        col_r1, col_r2 = st.columns([4, 1])
        with col_r2:
            if st.button("🔄 刷新數據", use_container_width=True, key="refresh_all"): st.cache_data.clear(); st.rerun()
        
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
                df_final = df_dept[df_dept['日期格式'].dt.year == query_year].copy()
                
                if not df_final.empty:
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
                    
                    # Chart 4: V160/V161 Logic
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
                    
                    monthly_counts = df_final[df_final['加油量'] > 0].groupby('月份')['油品類別'].nunique()

                    fig = go.Figure()
                    device_color_map = {dev: DASH_PALETTE[i % len(DASH_PALETTE)] for i, dev in enumerate(unique_devices)}
                    
                    for dev in unique_devices:
                        dev_data = df_final[df_final['設備名稱備註'] == dev]
                        dev_grouped = dev_data.groupby(['月份', '油品類別'])['加油量'].sum().reset_index()
                        merged_dev = pd.merge(base_x, dev_grouped, on=['月份', '油品類別'], how='left').fillna(0)
                        
                        def get_text(row):
                            val = row['加油量']
                            m = row['月份']
                            if val <= 0: return ""
                            if monthly_counts.get(m, 0) <= 1: return "" 
                            return f"{val:.1f}"

                        text_vals = merged_dev.apply(get_text, axis=1)
                        fig.add_trace(go.Bar(x=[merged_dev['月份'], merged_dev['油品類別']], y=merged_dev['加油量'], name=dev, marker_color=device_color_map[dev], text=text_vals, texttemplate='%{text}', textposition='inside', textangle=0))
                    
                    total_grouped = df_final.groupby(['月份', '油品類別'])['加油量'].sum().reset_index()
                    merged_total = pd.merge(base_x, total_grouped, on=['月份', '油品類別'], how='left').fillna(0)
                    label_data = merged_total[merged_total['加油量'] > 0]
                    fig.add_trace(go.Scatter(x=[label_data['月份'], label_data['油品類別']], y=label_data['加油量'], text=label_data['加油量'].apply(lambda x: f"{x:.1f}"), mode='text', textposition='top center', textfont=dict(size=14, color='black'), showlegend=False))

                    fig.update_layout(barmode='stack', font=dict(size=14), xaxis=dict(title="月份 / 油品"), yaxis=dict(title="加油量 (公升)"), height=550, margin=dict(t=50, b=120))
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # V163.0: 設備申報資訊統計區 (Header Yellow #FAD7A0, Unit Dark Gray, Reordered Fields)
                    st.markdown("---")
                    st.subheader(f"📋 {query_dept} - 設備申報資訊統計區", anchor=False)
                    target_devices = df_equip[df_equip['填報單位'] == query_dept]
                    if not target_devices.empty:
                        device_list = []
                        for _, row in target_devices.iterrows():
                            d_name = row['設備名稱備註']
                            d_id = row.get('設備編號', '無編號')
                            d_unit = row.get('填報單位', '-')
                            d_sub = row.get('設備所屬單位/部門', '-')
                            d_keeper = row.get('保管人', '-')
                            d_loc = row.get('設備詳細位置/樓層', '-')
                            d_qty = row.get('設備數量', '1')
                            d_prop = row.get('校內財產編號', '-')
                            raw_fuel = row.get('原燃物料名稱', '-')
                            d_fuel = '汽油' if '汽油' in raw_fuel else ('柴油' if '柴油' in raw_fuel else raw_fuel)
                            d_vol = df_final[df_final['設備名稱備註'] == d_name]['加油量'].sum()
                            d_count = len(df_final[df_final['設備名稱備註'] == d_name])
                            status_html = '<span class="alert-status">⚠️ 尚未申報</span>' if d_count == 0 else ""
                            device_list.append({ "id": d_id, "name": d_name, "vol": d_vol, "fuel": d_fuel, "unit": d_unit, "sub": d_sub, "keeper": d_keeper, "loc": d_loc, "qty": d_qty, "prop": d_prop, "count": d_count, "status": status_html })
                        
                        for k in range(0, len(device_list), 2):
                            d_cols = st.columns(2)
                            for m in range(2):
                                if k + m < len(device_list):
                                    item = device_list[k + m]
                                    with d_cols[m]:
                                        # V163: Layout adjusted to Grid 2x2. First row: Fuel, Sub. Second row: Keeper, Loc.
                                        st.markdown(f"""
                                        <div class="dev-card-v148">
                                            <div class="dev-header" style="background-color: #FAD7A0;">
                                                <div class="dev-header-left"><div class="dev-id">{item['id']}</div><div class="dev-name-row"><span class="dev-name">{item['name']}</span><span class="qty-badge">數量: {item['qty']}</span></div></div>
                                                <div class="dev-header-right"><div class="dev-vol">{item['vol']:.1f}<span class="dev-unit" style="color:#333333;">公升</span></div></div>
                                            </div>
                                            <div class="dev-body">
                                                <div class="dev-item"><span class="dev-label">燃料種類:</span><span class="dev-val">{item['fuel']}</span></div>
                                                <div class="dev-item"><span class="dev-label">設備所屬單位/部門:</span><span class="dev-val">{item['sub']}</span></div>
                                                <div class="dev-item"><span class="dev-label">保管人:</span><span class="dev-val">{item['keeper']}</span></div>
                                                <div class="dev-item"><span class="dev-label">設備詳細位置/樓層:</span><span class="dev-val">{item['loc']}</span></div>
                                            </div>
                                            <div class="dev-footer"><div class="dev-count">年度申報次數: {item['count']} 次</div><div>{item['status']}</div></div>
                                        </div>
                                        """, unsafe_allow_html=True)

                    st.markdown("---")
                    # V163.0: 油品設備佔比分析 (Fixed Height 600, Black Text)
                    st.subheader("🍩 油品設備用油量佔比分析", anchor=False)
                    
                    gas_df = df_final[(df_final['原燃物料名稱'].str.contains('汽油', na=False)) & (df_final['加油量'] > 0)]
                    if not gas_df.empty:
                        fig_gas = px.pie(gas_df, values='加油量', names='設備名稱備註', title='⛽ 汽油設備用油量分析', color_discrete_sequence=px.colors.sequential.Teal, hole=0.5)
                        pull_g = [0.1 if v < gas_df['加油量'].sum()*0.05 else 0 for v in gas_df['加油量']]
                        fig_gas.update_traces(textinfo='percent+label', textfont=dict(color='black', size=16), textposition='outside', insidetextorientation='horizontal', pull=pull_g, hovertemplate='<b>項目: %{label}</b><br>統計加油量: %{value:.1f} L<br>百分比: %{percent:.1%}<extra></extra>')
                        fig_gas.update_layout(height=600, legend=dict(orientation="h", yanchor="top", y=-0.1, xanchor="center", x=0.5), margin=dict(l=40, r=40, t=40, b=40))
                        st.plotly_chart(fig_gas, use_container_width=True)
                    else: st.info("無汽油使用紀錄")
                    
                    st.write("") 

                    diesel_df = df_final[(df_final['原燃物料名稱'].str.contains('柴油', na=False)) & (df_final['加油量'] > 0)]
                    if not diesel_df.empty:
                        fig_diesel = px.pie(diesel_df, values='加油量', names='設備名稱備註', title='🚛 柴油設備用油量分析', color_discrete_sequence=px.colors.sequential.Oranges, hole=0.5)
                        pull_d = [0.1 if v < diesel_df['加油量'].sum()*0.05 else 0 for v in diesel_df['加油量']]
                        fig_diesel.update_traces(textinfo='percent+label', textfont=dict(color='black', size=16), textposition='outside', insidetextorientation='horizontal', pull=pull_d, hovertemplate='<b>項目: %{label}</b><br>統計加油量: %{value:.1f} L<br>百分比: %{percent:.1%}<extra></extra>')
                        fig_diesel.update_layout(height=600, legend=dict(orientation="h", yanchor="top", y=-0.1, xanchor="center", x=0.5), margin=dict(l=40, r=40, t=40, b=40))
                        st.plotly_chart(fig_diesel, use_container_width=True)
                    else: st.info("無柴油使用紀錄")
                    
                    st.markdown("---")
                    # V161/V163: 碳排結構 (Height=800)
                    st.subheader(f"🌞 單位油料使用碳排放量(公噸二氧化碳當量)結構", anchor=False)
                    
                    df_final['CO2e'] = df_final.apply(lambda r: r['加油量']*0.0022 if '汽油' in str(r['原燃物料名稱']) else r['加油量']*0.0027, axis=1)
                    
                    treemap_data = df_final.groupby(['設備名稱備註'])['CO2e'].sum().reset_index()
                    fig_tree = px.treemap(treemap_data, path=['設備名稱備註'], values='CO2e', title=f"{query_dept} - 設備碳排放量權重分析", color='CO2e', color_continuous_scale='Teal')
                    fig_tree.update_traces(texttemplate='%{label}<br>%{value:.4f}<br>%{percentEntry:.1%}', textfont=dict(size=24))
                    fig_tree.update_coloraxes(showscale=False)
                    fig_tree.update_layout(height=800)
                    st.plotly_chart(fig_tree, use_container_width=True)

                    st.markdown("---")
                    st.subheader(f"📋 {query_year}年度 填報明細")
                    df_display = df_final[["加油日期", "設備名稱備註", "原燃物料名稱", "油卡編號", "加油量", "填報人", "備註"]].sort_values(by='加油日期', ascending=False).rename(columns={'加油量': '加油量(公升)'})
                    st.dataframe(df_display.style.format({"加油量(公升)": "{:.2f}"}), use_container_width=True)
                else: st.warning(f"⚠️ {query_dept} 在 {query_year} 年度尚無填報紀錄。")
        else: st.info("尚無該年度資料，無法顯示儀表板。")

    st.markdown('<div class="contact-footer">管理員系統版本 V163.0 (Fuel Final Refined - Layout & Rotation Fixed)</div>', unsafe_allow_html=True)

def render_admin_dashboard():
    """ 顯示管理員後台 """
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

    admin_tabs = st.tabs(["🔍 申報資料異動", "⚠️ 篩選未申報名單", "📝 全校燃油設備總覽", "📊 全校油料使用儀表板"])

    # === Tab 1: 申報資料異動 ===
    with admin_tabs[0]:
        st.subheader("🔍 申報資料異動")
        if not df_year.empty:
            df_year['加油日期'] = pd.to_datetime(df_year['加油日期']).dt.date
            edited = st.data_editor(df_year, column_config={"佐證資料": st.column_config.LinkColumn("佐證", display_text="🔗"), "加油日期": st.column_config.DateColumn("日期", format="YYYY-MM-DD"), "加油量": st.column_config.NumberColumn("油量", format="%.2f"), "填報時間": st.column_config.TextColumn("填報時間", disabled=True)}, num_rows="dynamic", use_container_width=True, key="editor_v122")
            
            df_stats = df_year.groupby(['設備名稱備註'])['加油量'].sum().reset_index()
            df_export = pd.merge(df_equip, df_stats, on='設備名稱備註', how='left')
            df_export['加油量'] = df_export['加油量'].fillna(0)
            df_export.rename(columns={'加油量': f'{selected_admin_year}年度總加油量'}, inplace=True)
            
            target_cols = ['填報單位', '設備名稱備註', '校內財產編號', '原燃物料名稱', '保管人', '設備所屬單位/部門', '設備詳細位置/樓層', '設備數量', '設備編號', f'{selected_admin_year}年度總加油量']
            final_cols = [c for c in target_cols if c in df_export.columns]
            df_final_export = df_export[final_cols]
            
            csv_data = df_final_export.to_csv(index=False).encode('utf-8-sig')
            st.download_button(label="📥 下載設備年度加油量統計表 (CSV)", data=csv_data, file_name=f"{selected_admin_year}_設備年度加油統計.csv", mime="text/csv", key="dl_stats_csv_v139")

            if st.button("💾 儲存變更", type="primary"):
                try:
                    df_all_data = df_records.copy()
                    df_all_data['temp_date'] = pd.to_datetime(df_all_data['加油日期'], errors='coerce')
                    df_all_data['temp_year'] = df_all_data['temp_date'].dt.year.fillna(0).astype(int)
                    df_keep = df_all_data[df_all_data['temp_year'] != selected_admin_year].copy()
                    df_new = edited.copy()
                    df_final = pd.concat([df_keep, df_new], ignore_index=True)
                    if 'temp_date' in df_final.columns: del df_final['temp_date']
                    if 'temp_year' in df_final.columns: del df_final['temp_year']
                    if '加油日期' in df_final.columns: df_final['加油日期'] = df_final['加油日期'].astype(str)
                    cols_to_write = df_records.columns.tolist()
                    df_final = df_final[cols_to_write]
                    df_final = df_final.sort_values(by='加油日期', ascending=False)

                    ws_record.clear()
                    ws_record.update([df_final.columns.tolist()] + df_final.astype(str).values.tolist())
                    st.success("✅ 更新成功！資料已安全合併存檔。"); st.cache_data.clear(); time.sleep(1); st.rerun()
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

    # === Tab 3: 全校總覽 ===
    with admin_tabs[2]:
        if not df_year.empty and not df_equip.empty:
            total_eq = int(df_equip['設備數量_num'].sum())
            gas_eq = int(df_equip[df_equip['原燃物料名稱'].str.contains('汽油', na=False)]['設備數量_num'].sum())
            diesel_eq = int(df_equip[df_equip['原燃物料名稱'].str.contains('柴油', na=False)]['設備數量_num'].sum())
            k1, k2, k3 = st.columns(3)
            k1.markdown(f"""<div class="top-kpi-card"><div class="top-kpi-title">🚜 全校燃油設備總數</div><div class="top-kpi-value">{total_eq}</div></div>""", unsafe_allow_html=True)
            k2.markdown(f"""<div class="top-kpi-card"><div class="top-kpi-title">⛽ 全校汽油設備數</div><div class="top-kpi-value">{gas_eq}</div></div>""", unsafe_allow_html=True)
            k3.markdown(f"""<div class="top-kpi-card"><div class="top-kpi-title">🚛 全校柴油設備數</div><div class="top-kpi-value">{diesel_eq}</div></div>""", unsafe_allow_html=True)
            st.markdown("---")

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
                            count_tot = int(eq_sums.get(category, 0))
                            count_gas = int(eq_gas_sums.get(category, 0))
                            count_dsl = int(eq_dsl_sums.get(category, 0))
                            gas_vol = fuel_sums.loc[category, '汽油'] if category in fuel_sums.index and '汽油' in fuel_sums.columns else 0
                            diesel_vol = fuel_sums.loc[category, '柴油'] if category in fuel_sums.index and '柴油' in fuel_sums.columns else 0
                            total_vol = gas_vol + diesel_vol
                            header_color = MORANDI_COLORS.get(category, "#CFD8DC")
                            st.markdown(f"""<div class="stat-card-v119"><div class="stat-header" style="background-color: {header_color};"><span class="stat-title">{category}</span><span class="stat-count">{count_tot}</span></div><div class="stat-body-split"><div class="stat-col-left"><div class="stat-item"><span class="stat-item-label">⛽ 汽油設備數</span><span class="stat-item-val">{count_gas}</span></div><div class="stat-item"><span class="stat-item-label">🚛 柴油設備數</span><span class="stat-item-val">{count_dsl}</span></div><div class="stat-item"><span class="stat-item-label">🔥 燃油設備數</span><span class="stat-item-val">{count_tot}</span></div></div><div class="stat-col-right"><div class="stat-item"><span class="stat-item-label">汽油加油量(公升)</span><span class="stat-item-val">{gas_vol:,.1f}</span></div><div class="stat-item"><span class="stat-item-label">柴油加油量(公升)</span><span class="stat-item-val">{diesel_vol:,.1f}</span></div><div class="stat-item"><span class="stat-item-label">總計加油量(公升)</span><span class="stat-item-val">{total_vol:,.1f}</span></div></div></div></div>""", unsafe_allow_html=True)
            
            st.markdown("---")
            st.subheader("📋 各類設備詳細申報紀錄一覽")
            
            for category in DEVICE_ORDER:
                target_devices = df_equip[df_equip['統計類別'] == category]
                if not target_devices.empty:
                    st.markdown(f"<h3 style='color: #2874A6; font-size: 1.6rem; margin-top: 40px; margin-bottom: 25px;'>{category}</h3>", unsafe_allow_html=True)
                    
                    device_list = []
                    for _, row in target_devices.iterrows():
                        d_name = row['設備名稱備註']
                        d_id = row.get('設備編號', '無編號')
                        d_unit = row.get('填報單位', '-')
                        d_sub = row.get('設備所屬單位/部門', '-')
                        d_keeper = row.get('保管人', '-')
                        d_qty = row.get('設備數量', '1')
                        raw_fuel = row.get('原燃物料名稱', '-')
                        d_fuel = '汽油' if '汽油' in raw_fuel else ('柴油' if '柴油' in raw_fuel else raw_fuel)
                        d_vol = df_year[df_year['設備名稱備註'] == d_name]['加油量'].sum()
                        d_count = len(df_year[df_year['設備名稱備註'] == d_name])
                        status_html = '<span class="alert-status">⚠️ 尚未申報</span>' if d_count == 0 else ""

                        device_list.append({ "id": d_id, "name": d_name, "vol": d_vol, "fuel": d_fuel, "unit": d_unit, "sub": d_sub, "keeper": d_keeper, "qty": d_qty, "count": d_count, "status": status_html })
                    
                    for k in range(0, len(device_list), 2):
                        d_cols = st.columns(2)
                        for m in range(2):
                            if k + m < len(device_list):
                                item = device_list[k + m]
                                with d_cols[m]:
                                    # V163: Revert to Grid (CSS .dev-body is grid 1fr 1fr)
                                    st.markdown(f"""<div class="dev-card-v148"><div class="dev-header" style="background-color: {MORANDI_COLORS.get(category, '#34495E')};"><div class="dev-header-left"><div class="dev-id">{item['id']}</div><div class="dev-name-row"><span class="dev-name">{item['name']}</span><span class="qty-badge">數量: {item['qty']}</span></div></div><div class="dev-header-right"><div class="dev-vol">{item['vol']:.1f}<span class="dev-unit">公升</span></div></div></div><div class="dev-body"><div class="dev-item"><span class="dev-label">燃料種類:</span><span class="dev-val">{item['fuel']}</span></div><div class="dev-item"><span class="dev-label">填報單位:</span><span class="dev-val">{item['unit']}</span></div><div class="dev-item"><span class="dev-label">所屬部門:</span><span class="dev-val">{item['sub']}</span></div><div class="dev-item"><span class="dev-label">保管人:</span><span class="dev-val">{item['keeper']}</span></div></div><div class="dev-footer"><div class="dev-count">年度申報次數: {item['count']} 次</div><div>{item['status']}</div></div></div>""", unsafe_allow_html=True)

            st.markdown("---")
            st.subheader("🍩 油品設備用油量佔比分析")
            color_map = { "公務車輛(GV-1-)": "#B0C4DE", "乘坐式割草機(GV-2-)": "#F5CBA7", "乘坐式農用機具(GV-3-)": "#D7BDE2", "鍋爐(GS-1-)": "#E6B0AA", "發電機(GS-2-)": "#A9CCE3", "肩背或手持式割草機、吹葉機(GS-3-)": "#A3E4D7", "肩背或手持式農用機具(GS-4-)": "#F9E79F" }
            
            # V163: Rotation 210, Black Text, 1 decimal
            gas_data = df_year[(df_year['油品大類'] == '汽油') & (df_year['統計類別'].isin(DEVICE_ORDER))].groupby('統計類別')['加油量'].sum().reset_index()
            if not gas_data.empty:
                gas_data = gas_data[gas_data['加油量'] > 0]
                st.markdown('<div class="pie-chart-box">', unsafe_allow_html=True)
                fig_g = px.pie(gas_data, values='加油量', names='統計類別', title='⛽ 汽油用量佔比', hole=0.4, color='統計類別', color_discrete_map=color_map)
                fig_g.update_layout(height=650, font=dict(size=18), legend=dict(font=dict(size=16)), margin=dict(l=80, r=80, t=50, b=50))
                pull_g = [0.1 if v < gas_data['加油量'].sum()*0.05 else 0 for v in gas_data['加油量']]
                # V163: Rotation 210
                fig_g.update_traces(textinfo='percent+label', textfont=dict(size=16, color='black'), textposition='auto', insidetextorientation='horizontal', pull=pull_g, hovertemplate='<b>項目: %{label}</b><br>統計加油量: %{value:.1f} L<br>百分比: %{percent:.1%}<extra></extra>', rotation=210)
                st.plotly_chart(fig_g, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)
            else: st.info("無汽油數據")
            
            dsl_data = df_year[(df_year['油品大類'] == '柴油') & (df_year['統計類別'].isin(DEVICE_ORDER))].groupby('統計類別')['加油量'].sum().reset_index()
            if not dsl_data.empty:
                dsl_data = dsl_data[dsl_data['加油量'] > 0]
                st.markdown('<div class="pie-chart-box">', unsafe_allow_html=True)
                fig_d = px.pie(dsl_data, values='加油量', names='統計類別', title='🚛 柴油用量佔比', hole=0.4, color='統計類別', color_discrete_map=color_map)
                fig_d.update_layout(height=650, font=dict(size=18), legend=dict(font=dict(size=16)), margin=dict(l=80, r=80, t=50, b=50))
                pull_d = [0.1 if v < dsl_data['加油量'].sum()*0.05 else 0 for v in dsl_data['加油量']]
                # V163: Rotation 210
                fig_d.update_traces(textinfo='percent+label', textfont=dict(size=16, color='black'), textposition='auto', insidetextorientation='horizontal', pull=pull_d, hovertemplate='<b>項目: %{label}</b><br>統計加油量: %{value:.1f} L<br>百分比: %{percent:.1%}<extra></extra>', rotation=210)
                st.plotly_chart(fig_d, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)
            else: st.info("無柴油數據")
        else: st.warning("尚無資料可供統計。")

    # === Tab 4: 儀表板 (V163 Fix) ===
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

            # Chart 4: V163 Fix (Black Text)
            st.subheader("📈 全校逐月加油量統計")
            monthly = df_year.groupby(['月份', '油品大類'])['加油量'].sum().reset_index()
            full_months = pd.DataFrame({'月份': range(1, 13)})
            monthly = full_months.merge(monthly, on='月份', how='left').fillna({'加油量':0, '油品大類':'汽油'})
            
            fig_month = px.bar(monthly, x='月份', y='加油量', color='油品大類', barmode='group', text_auto='.2f', color_discrete_sequence=DASH_PALETTE)
            fig_month.update_layout(xaxis=dict(tickmode='linear', tick0=1, dtick=1, title_font=dict(size=20), tickfont=dict(size=18, color='#566573')), yaxis=dict(title="加油量(公升)", title_font=dict(size=20), tickfont=dict(size=18, color='#566573')), font=dict(size=18), showlegend=True, height=500, margin=dict(t=50))
            # V163: Black text
            fig_month.update_traces(textfont=dict(color='black', size=14), textangle=0, textposition='outside')
            st.plotly_chart(fig_month, use_container_width=True)

            st.markdown("---")

            # Chart 5: V163 Fix (Black Text)
            st.subheader("🏆 全校前十大加油量單位")
            top_fuel = st.radio("選擇油品類型", ["汽油", "柴油"], horizontal=True, label_visibility="collapsed")
            df_top = df_year[df_year['油品大類'] == top_fuel]
            if not df_top.empty:
                top10_data = df_top.groupby('填報單位')['加油量'].sum().nlargest(10).reset_index()
                fig_top = px.bar(top10_data, x='填報單位', y='加油量', text_auto='.2f', title=f"{top_fuel}用量前十大單位", color_discrete_sequence=DASH_PALETTE)
                fig_top.update_layout(xaxis=dict(categoryorder='total descending', title_font=dict(size=20), tickfont=dict(size=18, color='#566573')), yaxis=dict(title="加油量(公升)", title_font=dict(size=20), tickfont=dict(size=18, color='#566573')), font=dict(size=18), height=600, margin=dict(t=50))
                # V163: Black text
                fig_top.update_traces(selector=dict(type='bar'), width=0.5, textposition='outside', textangle=0, textfont=dict(color='black', size=14))
                st.plotly_chart(fig_top, use_container_width=True)
            else: st.info("無此油品數據。")

            st.markdown("---")
            # Chart 6: V163 Fix (Rotation 210, Black Text, 1 decimal)
            st.subheader("🍩 全校加油量單位佔比")
            
            df_gas = df_year[(df_year['油品大類'] == '汽油') & (df_year['加油量'] > 0)]
            if not df_gas.empty:
                pull_dg = [0.1 if v < df_gas['加油量'].sum()*0.05 else 0 for v in df_gas['加油量']]
                fig_dg = px.pie(df_gas, values='加油量', names='填報單位', title='⛽ 汽油用量分佈', hole=0.4, color_discrete_sequence=DASH_PALETTE)
                fig_dg.update_layout(height=650, font=dict(size=18), legend=dict(font=dict(size=16)), margin=dict(t=80, l=100, r=100, b=40))
                # V163: Rotation 210
                fig_dg.update_traces(textposition='outside', textinfo='label+percent', hovertemplate='<b>項目: %{label}</b><br>統計加油量: %{value:.2f} L<br>百分比: %{percent:.1%}<extra></extra>', textfont=dict(size=16, color='black'), insidetextorientation='horizontal', pull=pull_dg, rotation=210)
                st.plotly_chart(fig_dg, use_container_width=True)
            else: st.info("無汽油數據")
            
            df_dsl = df_year[(df_year['油品大類'] == '柴油') & (df_year['加油量'] > 0)]
            if not df_dsl.empty:
                pull_dd = [0.1 if v < df_dsl['加油量'].sum()*0.05 else 0 for v in df_dsl['加油量']]
                fig_dd = px.pie(df_dsl, values='加油量', names='填報單位', title='🚛 柴油用量分佈', hole=0.4, color_discrete_sequence=DASH_PALETTE)
                fig_dd.update_layout(height=650, font=dict(size=18), legend=dict(font=dict(size=16)), margin=dict(t=80, l=100, r=100, b=40))
                # V163: Rotation 210
                fig_dd.update_traces(textposition='outside', textinfo='label+percent', hovertemplate='<b>項目: %{label}</b><br>統計加油量: %{value:.2f} L<br>百分比: %{percent:.1%}<extra></extra>', textfont=dict(size=16, color='black'), insidetextorientation='horizontal', pull=pull_dd, rotation=210)
                st.plotly_chart(fig_dd, use_container_width=True)
            else: st.info("無柴油數據")

            st.markdown("---")
            # Chart 7: V158 Fix (Height=700)
            st.subheader("🌍 全校油料使用碳排放量(公噸二氧化碳當量)結構")
            df_year['CO2e'] = df_year.apply(lambda r: r['加油量']*0.0022 if '汽油' in str(r['原燃物料名稱']) else r['加油量']*0.0027, axis=1)
            if not df_year.empty:
                fig_tree = px.treemap(df_year, path=['填報單位', '設備名稱備註'], values='CO2e', color='填報單位', color_discrete_sequence=DASH_PALETTE)
                fig_tree.update_traces(texttemplate='%{label}<br>%{value:.4f}<br>%{percentRoot:.1%}', textfont=dict(size=24))
                fig_tree.update_layout(height=700)
                st.plotly_chart(fig_tree, use_container_width=True)
            else: st.info("無數據")
        else: st.info("尚無該年度資料，無法顯示儀表板。")

    st.markdown('<div class="contact-footer">管理員系統版本 V163.0 (Fuel Final Refined - Layout & Rotation Fixed)</div>', unsafe_allow_html=True)

# ==========================================
# 5. 主程式入口
# ==========================================
if __name__ == "__main__":
    if username == 'admin':
        main_tabs = st.tabs(["👑 管理員後台", "⛽ 外部填報系統"])
        with main_tabs[0]: render_admin_dashboard()
        with main_tabs[1]: render_user_interface()
    else: render_user_interface()