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

def get_taiwan_time():
    return datetime.utcnow() + timedelta(hours=8)

# ==========================================
# 1. CSS 樣式表 (V106: 針對批次卡片與統計卡片優化)
# ==========================================
st.markdown("""
<style>
    :root {
        color-scheme: light;
        --btn-bg: #B0BEC5; --btn-border: #2C3E50; --btn-text: #17202A;      
        --orange-bg: #E67E22; --orange-dark: #D35400; --orange-text: #FFFFFF;
        --bg-color: #EAEDED; --card-bg: #FFFFFF; --text-main: #2C3E50;
        --border-color: #BDC3C7; --morandi-red: #A93226; 
        
        /* KPI Colors */
        --kpi-gas: #52BE80; --kpi-diesel: #F4D03F; --kpi-total: #5DADE2; --kpi-co2: #AF7AC5;
    }

    [data-testid="stAppViewContainer"] { background-color: var(--bg-color); color: var(--text-main); }
    [data-testid="stHeader"] { background-color: rgba(0,0,0,0); }
    [data-testid="stSidebar"] { background-color: var(--card-bg); border-right: 1px solid var(--border-color); }
    h1, h2, h3, h4, h5, h6, p, span, label, div { color: var(--text-main); }

    /* 輸入元件 */
    div[data-baseweb="input"] > div, div[data-baseweb="base-input"] > input, textarea, input {
        background-color: #FFFFFF !important; border-color: #BDC3C7 !important; color: #000000 !important; font-size: 1.15rem !important;
    }
    div[data-baseweb="select"] > div { border-color: #BDC3C7 !important; background-color: #FFFFFF !important; }
    ul[data-baseweb="menu"] { background-color: #FFFFFF !important; }

    /* 按鈕 */
    div.stButton > button, button[kind="primary"], [data-testid="stFormSubmitButton"] > button {
        background-color: var(--orange-bg) !important; color: var(--orange-text) !important;
        border: 2px solid var(--orange-dark) !important; border-radius: 12px !important;
        font-size: 1.3rem !important; font-weight: 800 !important; padding: 0.7rem 1.5rem !important;
        box-shadow: 0 4px 6px rgba(230, 126, 34, 0.3) !important;
    }
    div.stButton > button:hover { transform: translateY(-2px); }

    /* --- V106: 外部填報 - 批次申報卡片 (7:3 版面專用) --- */
    .batch-card-final {
        background-color: #FFFFFF;
        border: 1px solid #BDC3C7;
        border-radius: 10px;
        overflow: hidden;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        height: 100%;
        display: flex;
        flex-direction: column;
        border-left: 5px solid #E67E22;
        margin-bottom: 5px; /* 微調間距 */
    }
    .batch-header-final {
        padding: 10px 15px;
        font-weight: 800;
        color: #2C3E50;
        display: flex;
        justify-content: space-between;
        align-items: center;
        border-bottom: 1px solid rgba(0,0,0,0.1);
        font-size: 1.15rem;
    }
    .batch-qty-badge {
        font-size: 0.95rem;
        background-color: rgba(255,255,255,0.6);
        padding: 2px 10px;
        border-radius: 12px;
        border: 1px solid rgba(0,0,0,0.1);
        color: #2C3E50;
    }
    .batch-body-final {
        background-color: #FFFFFF;
        padding: 12px 15px;
        font-size: 1rem;
        color: #566573;
        line-height: 1.6;
        flex-grow: 1;
    }
    /* 兩列資訊排版 */
    .batch-row {
        display: flex;
        justify-content: space-between;
        margin-bottom: 5px;
    }
    .batch-item {
        flex: 1;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        margin-right: 5px;
    }

    /* --- V106: 後台 - 統計模式專用卡片 --- */
    .stat-card-v106 {
        background-color: #FFFFFF;
        border: 1px solid #BDC3C7;
        border-radius: 12px;
        overflow: hidden;
        box-shadow: 0 3px 6px rgba(0,0,0,0.08);
        margin-bottom: 15px;
        height: 100%;
    }
    .stat-header {
        padding: 12px 20px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        border-bottom: 1px solid rgba(0,0,0,0.1);
    }
    .stat-title {
        font-size: 1.15rem;
        font-weight: 800;
        color: #2C3E50;
    }
    .stat-count {
        font-size: 1.8rem;
        font-weight: 900;
        color: var(--morandi-red);
    }
    .stat-body {
        padding: 15px 20px;
        font-size: 0.95rem;
        color: #566573;
        display: flex;
        flex-direction: column;
        gap: 8px;
    }
    .stat-row {
        display: flex;
        justify-content: space-between;
        border-bottom: 1px dashed #EAEDED;
        padding-bottom: 5px;
    }
    .stat-row:last-child {
        border-bottom: none;
        padding-top: 5px;
        font-weight: bold;
        color: #2C3E50;
    }

    /* 一般設備卡片 (後台明細用) */
    .equip-card { background-color: #FFFFFF; border: 1px solid #BDC3C7; border-radius: 10px; overflow: hidden; height: 100%; margin-bottom: 10px; box-shadow: 0 2px 5px rgba(0,0,0,0.05); }
    .equip-header { padding: 10px 15px; display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #BDC3C7; }
    .equip-code { font-size: 1.05rem; font-weight: 800; color: #2C3E50; }
    .equip-name { font-size: 0.95rem; color: #455A64; font-weight: 600;}
    .equip-vol { font-size: 1.6rem; font-weight: 900; color: var(--morandi-red); line-height: 1.2;} 
    .equip-fuel-type { font-size: 0.85rem; color: #566573; font-weight: bold; background: rgba(255,255,255,0.6); padding: 2px 6px; border-radius: 4px; margin-left: 5px;}
    .equip-body { padding: 15px; display: flex; flex-direction: column; gap: 8px; }
    .equip-info { font-size: 0.9rem; line-height: 1.5; color: #34495E; }
    .equip-footer { padding: 10px 15px; border-top: 1px dashed #D7DBDD; display: flex; justify-content: space-between; align-items: center; }
    .status-warn { background-color: #FADBD8; color: #943126; border: 1px solid #F1948A; padding: 2px 8px; border-radius: 10px; font-weight: bold; font-size: 0.8rem; }
    .count-text { font-size: 0.85rem; color: #2C3E50; font-weight: 800; margin-right: 8px; }

    /* 未申報名單區塊 (色塊樣式) */
    .unreported-block { padding: 15px 20px; border-radius: 12px; margin-bottom: 20px; color: #2C3E50; box-shadow: 0 2px 6px rgba(0,0,0,0.08); border: 1px solid rgba(0,0,0,0.05); }
    .unreported-title { font-size: 1.6rem; font-weight: 900; margin-bottom: 12px; border-bottom: 2px solid rgba(0,0,0,0.1); padding-bottom: 8px; }

    /* 其他通用 */
    .device-info-box { background-color: var(--card-bg); border: 2px solid #5DADE2; border-radius: 10px; padding: 20px; margin-bottom: 20px; }
    .alert-box { background-color: #FCF3CF; border: 2px solid #F1C40F; padding: 15px; border-radius: 10px; margin-bottom: 20px; color: #9A7D0A !important; font-weight: bold; text-align: center; }
    .login-header { font-size: 2.2rem; font-weight: 800; color: var(--text-main) !important; text-align: center; margin-bottom: 20px; padding: 25px; background-color: var(--card-bg); border: 2px solid var(--border-color); border-radius: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    .privacy-box { background-color: #F8F9F9; border: 1px solid #BDC3C7; padding: 15px; border-radius: 10px; font-size: 0.9rem; color: #566573; margin-bottom: 10px; }
    .privacy-title { font-weight: bold; color: #2C3E50; margin-bottom: 5px; font-size: 1rem; }
    button[data-baseweb="tab"] div p { font-size: 1.6rem !important; font-weight: 900 !important; color: var(--text-sub); }
    button[data-baseweb="tab"][aria-selected="true"] div p { color: #E67E22 !important; border-bottom: 3px solid #E67E22; }
    input[aria-label="搜尋框"] { height: 50px !important; font-size: 1.2rem !important; }
</style>
""", unsafe_allow_html=True)

# ☁️ 設定區
SHEET_ID = "1gqDU21YJeBoBOd8rMYzwwZ45offXWPGEODKTF6B8k-Y" 
DRIVE_FOLDER_ID = "1DCmR0dXOdFBdTrgnvCYFPtNq_bGzSJeB" 
VIP_UNITS = ["總務處事務組", "民雄總務", "新民聯辦", "產推處產學營運組"]
FLEET_CARDS = {"總務處事務組-柴油": "TZI510508", "總務處事務組-汽油": "TZI510509", "民雄總務": "TZI510594", "新民聯辦": "TZI510410", "產推處產學營運組": "TZI510244"}
DEVICE_ORDER = ["公務車輛(GV-1-)", "乘坐式割草機(GV-2-)", "乘坐式農用機具(GV-3-)", "鍋爐(GS-1-)", "發電機(GS-2-)", "肩背或手持式割草機、吹葉機(GS-3-)", "肩背或手持式農用機具(GS-4-)"]
DEVICE_CODE_MAP = {"GV-1": "公務車輛(GV-1-)", "GV-2": "乘坐式割草機(GV-2-)", "GV-3": "乘坐式農用機具(GV-3-)", "GS-1": "鍋爐(GS-1-)", "GS-2": "發電機(GS-2-)", "GS-3": "肩背或手持式割草機、吹葉機(GS-3-)", "GS-4": "肩背或手持式農用機具(GS-4-)"}

# 高級莫蘭迪色卡 (解決撞色問題)
MORANDI_COLORS = {
    "公務車輛(GV-1-)": "#B0C4DE", "乘坐式割草機(GV-2-)": "#F5CBA7", "乘坐式農用機具(GV-3-)": "#D7BDE2",
    "鍋爐(GS-1-)": "#E6B0AA", "發電機(GS-2-)": "#A9CCE3",
    "肩背或手持式割草機、吹葉機(GS-3-)": "#A3E4D7", "肩背或手持式農用機具(GS-4-)": "#F9E79F"
}
UNREPORTED_COLORS = ["#D5DBDB", "#FAD7A0", "#D2B4DE", "#AED6F1", "#A3E4D7", "#F5B7B1"]

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
    c1, c2 = st.columns(2)
    with c1:
        st.info("⛽ 車輛/機具用油")
        if st.button("前往「燃油設備填報區」", use_container_width=True, type="primary"): st.session_state['current_page'] = 'fuel'; st.rerun()
    with c2:
        st.info("❄️ 冷氣/冰水主機")
        st.button("前往「冷媒類設備填報區」", use_container_width=True, disabled=True)
    if username == 'admin':
        st.markdown("---"); st.markdown("### 👑 超級管理員專區")
        if st.button("進入「管理員後台」", use_container_width=True): st.session_state['current_page'] = 'admin_dashboard'; st.rerun()
    st.markdown('<div class="contact-footer">如有填報疑問，請電洽環安中心林小姐(分機 7137)，謝謝</div>', unsafe_allow_html=True)

# ------------------------------------------
# ⛽ 外部填報區
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
            
            privacy_html = """<div class="privacy-box"><div class="privacy-title">📜 個人資料蒐集、處理及利用告知聲明</div>1. <strong>蒐集機關</strong>：國立嘉義大學。<br>2. <strong>蒐集目的</strong>：進行本校公務車輛/機具之加油紀錄管理、校園溫室氣體（碳）盤查統計、稽核佐證資料蒐集及後續能源使用分析。<br>3. <strong>個資類別</strong>：填報人姓名。<br>4. <strong>利用期間</strong>：姓名保留至填報年度後第二年1月1日，期滿即進行「去識別化」刪除，其餘數據永久保存。<br>5. <strong>利用對象</strong>：本校教師、行政人員及碳盤查查驗人員。<br>6. <strong>您有權依個資法請求查詢、更正或刪除您的個資。如不提供，將無法完成填報。</strong><br></div>"""

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
                            # V106.0: 7:3 Layout + 資訊卡修正
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
                                st.write("") # Vertical spacer
                                st.write("") 
                                vol = st.number_input(f"加油量", min_value=0.0, step=0.1, key=f"b_v_{row['校內財產編號']}_{idx}", label_visibility="collapsed")
                                batch_inputs[idx] = vol
                                
                        st.markdown("---")
                        st.markdown("**📂 上傳中油加油明細 (只需一份)**")
                        f_file = st.file_uploader("支援 PDF/JPG/PNG", type=['pdf', 'jpg', 'png', 'jpeg'])
                        st.markdown("---"); st.markdown(privacy_html, unsafe_allow_html=True)
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
                                    for idx, vol in batch_inputs.items():
                                        row = filtered_equip.loc[idx]
                                        rows_to_append.append([current_time, selected_dept, p_name, p_ext, row['設備名稱備註'], str(row.get('校內財產編號','-')), row['原燃物料名稱'], fleet_id, str(batch_date), vol, "是", f"批次申報-{target_sub_cat}", file_link])
                                    if rows_to_append:
                                        ws_record.append_rows(rows_to_append); st.success(f"✅ 批次申報成功！已寫入 {len(rows_to_append)} 筆紀錄。")
                                        st.balloons(); st.session_state['reset_counter'] += 1; time.sleep(1.5); st.rerun()
                                    else: st.warning("沒有可寫入的紀錄。")
                                except Exception as e: st.error(f"失敗: {e}")

            # --- 一般申報模式 ---
            else:
                filtered = df_equip[df_equip['填報單位'] == selected_dept]
                devices = sorted([x for x in filtered['設備名稱備註'].unique()])
                dynamic_key = f"vehicle_selector_{st.session_state['reset_counter']}"
                selected_device = c2.selectbox("車輛/機具名稱", devices, index=None, placeholder="請選擇車輛...", key=dynamic_key)
                if selected_device:
                    row = filtered[filtered['設備名稱備註'] == selected_device].iloc[0]
                    info_html = f"""<div class="device-info-box"><div style="border-bottom: 1px solid #BDC3C7; padding-bottom: 10px; margin-bottom: 10px; font-weight: bold; font-size: 1.2rem; color: #5DADE2;">📋 設備詳細資料</div><div><strong>🏢 部門：</strong>{row.get('設備所屬單位/部門', '-')}</div><div><strong>👤 保管人：</strong>{row.get('保管人', '-')}</div><div><strong>🔢 財產編號：</strong>{row.get('校內財產編號', '-')}</div><div><strong>📍 位置：</strong>{row.get('設備詳細位置/樓層', '-')}</div><div><strong>⛽ 燃料：</strong>{row.get('原燃物料名稱', '-')}</div><div><strong>📊 數量：</strong>{row.get('設備數量', '-')}</div></div>"""
                    st.markdown(info_html, unsafe_allow_html=True)
                    st.markdown("#### 步驟 2：填寫資料")
                    report_mode = st.radio("請選擇申報類型", ["用油量申報 (含單筆/多筆/油卡)", "無使用"], horizontal=True)
                    
                    if report_mode == "用油量申報 (含單筆/多筆/油卡)":
                        c_btn1, c_btn2, _ = st.columns([1, 1, 3])
                        with c_btn1: 
                            if st.button("➕ 增加一列"): st.session_state['multi_row_count'] += 1
                        with c_btn2: 
                            if st.button("➖ 減少一列") and st.session_state['multi_row_count'] > 1: st.session_state['multi_row_count'] -= 1

                    with st.form("entry_form", clear_on_submit=True):
                        col_p1, col_p2 = st.columns(2)
                        p_name = col_p1.text_input("👤 填報人姓名 (必填)")
                        p_ext = col_p2.text_input("📞 聯絡分機 (必填)")
                        fuel_card_id = ""; data_entries = []; f_files = None; note_input = ""
                        
                        if report_mode == "用油量申報 (含單筆/多筆/油卡)":
                            fuel_card_id = st.text_input("💳 油卡編號 (選填)")
                            for i in range(st.session_state['multi_row_count']):
                                c_d, c_v = st.columns(2)
                                _date = c_d.date_input(f"📅 日期 {i+1}", datetime.today(), key=f"d_{i}")
                                _vol = c_v.number_input(f"💧 油量 {i+1}", min_value=0.0, step=0.1, key=f"v_{i}")
                                data_entries.append({"date": _date, "vol": _vol})
                            is_shared = st.checkbox("與其他設備共用加油單")
                            note_input = st.text_input("備註內容")
                            f_files = st.file_uploader("上傳佐證", accept_multiple_files=True)
                        else:
                            st.info("ℹ️ 您選擇了「無使用」，請選擇無使用的期間。")
                            c_s, c_e = st.columns(2)
                            d_start = c_s.date_input("開始日期", datetime(datetime.now().year, 1, 1))
                            d_end = c_e.date_input("結束日期", datetime.now())
                            data_entries.append({"date": d_end, "vol": 0.0})
                            note_input = f"無使用 (期間: {d_start} ~ {d_end})"
                            is_shared = False

                        st.markdown("---"); st.markdown(privacy_html, unsafe_allow_html=True)
                        agree = st.checkbox("同意個資聲明")
                        submitted = st.form_submit_button("🚀 確認送出", use_container_width=True)
                        
                        if submitted:
                            if not agree: st.error("❌ 請勾選同意聲明")
                            elif not p_name or not p_ext: st.warning("⚠️ 姓名與分機為必填")
                            elif report_mode == "用油量申報 (含單筆/多筆/油卡)" and not f_files: st.error("⚠️ 請上傳佐證")
                            else:
                                if data_entries[0]['vol'] <= 0 and report_mode == "用油量申報 (含單筆/多筆/油卡)": st.warning("⚠️ 油量需大於 0")
                                else:
                                    valid=True; links=[]
                                    if f_files:
                                        for idx, f in enumerate(f_files):
                                            try:
                                                f.seek(0); clean_name = f"{selected_dept}_{selected_device}_{data_entries[0]['date']}_{idx+1}.{f.name.split('.')[-1]}"
                                                meta = {'name': clean_name, 'parents': [DRIVE_FOLDER_ID]}
                                                media = MediaIoBaseUpload(f, mimetype=f.type, resumable=True)
                                                file = drive_service.files().create(body=meta, media_body=media, fields='webViewLink').execute()
                                                links.append(file.get('webViewLink'))
                                            except: valid=False; st.error("上傳失敗"); break
                                    
                                    if valid:
                                        rows = []; now_str = get_taiwan_time().strftime("%Y-%m-%d %H:%M:%S")
                                        final_link = "\n".join(links) if links else "無"
                                        for e in data_entries:
                                            if e['vol']>0 or report_mode=="無使用":
                                                rows.append([now_str, selected_dept, p_name, p_ext, selected_device, str(row.get('校內財產編號','-')), str(row.get('原燃物料名稱','-')), fuel_card_id, str(e['date']), e['vol'], "是" if is_shared else "-", note_input, final_link])
                                        if rows: ws_record.append_rows(rows); st.success("✅ 申報成功！"); time.sleep(1); st.rerun()

    with tabs[1]: st.info("📊 外部看板功能維持不變")

# ------------------------------------------
# 👑 超級管理員專區 (V106.0: 校正回歸)
# ------------------------------------------
elif st.session_state['current_page'] == 'admin_dashboard' and username == 'admin':
    st.title("👑 超級管理員後台")
    
    # 1. 核心資料預處理 (Core Data Pipeline)
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

    admin_tabs = st.tabs(["📝 全校燃油設備總覽", "🔍 申報資料異動", "📊 動態管理儀表板"])

    # === Tab A: 全校總覽 (雙模式) ===
    with admin_tabs[0]:
        view_mode = st.radio("檢視模式", ["📋 設備明細檢視", "📊 設備類型統計"], horizontal=True, label_visibility="collapsed")
        st.markdown("---")

        if view_mode == "📊 設備類型統計":
            if not df_year.empty and not df_equip.empty:
                # 1. KPI
                total_eq = len(df_equip)
                gas_eq = len(df_equip[df_equip['原燃物料名稱'].str.contains('汽油', na=False)])
                diesel_eq = len(df_equip[df_equip['原燃物料名稱'].str.contains('柴油', na=False)])
                
                k1, k2, k3 = st.columns(3)
                k1.metric("🔥 全校燃油設備總數", total_eq)
                k2.metric("⛽ 全校汽油設備數", gas_eq)
                k3.metric("🚛 全校柴油設備數", diesel_eq)
                st.markdown("---")

                # 2. 統計卡片
                st.subheader("📂 各類設備用油統計")
                cols = st.columns(3)
                eq_counts = df_equip.groupby('統計類別').size()
                fuel_sums = df_year.groupby(['統計類別', '油品大類'])['加油量'].sum().unstack(fill_value=0)
                
                for idx, category in enumerate(DEVICE_ORDER):
                    with cols[idx % 3]:
                        count = eq_counts.get(category, 0)
                        gas_vol = fuel_sums.loc[category, '汽油'] if category in fuel_sums.index and '汽油' in fuel_sums.columns else 0
                        diesel_vol = fuel_sums.loc[category, '柴油'] if category in fuel_sums.index and '柴油' in fuel_sums.columns else 0
                        total_vol = gas_vol + diesel_vol
                        header_color = MORANDI_COLORS.get(category, "#CFD8DC")
                        
                        st.markdown(f"""
                        <div class="stat-card-v106">
                            <div class="stat-header" style="background-color: {header_color};">
                                <span class="stat-title">{category}</span>
                                <span class="stat-count">{count}</span>
                            </div>
                            <div class="stat-body">
                                <div class="stat-row"><span>⛽ 汽油用量</span><span>{gas_vol:,.1f} L</span></div>
                                <div class="stat-row"><span>🚛 柴油用量</span><span>{diesel_vol:,.1f} L</span></div>
                                <div class="stat-row"><span>💧 總計用量</span><span>{total_vol:,.1f} L</span></div>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
                
                st.markdown("---")
                # 3. 環形圖
                st.subheader("🍩 油品設備用油量佔比分析")
                c_pie1, c_pie2 = st.columns(2)
                
                gas_data = df_year[df_year['油品大類'] == '汽油'].groupby('統計類別')['加油量'].sum().reset_index()
                if not gas_data.empty:
                    fig_g = px.pie(gas_data, values='加油量', names='統計類別', title='⛽ 汽油用量佔比', hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
                    c_pie1.plotly_chart(fig_g, use_container_width=True)
                else: c_pie1.info("無汽油數據")
                
                dsl_data = df_year[df_year['油品大類'] == '柴油'].groupby('統計類別')['加油量'].sum().reset_index()
                if not dsl_data.empty:
                    fig_d = px.pie(dsl_data, values='加油量', names='統計類別', title='🚛 柴油用量佔比', hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
                    c_pie2.plotly_chart(fig_d, use_container_width=True)
                else: c_pie2.info("無柴油數據")
            else: st.warning("尚無資料可供統計。")

        else: # 明細檢視
            with st.expander("🔍 篩選未申報名單 (點擊展開)", expanded=False):
                c_f1, c_f2 = st.columns(2)
                d_start = c_f1.date_input("查詢起始日", date(selected_admin_year, 1, 1))
                d_end = c_f2.date_input("查詢結束日", date.today())
                
                if st.button("開始篩選未申報單位"):
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

            if not df_year.empty and not df_equip.empty:
                annual = df_year.groupby('設備名稱備註')['加油量'].sum().reset_index().rename(columns={'加油量':'年度用油量'})
                last_dt = df_year.groupby('設備名稱備註')['日期格式'].max().reset_index()
                cnt = df_year.groupby('設備名稱備註').size().reset_index(name='申報次數')
                
                t_cols = ['設備編號', '設備名稱備註', '原燃物料名稱', '設備數量', '設備所屬單位/部門', '保管人', '設備詳細位置/樓層', '統計類別']
                ex_cols = [c for c in t_cols if c in df_equip.columns]
                res = pd.merge(df_equip[ex_cols], annual, on='設備名稱備註', how='left')
                res = pd.merge(res, last_dt, on='設備名稱備註', how='left')
                res = pd.merge(res, cnt, on='設備名稱備註', how='left')
                res['年度用油量'] = res['年度用油量'].fillna(0); res['申報次數'] = res['申報次數'].fillna(0).astype(int)
                
                for category in DEVICE_ORDER:
                    sub = res[res['統計類別'] == category]
                    if not sub.empty:
                        h_col = MORANDI_COLORS.get(category, "#CFD8DC")
                        st.markdown(f"### 📂 {category}")
                        cols = st.columns(2)
                        for i, row in sub.reset_index().iterrows():
                            with cols[i%2]:
                                last = row['日期格式']
                                if pd.isna(last): stat_html = '<span class="status-badge status-warn">⚠️ 尚無紀錄</span>'; d_str = "無"
                                else:
                                    diff = (datetime.now() - last).days
                                    stat_html = f'<span class="status-badge status-warn">⚠️ 逾期未填</span>' if diff > 180 else ''
                                    d_str = last.strftime("%Y-%m-%d")
                                ft = "⛽" if "汽油" in str(row['原燃物料名稱']) else "🚛"
                                st.markdown(f"""<div class="equip-card"><div class="equip-header" style="background-color: {h_col};"><div class="equip-title-group"><div class="equip-code">{row.get('設備編號','-')}</div><div class="equip-name">{row.get('設備名稱備註','-')}</div></div><div class="equip-fuel-group"><div class="equip-vol">{row['年度用油量']:,.2f}</div><span class="equip-fuel-type">{ft} {row.get('原燃物料名稱','')} (公升)</span></div></div><div class="equip-body"><div class="equip-info">🏢 部門: {row.get('設備所屬單位/部門','-')} | 👤 保管人: {row.get('保管人','-')}<br>📍 位置: {row.get('設備詳細位置/樓層','-')} | 📊 數量: {row.get('設備數量','-')}</div><div class="equip-footer"><div class="last-date">最後申報日期: {d_str}</div><div style="display:flex; align-items:center;"><span class="count-text">申報次數: <b>{row['申報次數']}</b></span>{stat_html}</div></div></div></div>""", unsafe_allow_html=True)
            else: st.warning("尚無資料可供統計。")

    # === Tab B: 申報資料異動 ===
    with admin_tabs[1]:
        st.subheader("🔍 申報資料異動")
        if not df_year.empty:
            df_year['加油日期'] = pd.to_datetime(df_year['加油日期']).dt.date
            edited = st.data_editor(df_year, column_config={"佐證資料": st.column_config.LinkColumn("佐證", display_text="🔗"), "加油日期": st.column_config.DateColumn("日期", format="YYYY-MM-DD"), "加油量": st.column_config.NumberColumn("油量", format="%.2f"), "填報時間": st.column_config.TextColumn("填報時間", disabled=True)}, num_rows="dynamic", use_container_width=True, key="editor_v106")
            if st.button("💾 儲存變更", type="primary"):
                try:
                    ws_record.clear()
                    exp = edited.copy(); exp['加油日期'] = exp['加油日期'].astype(str)
                    ws_record.update([exp.columns.tolist()] + exp.astype(str).values.tolist())
                    st.success("✅ 更新成功！"); st.cache_data.clear(); time.sleep(1); st.rerun()
                except Exception as e: st.error(f"更新失敗: {e}")
        else: st.info(f"{selected_admin_year} 年度尚無資料。")

    # === Tab C: 儀表板 (留白) ===
    with admin_tabs[2]:
        st.info("🚧 動態管理儀表板 - 架構重設中...")

    st.markdown('<div class="contact-footer">管理員系統版本 V106.0 (Alignment & Fixes)</div>', unsafe_allow_html=True)