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
# 0. 系統設定 (V165.2)
# ==========================================
st.set_page_config(page_title="燃油使用填報", page_icon="⛽", layout="wide")

def get_taiwan_time():
    return datetime.utcnow() + timedelta(hours=8)

# ==========================================
# 1. CSS 樣式表 (V165.2 優化版)
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
        --axis-gray: #424949; /* 深灰色座標軸 */
        --btn-light-blue: #D6EAF8; /* 淺藍色按鈕底 */
    }

    [data-testid="stAppViewContainer"] { background-color: #EAEDED; color: var(--text-main); }
    [data-testid="stHeader"] { background-color: rgba(0,0,0,0); }
    [data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 1px solid #BDC3C7; }

    /* 輸入元件優化 */
    div[data-baseweb="input"] > div, div[data-baseweb="base-input"] > input, textarea, input {
        background-color: #FFFFFF !important; border-color: #BDC3C7 !important; color: #000000 !important; font-size: 1.2rem !important;
    }
    div[data-baseweb="select"] > div { border-color: #BDC3C7 !important; background-color: #FFFFFF !important; }
    
    /* 申報類型按鈕樣式 (Radio 轉 Button) */
    div[role="radiogroup"] {
        display: flex;
        gap: 10px;
        background-color: transparent !important;
        border: none !important;
    }
    div[role="radiogroup"] label {
        background-color: var(--btn-light-blue) !important;
        color: #000000 !important;
        border: 1px solid #AED6F1 !important;
        border-radius: 8px !important;
        padding: 10px 20px !important;
        font-size: 1.3rem !important;
        font-weight: 800 !important;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        cursor: pointer;
        transition: all 0.2s;
    }
    div[role="radiogroup"] label:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 6px rgba(0,0,0,0.15);
    }
    div[role="radiogroup"] label[data-checked="true"] {
        background-color: #5DADE2 !important; /* 選中時變深藍 */
        color: #FFFFFF !important;
        border: 1px solid #2E86C1 !important;
    }

    /* 送出按鈕 */
    div.stButton > button[kind="primary"], [data-testid="stFormSubmitButton"] > button {
        background-color: var(--orange-bg) !important; 
        color: #FFFFFF !important; border: 2px solid var(--orange-dark) !important; border-radius: 12px !important;
        font-size: 1.4rem !important; font-weight: 900 !important; padding: 0.8rem 2rem !important;
        box-shadow: 0 4px 6px rgba(230, 126, 34, 0.3) !important; width: 100%; 
    }
    div.stButton > button[kind="secondary"] {
        background-color: #FFFFFF !important; color: var(--text-main) !important; border: 2px solid #BDC3C7 !important;
    }

    /* --- 設備詳細卡片 (V165.1 字體放大版) --- */
    .dev-card-v165 {
        background-color: #FFFFFF; border: 1px solid #BDC3C7; border-radius: 12px; overflow: hidden;
        box-shadow: 0 3px 6px rgba(0,0,0,0.08); margin-bottom: 20px; display: flex; flex-direction: column;
    }
    .horizontal-card-v165 {
        display: flex; border: 2px solid #BDC3C7; border-radius: 15px; overflow: hidden; 
        margin-bottom: 25px; box-shadow: 0 5px 10px rgba(0,0,0,0.1); background-color: #FFFFFF; min-height: 260px;
    }
    .card-left { 
        flex: 3; background-color: var(--morandi-blue); color: #FFFFFF; 
        display: flex; flex-direction: column; justify-content: center; align-items: center; 
        padding: 20px; text-align: center; border-right: 1px solid #2C3E50; 
    }
    .dept-text { font-size: 1.8rem; font-weight: 800; margin-bottom: 10px; line-height: 1.4; } /* 放大 */
    .cat-text { font-size: 1.4rem; font-weight: 900; margin-top: 15px; opacity: 0.95; } /* 放大 */
    
    .card-right { flex: 7; padding: 25px 35px; display: flex; flex-direction: column; justify-content: center; }
    .dev-title {
        font-size: 1.8rem; font-weight: 900; color: #2C3E50; margin-bottom: 15px; 
        border-bottom: 3px solid #E67E22; padding-bottom: 8px;
    }
    .info-row { display: flex; align-items: center; padding: 8px 0; font-size: 1.35rem; color: #566573; border-bottom: 1px dashed #F2F3F4; } /* 放大 */
    .info-label { font-weight: 800; margin-right: 15px; min-width: 160px; color: #2E4053; }
    .info-value { font-weight: 600; color: #17202A; flex: 1; }

    /* 表單區塊設計 */
    .form-section {
        background-color: #F8F9F9; border: 1px solid #D5DBDB; border-radius: 12px;
        padding: 20px; margin-bottom: 15px;
    }
    .form-header {
        font-size: 1.3rem; font-weight: bold; color: #2C3E50; margin-bottom: 15px;
        border-left: 5px solid #E67E22; padding-left: 10px;
    }

    /* 看板相關 */
    .dashboard-main-title { font-size: 2rem; font-weight: 900; text-align: center; color: #2C3E50; margin-bottom: 25px; background-color: #F8F9F9; padding: 15px; border-radius: 12px; border: 1px solid #BDC3C7; }
    .kpi-card { padding: 20px; border-radius: 15px; text-align: center; background-color: #FFFFFF; box-shadow: 0 4px 10px rgba(0,0,0,0.1); border: 1px solid #BDC3C7; height: 100%; }
    .kpi-title { font-size: 1.3rem; font-weight: bold; color: var(--text-sub); margin-bottom: 5px; }
    .kpi-value { font-size: 3rem; font-weight: 800; color: var(--text-main); }
    .kpi-unit { font-size: 1.1rem; color: var(--text-sub); }

    /* 設備統計卡片 */
    .dev-stat-header { padding: 12px 15px; display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid rgba(0,0,0,0.1); color: #2C3E50; font-weight: 800; }
    .dev-stat-body { padding: 15px; font-size: 1.1rem; display: grid; gap: 8px; }
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
# 3. 資料庫連線
# ==========================================
SHEET_ID = "1gqDU21YJeBoBOd8rMYzwwZ45offXWPGEODKTF6B8k-Y" 
DRIVE_FOLDER_ID = "1Uryuk3-9FHJ39w5Uo8FYxuh9VOFndeqD"
VIP_UNITS = ["總務處事務組", "民雄總務", "新民聯辦", "產推處產學營運組"]
DEVICE_CODE_MAP = {"GV-1": "公務車輛(GV-1-)", "GV-2": "乘坐式割草機(GV-2-)", "GV-3": "乘坐式農用機具(GV-3-)", "GS-1": "鍋爐(GS-1-)", "GS-2": "發電機(GS-2-)", "GS-3": "肩背或手持式割草機、吹葉機(GS-3-)", "GS-4": "肩背或手持式農用機具(GS-4-)"}
MORANDI_COLORS = { "公務車輛(GV-1-)": "#B0C4DE", "乘坐式割草機(GV-2-)": "#F5CBA7", "乘坐式農用機具(GV-3-)": "#D7BDE2", "鍋爐(GS-1-)": "#E6B0AA", "發電機(GS-2-)": "#A9CCE3", "肩背或手持式割草機、吹葉機(GS-3-)": "#A3E4D7", "肩背或手持式農用機具(GS-4-)": "#F9E79F" }
DASH_PALETTE = ['#B0C4DE', '#F5CBA7', '#A9CCE3', '#E6B0AA', '#D7BDE2', '#A3E4D7', '#F9E79F', '#95A5A6', '#85C1E9', '#D2B4DE', '#F1948A', '#76D7C4']

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
except Exception as e: st.error(f"連線失敗: {e}"); st.stop()

@st.cache_data(ttl=600)
def load_fuel_data():
    max_retries = 3; delay = 2; df_e = pd.DataFrame(); df_r = pd.DataFrame()
    for attempt in range(max_retries):
        try: df_e = pd.DataFrame(ws_equip.get_all_records()).astype(str); break
        except: time.sleep(delay)
    
    if '設備編號' in df_e.columns: df_e['統計類別'] = df_e['設備編號'].apply(lambda c: next((v for k, v in DEVICE_CODE_MAP.items() if str(c).startswith(k)), "其他/未分類"))
    else: df_e['統計類別'] = "未設定I欄"
    
    for attempt in range(max_retries):
        try: data = ws_record.get_all_values(); df_r = pd.DataFrame(data[1:], columns=data[0]) if len(data) > 1 else pd.DataFrame(columns=data[0]); break
        except: time.sleep(delay)
    return df_e, df_r

df_equip, df_records = load_fuel_data()

if 'multi_row_count' not in st.session_state: st.session_state['multi_row_count'] = 1

# ==========================================
# 4. 介面邏輯 (Front-End)
# ==========================================
def render_user_interface():
    st.markdown("### ⛽ 燃油設備填報專區")
    tabs = st.tabs(["📝 新增填報", "📊 動態查詢看板"])
    
    # --- Tab 1: 填報 (V165.1 改版) ---
    with tabs[0]:
        st.markdown('<div class="alert-box">📢 請「誠實申報」，以保障單位及自身權益！</div>', unsafe_allow_html=True)
        
        # 1. 選擇設備
        c1, c2 = st.columns(2)
        units = sorted([x for x in df_equip['填報單位'].unique() if x != '-' and x != '填報單位'])
        selected_dept = c1.selectbox("填報單位", units, index=None, placeholder="請選擇單位...", key="dept_selector")
        
        if selected_dept:
            eq_list = df_equip[df_equip['填報單位'] == selected_dept]['設備名稱備註'].unique()
            target_eq = c2.selectbox("設備名稱", eq_list, index=None, placeholder="請選擇設備...")
            
            if target_eq:
                eq_info = df_equip[(df_equip['填報單位'] == selected_dept) & (df_equip['設備名稱備註'] == target_eq)].iloc[0]
                
                # (1) 確認設備資訊 (字體加大版)
                st.markdown("#### 步驟 2：確認設備資訊")
                icon_map = {"公務車輛": "🚙", "割草機": "🌱", "農用": "🚜", "鍋爐": "🔥", "發電機": "⚡"}
                cat_icon = next((v for k, v in icon_map.items() if k in eq_info['統計類別']), "🔧")
                
                st.markdown(f"""
                <div class="horizontal-card-v165">
                    <div class="card-left">
                        <div class="dept-text">{eq_info['設備所屬單位/部門']}</div>
                        <div style="font-size:4rem;">{cat_icon}</div>
                        <div class="cat-text">{eq_info['統計類別'].split('(')[0]}</div>
                    </div>
                    <div class="card-right">
                        <div class="dev-title">{eq_info['設備名稱備註']}</div>
                        <div class="info-row"><span class="info-label">🔢 校內財產編號</span><span class="info-value">{eq_info['校內財產編號']}</span></div>
                        <div class="info-row"><span class="info-label">👤 保管人</span><span class="info-value">{eq_info['保管人']}</span></div>
                        <div class="info-row"><span class="info-label">⛽ 原燃物料名稱</span><span class="info-value" style="color:#C0392B;">{eq_info['原燃物料名稱']}</span></div>
                        <div class="info-row"><span class="info-label">📦 設備數量</span><span class="info-value">{eq_info['設備數量']}</span></div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

                # (2) 填寫加油紀錄 (改版：區分申報類型)
                st.markdown("#### 步驟 3：填寫加油紀錄")
                
                with st.form("main_form", clear_on_submit=True):
                    # A. 申報類型 (淺藍色按鈕樣式)
                    report_type = st.radio(
                        "請選擇申報類型：",
                        ["用油量申報 (含單筆/多筆/油卡)", "本月無使用 (零申報)"],
                        horizontal=True
                    )
                    st.write("") # Spacer

                    # 基本資料
                    c_basic1, c_basic2 = st.columns(2)
                    p_name = c_basic1.text_input("👤 填報人姓名 (必填)")
                    p_ext = c_basic2.text_input("📞 聯絡分機 (必填)")
                    
                    st.markdown("---")

                    # B. 根據類型顯示不同表單
                    if "無使用" in report_type:
                        st.info("💡 您選擇了「無使用」，系統將自動記錄加油量為 0。")
                        no_use_date = st.date_input("選擇申報月份 (日期請選該月任一天)", datetime.today())
                        st.text_input("備註", value="本月無使用，故無加油紀錄。", disabled=True)
                        
                    else:
                        # 用油量申報 (支援多筆)
                        st.markdown("**⛽ 請填寫加油明細 (可點擊下方按鈕增加筆數)**")
                        for i in range(st.session_state['multi_row_count']):
                            st.markdown(f"""<div class="form-section"><div class="form-header">🧾 加油單據 #{i+1}</div>""", unsafe_allow_html=True)
                            c_r1, c_r2, c_r3, c_r4 = st.columns([2, 2, 2, 3])
                            with c_r1: d_val = st.date_input(f"加油日期", datetime.today(), key=f"d_{i}")
                            with c_r2: c_val = st.text_input(f"油卡編號 (選填)", key=f"c_{i}")
                            with c_r3: v_val = st.number_input(f"加油量 (公升)", min_value=0.01, step=0.1, key=f"v_{i}")
                            with c_r4: f_val = st.file_uploader(f"上傳憑證", type=['pdf', 'jpg', 'png'], key=f"f_{i}")
                            st.checkbox("此單據與其他設備共用?", key=f"s_{i}")
                            st.markdown("</div>", unsafe_allow_html=True)

                        c_add, c_dummy = st.columns([1, 4])
                        if c_add.form_submit_button("➕ 增加一筆單據", type="secondary"): 
                            st.session_state['multi_row_count'] += 1; st.rerun()
                        
                        st.text_area("備註 (若有共用油單或其他特殊狀況請說明)", key="s_note")

                    # 隱私聲明
                    st.markdown('<div class="correction-note">如有資料誤繕情形，請重新登錄1次資訊，並於備註欄填寫：「前筆資料誤繕，請刪除。」</div>', unsafe_allow_html=True)
                    st.markdown("""<div class="privacy-box"><div class="privacy-title">📜 個資聲明</div>本表單蒐集之姓名僅供公務聯絡及稽核使用，保存期限依相關規定辦理。</div>""", unsafe_allow_html=True)
                    
                    # 送出按鈕
                    submitted = st.form_submit_button("✅ 確認送出申報")

                    if submitted:
                        if not p_name or not p_ext:
                            st.error("❌ 請填寫姓名與分機")
                        else:
                            valid_rows = []
                            t_now = get_taiwan_time().strftime("%Y-%m-%d %H:%M:%S")
                            
                            with st.spinner("資料上傳中..."):
                                # 處理無使用
                                if "無使用" in report_type:
                                    valid_rows.append([t_now, selected_dept, p_name, p_ext, target_eq, eq_info['校內財產編號'], eq_info['原燃物料名稱'], "無", str(no_use_date), 0, "否", "本月無使用", ""])
                                
                                # 處理用油申報
                                else:
                                    for i in range(st.session_state['multi_row_count']):
                                        vol = st.session_state.get(f"v_{i}")
                                        file_val = st.session_state.get(f"f_{i}")
                                        date_val = st.session_state.get(f"d_{i}")
                                        card_val = st.session_state.get(f"c_{i}", "")
                                        is_shared = "是" if st.session_state.get(f"s_{i}") else "否"
                                        
                                        if vol and vol > 0:
                                            file_link = ""
                                            if file_val:
                                                try:
                                                    f_meta = {'name': f"{selected_dept}_{target_eq}_{date_val}_{file_val.name}", 'parents': [DRIVE_FOLDER_ID]}; media = MediaIoBaseUpload(file_val, mimetype=file_val.type)
                                                    u_f = drive_service.files().create(body=f_meta, media_body=media, fields='id, webViewLink').execute()
                                                    file_link = u_f.get('webViewLink')
                                                except: pass
                                            
                                            valid_rows.append([t_now, selected_dept, p_name, p_ext, target_eq, eq_info['校內財產編號'], eq_info['原燃物料名稱'], card_val, str(date_val), vol, is_shared, st.session_state.get("s_note", ""), file_link])

                            if valid_rows:
                                ws_record.append_rows(valid_rows)
                                st.success(f"🎉 申報成功！共新增 {len(valid_rows)} 筆紀錄。")
                                st.session_state['multi_row_count'] = 1 # Reset
                                time.sleep(2); st.rerun()
                            else:
                                st.error("❌ 無有效加油資料，請檢查加油量。")

    # === Tab 2: 看板 (V165.1 改版) ===
    with tabs[1]:
        st.markdown("### 📊 動態查詢看板 (年度檢視)")
        col_r1, col_r2 = st.columns([4, 1])
        with col_r2:
            if st.button("🔄 刷新數據", use_container_width=True): st.cache_data.clear(); st.rerun()

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
                    # 計算
                    if '原燃物料名稱' in df_final.columns:
                        gas_sum = df_final[df_final['原燃物料名稱'].str.contains('汽油', na=False)]['加油量'].sum()
                        diesel_sum = df_final[df_final['原燃物料名稱'].str.contains('柴油', na=False)]['加油量'].sum()
                        # 碳排係數：汽油 2.263, 柴油 2.606 (kgCO2e/L) -> 換算公噸 /1000
                        total_co2 = (gas_sum * 2.263 + diesel_sum * 2.606) / 1000
                    else: gas_sum = 0; diesel_sum = 0; total_co2 = 0
                    total_sum = df_final['加油量'].sum()
                    
                    st.markdown(f"<div class='dashboard-main-title'>{query_dept} - {query_year}年度 能源使用與碳排統計</div>", unsafe_allow_html=True)
                    r1c1, r1c2 = st.columns(2)
                    with r1c1: st.markdown(f"""<div class="kpi-card kpi-gas"><div class="kpi-title">⛽ 汽油使用量</div><div class="kpi-value">{gas_sum:,.1f}<span class="kpi-unit"> L</span></div></div>""", unsafe_allow_html=True)
                    with r1c2: st.markdown(f"""<div class="kpi-card kpi-diesel"><div class="kpi-title">🚛 柴油使用量</div><div class="kpi-value">{diesel_sum:,.1f}<span class="kpi-unit"> L</span></div></div>""", unsafe_allow_html=True)
                    st.markdown("<div style='margin-bottom: 20px;'></div>", unsafe_allow_html=True)
                    r2c1, r2c2 = st.columns(2)
                    with r2c1: st.markdown(f"""<div class="kpi-card kpi-total"><div class="kpi-title">💧 總用油量</div><div class="kpi-value">{total_sum:,.1f}<span class="kpi-unit"> L</span></div></div>""", unsafe_allow_html=True)
                    with r2c2: st.markdown(f"""<div class="kpi-card kpi-co2"><div class="kpi-title">☁️ 碳排放量</div><div class="kpi-value">{total_co2:,.4f}<span class="kpi-unit"> 公噸CO<sub>2</sub>e</span></div></div>""", unsafe_allow_html=True)
                    
                    st.markdown("---")

                    # (1) 年度 逐月油料統計 (字體放大 + 深灰座標)
                    st.subheader(f"📊 {query_year}年度 逐月油料統計", anchor=False)
                    df_final['月份'] = df_final['日期格式'].dt.month
                    df_final['油品類別'] = df_final['原燃物料名稱'].apply(lambda x: '汽油' if '汽油' in x else ('柴油' if '柴油' in x else '其他'))
                    
                    # 準備繪圖資料
                    monthly_data = df_final.groupby(['月份', '油品類別'])['加油量'].sum().reset_index()
                    fig = px.bar(monthly_data, x='月份', y='加油量', color='油品類別', 
                                 text_auto='.1f', # 自動顯示數值
                                 color_discrete_map={'汽油': '#52BE80', '柴油': '#F4D03F', '其他': '#BDC3C7'})
                    
                    # 優化字體設定
                    fig.update_layout(
                        xaxis=dict(tickmode='linear', dtick=1, title_font=dict(size=18, color='#424949'), tickfont=dict(size=16, color='#424949')),
                        yaxis=dict(title="加油量 (L)", title_font=dict(size=18, color='#424949'), tickfont=dict(size=16, color='#424949')),
                        legend=dict(font=dict(size=14)),
                        font=dict(family="Arial", size=14, color="#2C3E50"),
                        height=550,
                        margin=dict(t=50, b=50)
                    )
                    fig.update_traces(textfont=dict(size=16), textposition='outside') # 資料標籤放大
                    st.plotly_chart(fig, use_container_width=True)

                    # (2) 設備申報資訊統計區 (莫蘭迪標題)
                    st.markdown("---")
                    st.subheader(f"📋 {query_dept} - 設備申報資訊統計區", anchor=False)
                    target_devices = df_equip[df_equip['填報單位'] == query_dept]
                    
                    if not target_devices.empty:
                        # 準備列表
                        d_list = []
                        for _, row in target_devices.iterrows():
                            d_name = row['設備名稱備註']
                            d_cat = row.get('統計類別', '其他')
                            header_bg = MORANDI_COLORS.get(d_cat, '#D5DBDB') # 莫蘭迪底色
                            d_vol = df_final[df_final['設備名稱備註'] == d_name]['加油量'].sum()
                            d_count = len(df_final[df_final['設備名稱備註'] == d_name])
                            status = '<span class="alert-status">⚠️ 0 申報</span>' if d_count == 0 else f"{d_count} 次"
                            d_list.append({'name': d_name, 'bg': header_bg, 'vol': d_vol, 'count': d_count, 'status': status, 'cat': d_cat})

                        # 顯示卡片
                        for k in range(0, len(d_list), 2):
                            cols = st.columns(2)
                            for m in range(2):
                                if k+m < len(d_list):
                                    item = d_list[k+m]
                                    with cols[m]:
                                        st.markdown(f"""
                                        <div class="dev-card-v165">
                                            <div class="dev-stat-header" style="background-color: {item['bg']};">
                                                <span>{item['name']}</span>
                                                <small>{item['cat'].split('(')[0]}</small>
                                            </div>
                                            <div class="dev-stat-body">
                                                <div style="display:flex; justify-content:space-between;">
                                                    <span>累積加油量:</span>
                                                    <span style="font-weight:900; color:#C0392B;">{item['vol']:,.1f} L</span>
                                                </div>
                                                <div style="display:flex; justify-content:space-between;">
                                                    <span>申報次數:</span>
                                                    <span>{item['status']}</span>
                                                </div>
                                            </div>
                                        </div>
                                        """, unsafe_allow_html=True)

                    # (3) 油品設備用油量佔比 (水平長條 + 排序 + 點選顯示總計)
                    st.markdown("---")
                    st.subheader("📊 油品設備用油量佔比分析", anchor=False) # 刪除括號文字
                    
                    col_bar1, col_bar2 = st.columns(2)
                    
                    # 汽油
                    gas_df = df_final[(df_final['原燃物料名稱'].str.contains('汽油', na=False)) & (df_final['加油量'] > 0)]
                    if not gas_df.empty:
                        gas_grouped = gas_df.groupby('設備名稱備註')['加油量'].sum().reset_index().sort_values('加油量', ascending=True)
                        # 計算百分比供 Hover 使用
                        total_g = gas_grouped['加油量'].sum()
                        gas_grouped['佔比'] = (gas_grouped['加油量'] / total_g * 100).round(1)
                        
                        fig_g = px.bar(gas_grouped, x='加油量', y='設備名稱備註', orientation='h', 
                                       title='⛽ 汽油設備用量排名', text='加油量',
                                       color='加油量', color_continuous_scale='Teal')
                        fig_g.update_traces(
                            texttemplate='%{x:,.1f}', textposition='outside',
                            hovertemplate='<b>%{y}</b><br>加油量: %{x:,.1f} L<br>佔比: %{customdata[0]}%<extra></extra>',
                            customdata=gas_grouped[['佔比']]
                        )
                        fig_g.update_layout(xaxis_title="加油量 (L)", yaxis_title=None, height=400)
                        with col_bar1: st.plotly_chart(fig_g, use_container_width=True)
                    else: 
                        with col_bar1: st.info("無汽油使用紀錄")

                    # 柴油
                    diesel_df = df_final[(df_final['原燃物料名稱'].str.contains('柴油', na=False)) & (df_final['加油量'] > 0)]
                    if not diesel_df.empty:
                        dsl_grouped = diesel_df.groupby('設備名稱備註')['加油量'].sum().reset_index().sort_values('加油量', ascending=True)
                        total_d = dsl_grouped['加油量'].sum()
                        dsl_grouped['佔比'] = (dsl_grouped['加油量'] / total_d * 100).round(1)

                        fig_d = px.bar(dsl_grouped, x='加油量', y='設備名稱備註', orientation='h', 
                                       title='🚛 柴油設備用量排名', text='加油量',
                                       color='加油量', color_continuous_scale='Oranges')
                        fig_d.update_traces(
                            texttemplate='%{x:,.1f}', textposition='outside',
                            hovertemplate='<b>%{y}</b><br>加油量: %{x:,.1f} L<br>佔比: %{customdata[0]}%<extra></extra>',
                            customdata=dsl_grouped[['佔比']]
                        )
                        fig_d.update_layout(xaxis_title="加油量 (L)", yaxis_title=None, height=400)
                        with col_bar2: st.plotly_chart(fig_d, use_container_width=True)
                    else: 
                        with col_bar2: st.info("無柴油使用紀錄")

                    # (4) 碳排放結構 (矩形樹狀圖)
                    st.markdown("---")
                    st.subheader("🌍 單位油料使用碳排放量結構 (矩形樹狀圖)")
                    # 計算各設備碳排
                    df_final['CO2e_ton'] = df_final.apply(lambda r: (r['加油量']*2.263/1000) if '汽油' in str(r['原燃物料名稱']) else (r['加油量']*2.606/1000), axis=1)
                    
                    if df_final['CO2e_ton'].sum() > 0:
                        fig_tree = px.treemap(df_final, 
                                              path=['原燃物料名稱', '設備名稱備註'], 
                                              values='CO2e_ton',
                                              color='CO2e_ton', color_continuous_scale='RdBu_r',
                                              title=f"{query_year}年度 碳排放量分佈 (總計: {total_co2:.2f} 公噸CO2e)")
                        fig_tree.update_traces(
                            textinfo="label+value+percent entry",
                            texttemplate='<b>%{label}</b><br>%{value:.3f} 噸<br>(%{percentEntry:.1%})',
                            textfont=dict(size=16)
                        )
                        fig_tree.update_layout(height=600)
                        st.plotly_chart(fig_tree, use_container_width=True)
                    else: st.info("無碳排數據")

                    # (5) 年度 填報明細
                    st.markdown("---")
                    st.subheader("📝 年度 填報明細")
                    st.dataframe(
                        df_final[['加油日期', '設備名稱備註', '原燃物料名稱', '加油量', '填報人', '備註']].sort_values('加油日期', ascending=False),
                        use_container_width=True,
                        hide_index=True
                    )

                else: st.info("📭 該年度尚無相關申報紀錄。")

if __name__ == "__main__":
    render_user_interface()