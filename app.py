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

# ==========================================
# 0. 系統設定
# ==========================================
st.set_page_config(page_title="國立嘉義大學碳盤查平台", page_icon="🌍", layout="wide")

st.markdown("""
<style>
    /* =========================================
       🎨 V53.0 按鈕風格重塑版 (Orange Border Style)
       ========================================= */

    /* 1. 變數定義 */
    :root {
        /* 核心顏色 */
        --btn-border: #E67E22;    /* 粗橘色 */
        --btn-bg: #F2F4F4;        /* 淺灰底 */
        --btn-text: #2C3E50;      /* 深色文字 */
        --highlight-red: #C0392B;
        
        /* 淺色模式預設 */
        --bg-color: #EAEDED;
        --card-bg: #FFFFFF;
        --text-main: #2C3E50;
        --text-sub: #566573;
        --border-color: #BDC3C7;
        --kpi-header-bg: #D6DBDF;
    }

    /* 2. 深色模式偵測 (維持強制亮色邏輯，但適配背景) */
    @media (prefers-color-scheme: dark) {
        :root {
            --bg-color: #17202A;        
            --card-bg: #212F3D;         
            --text-main: #ECF0F1;       
            --text-sub: #B3B6B7;        
            --border-color: #566573;    
            --kpi-header-bg: #34495E;   
        }
    }

    /* 3. 全域背景與文字設定 */
    [data-testid="stAppViewContainer"] {
        background-color: var(--bg-color);
        color: var(--text-main);
    }
    
    [data-testid="stHeader"] { background-color: rgba(0,0,0,0); }
    
    [data-testid="stSidebar"] {
        background-color: var(--card-bg);
        border-right: 1px solid var(--border-color);
    }
    
    h1, h2, h3, h4, h5, h6, p, span, label {
        color: var(--text-main) !important;
    }

    /* 4. 輸入框專區 (強制白底黑字，確保手機深色模式可見) */
    div[data-baseweb="input"] > div, 
    div[data-baseweb="select"] > div, 
    div[data-baseweb="base-input"] > input,
    textarea, input {
        background-color: #FFFFFF !important; 
        border-color: #BDC3C7 !important;
        color: #000000 !important;            
        -webkit-text-fill-color: #000000 !important; 
        caret-color: #000000 !important;
        font-size: 1.15rem !important;
    }
    
    ul[data-baseweb="menu"] { background-color: #FFFFFF !important; }
    ul[data-baseweb="menu"] li { color: #000000 !important; }
    div[data-baseweb="select"] span { color: #000000 !important; -webkit-text-fill-color: #000000 !important;}

    /* 5. 按鈕專區 (🔥 V53.0 重點：粗橘框 + 淺灰底) */
    div.stButton > button {
        background-color: var(--btn-bg) !important; 
        color: var(--btn-text) !important;            
        border: 3px solid var(--btn-border) !important; /* 3px 粗框 */
        border-radius: 12px;
        font-size: 1.3rem !important;
        font-weight: 800 !important;
        padding: 0.6rem 1.5rem;
        transition: all 0.2s ease;
    }
    
    /* Hover 效果 */
    div.stButton > button:hover { 
        border-color: #D35400 !important; /* 深橘色 */
        background-color: #FFFFFF !important;
        color: #D35400 !important;
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(230, 126, 34, 0.3);
    }
    
    /* Active (點擊) 效果 */
    div.stButton > button:active {
        background-color: #E67E22 !important;
        color: #FFFFFF !important;
    }
    
    /* 登入標題區塊 */
    .login-header { 
        font-size: 2.2rem; 
        font-weight: 800; 
        color: var(--text-main) !important;
        text-align: center; 
        margin-bottom: 20px; 
        padding: 25px; 
        background-color: var(--card-bg); 
        border: 2px solid var(--border-color);
        border-radius: 15px; 
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }

    /* Tab 標籤 */
    button[data-baseweb="tab"] div p {
        font-size: 1.4rem !important;
        font-weight: 700 !important;
        color: var(--text-main) !important;
    }

    /* 6. KPI 互動卡片 */
    .kpi-header {
        font-size: 1.5rem;
        font-weight: 800;
        color: var(--text-main) !important;
        margin-bottom: 20px;
        text-align: center;
        background-color: var(--kpi-header-bg);
        padding: 12px;
        border-radius: 12px;
        letter-spacing: 1px;
    }
    
    .kpi-container {
        display: flex;
        justify-content: space-between;
        gap: 20px;
        margin-bottom: 25px;
        flex-wrap: wrap; 
    }
    
    .kpi-card {
        flex: 1;
        min-width: 220px; 
        padding: 20px;
        border-radius: 15px;
        text-align: center;
        background-color: var(--card-bg);
        box-shadow: 0 4px 10px rgba(0,0,0,0.1);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
        border: 1px solid var(--border-color);
        cursor: default;
    }
    
    .kpi-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 10px 20px rgba(0,0,0,0.15);
    }

    .kpi-card-total { border-bottom: 6px solid #5DADE2; } 
    .kpi-card-gas { border-bottom: 6px solid #58D68D; }   
    .kpi-card-diesel { border-bottom: 6px solid #F4D03F; } 
    
    .kpi-title { font-size: 1.2rem; font-weight: bold; margin-bottom: 5px; opacity: 0.8; color: var(--text-sub) !important; }
    .kpi-value { font-size: 2.8rem; font-weight: 800; line-height: 1.1; margin-bottom: 5px; color: var(--text-main) !important; }
    .kpi-unit { font-size: 1rem; font-weight: normal; color: var(--text-sub) !important; margin-left: 5px;}
    
    .kpi-sub { 
        font-size: 1rem; 
        color: var(--highlight-red) !important;
        font-weight: 700; 
        background-color: rgba(192, 57, 43, 0.1);
        padding: 4px 10px; 
        border-radius: 20px; 
        display: inline-block;
        margin-top: 5px;
    }

    .device-info-box {
        background-color: var(--card-bg);
        border: 2px solid #5DADE2;
        border-radius: 10px;
        padding: 20px;
        margin-bottom: 20px;
    }
    
    .pie-chart-box {
        background-color: var(--card-bg);
        border: 2px solid var(--border-color);
        border-radius: 15px;
        padding: 10px;
    }

    .alert-box {
        background-color: #FCF3CF;
        border: 2px solid #F1C40F;
        padding: 15px;
        border-radius: 10px;
        margin-bottom: 20px;
        color: #9A7D0A !important;
        font-weight: bold;
        text-align: center;
    }
    
    .setting-box {
        background-color: var(--card-bg);
        border: 2px dashed var(--border-color);
        padding: 15px;
        border-radius: 10px;
        margin-bottom: 20px;
        text-align: center;
    }
    
    div[role="radiogroup"] label {
        background-color: var(--card-bg);
        border: 1px solid var(--border-color);
        padding: 10px 20px;
        border-radius: 8px;
        margin-right: 10px;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

# ☁️ 設定區 (⚠️ 請務必填回學校的 ID)
SHEET_ID = "1gqDU21YJeBoBOd8rMYzwwZ45offXWPGEODKTF6B8k-Y" 
DRIVE_FOLDER_ID = "1DCmR0dXOdFBdTrgnvCYFPtNq_bGzSJeB" 

# ==========================================
# 1. 安全登入模組
# ==========================================
def clean_secrets(obj):
    if isinstance(obj, dict) or "AttrDict" in str(type(obj)):
        return {k: clean_secrets(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [clean_secrets(i) for i in obj]
    return obj

if 'current_page' not in st.session_state:
    st.session_state['current_page'] = 'home'

if 'reset_counter' not in st.session_state:
    st.session_state['reset_counter'] = 0

if 'multi_row_count' not in st.session_state:
    st.session_state['multi_row_count'] = 1

try:
    _raw_creds = st.secrets["credentials"]
    credentials_login = clean_secrets(_raw_creds)
    cookie_cfg = st.secrets["cookie"]
    
    authenticator = stauth.Authenticate(
        credentials_login,
        cookie_cfg["name"],
        cookie_cfg["key"],
        cookie_cfg["expiry_days"],
    )
    
    if st.session_state["authentication_status"] is not True:
        st.markdown('<div class="login-header">🏫 國立嘉義大學碳盤查<br>油料使用及冷媒填充回報平台</div>', unsafe_allow_html=True)
        st.markdown("<h3 style='text-align: center;'>登入系統 (Login)</h3>", unsafe_allow_html=True)
        st.markdown("---")
    
    authenticator.login('main')
    
    if st.session_state["authentication_status"] is False:
        st.error('❌ 帳號或密碼錯誤 (Username/Password Incorrect)')
        st.stop()
    elif st.session_state["authentication_status"] is None:
        st.info('🔒 請輸入帳號密碼登入 (Please Login)')
        st.stop()
        
    name = st.session_state["name"]
    username = st.session_state["username"]
    
    with st.sidebar:
        st.header(f"👤 {name}")
        st.success("☁️ 雲端連線正常")
        if st.button("🏠 返回主選單"):
            st.session_state['current_page'] = 'home'
            st.rerun()
        st.markdown("---")
        authenticator.logout('登出系統 (Logout)', 'sidebar')

except Exception as e:
    st.error(f"登入錯誤: {e}")
    st.stop()

# ==========================================
# 2. 雲端連線 (OAuth)
# ==========================================
@st.cache_resource
def init_google():
    oauth_info = st.secrets["gcp_oauth"]
    creds = Credentials(
        token=None, 
        refresh_token=oauth_info["refresh_token"],
        token_uri="https://oauth2.googleapis.com/token",
        client_id=oauth_info["client_id"],
        client_secret=oauth_info["client_secret"],
        scopes=["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/spreadsheets"]
    )
    gc = gspread.authorize(creds)
    drive_service = build('drive', 'v3', credentials=creds)
    return gc, drive_service

try:
    gc, drive_service = init_google()
    sh = gc.open_by_key(SHEET_ID)
    
    try: ws_equip = sh.worksheet("設備清單") 
    except: ws_equip = sh.sheet1 
    
    try: ws_record = sh.worksheet("填報紀錄")
    except: 
        ws_record = sh.add_worksheet(title="填報紀錄", rows="1000", cols="13")
        
    if len(ws_record.get_all_values()) == 0:
        ws_record.append_row([
            "填報時間", "填報單位", "填報人", "填報人分機", 
            "設備名稱備註", "校內財產編號", "原燃物料名稱", 
            "油卡編號", "加油日期", "加油量", "與其他設備共用加油單", "備註", "佐證資料"
        ])

except Exception as e:
    st.error(f"連線失敗: {e}")
    st.stop()

@st.cache_data(ttl=600)
def load_data():
    df_e = pd.DataFrame(ws_equip.get_all_records())
    df_e = df_e.astype(str)
    
    data = ws_record.get_all_values()
    if len(data) > 0:
        headers = data.pop(0)
        df_r = pd.DataFrame(data, columns=headers)
    else:
        df_r = pd.DataFrame()
        
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
        if st.button("前往「燃油設備填報區」", use_container_width=True, type="primary"):
            st.session_state['current_page'] = 'fuel'
            st.rerun()
    with col2:
        st.info("❄️ 冷氣/冰水主機")
        st.button("前往「冷媒類設備填報區」 (建置中)", use_container_width=True, disabled=True)
    
    st.markdown("""
        <div class="contact-footer">
        如有填報疑問，請電洽環安中心林小姐(分機 7137)，謝謝
        </div>
    """, unsafe_allow_html=True)

elif st.session_state['current_page'] == 'fuel':
    st.title("⛽ 燃油設備填報專區")
    
    tabs = st.tabs(["📝 新增填報", "📊 動態查詢看板"])

    # --- Tab 1: 填報 ---
    with tabs[0]:
        st.markdown("""
        <div class="alert-box">
            📢 宣導事項：請「誠實申報」，以保障單位及自身權益！
        </div>
        """, unsafe_allow_html=True)

        if not df_equip.empty:
            st.markdown("#### 步驟 1：選擇設備")
            c1, c2 = st.columns(2)
            units = sorted([x for x in df_equip['填報單位'].unique() if x != '-' and x != '填報單位'])
            
            selected_dept = c1.selectbox(
                "填報單位", 
                units, 
                index=None, 
                placeholder="請選擇單位...", 
                key="dept_selector"
            )
            
            if selected_dept:
                filtered = df_equip[df_equip['填報單位'] == selected_dept]
                devices = sorted([x for x in filtered['設備名稱備註'].unique()])
                
                dynamic_key = f"vehicle_selector_{st.session_state['reset_counter']}"
                
                selected_device = c2.selectbox(
                    "車輛/機具名稱", 
                    devices, 
                    index=None, 
                    placeholder="請選擇車輛...", 
                    key=dynamic_key
                )
                
                if selected_device:
                    row = filtered[filtered['設備名稱備註'] == selected_device].iloc[0]
                    
                    info_html = f"""
                    <div class="device-info-box">
                        <div style="border-bottom: 1px solid #BDC3C7; padding-bottom: 10px; margin-bottom: 10px; font-weight: bold; font-size: 1.2rem; color: #5DADE2;">
                            📋 設備詳細資料
                        </div>
                        <div><strong>🏢 部門：</strong>{row.get('設備所屬單位/部門', '-')}</div>
                        <div><strong>👤 保管人：</strong>{row.get('保管人', '-')}</div>
                        <div><strong>🔢 財產編號：</strong>{row.get('校內財產編號', '-')}</div>
                        <div><strong>📍 位置：</strong>{row.get('設備詳細位置/樓層', '-')}</div>
                        <div><strong>⛽ 燃料：</strong>{row.get('原燃物料名稱', '-')}</div>
                        <div><strong>📊 數量：</strong>{row.get('設備數量', '-')}</div>
                    </div>
                    """
                    st.markdown(info_html, unsafe_allow_html=True)
                    
                    st.markdown("#### 步驟 2：填寫資料")
                    
                    report_mode = st.radio(
                        "請選擇申報類型", 
                        ["用油量申報 (含單筆/多筆/油卡)", "本季無使用"], 
                        horizontal=True
                    )
                    
                    if report_mode == "用油量申報 (含單筆/多筆/油卡)":
                        st.markdown('<div class="setting-box">', unsafe_allow_html=True)
                        st.markdown("**🔧 設定明細筆數** (請先調整好筆數，再進行填寫)")
                        c_btn1, c_btn2, c_dummy = st.columns([1, 1, 3])
                        with c_btn1:
                            if st.button("➕ 增加一列", use_container_width=True):
                                if st.session_state['multi_row_count'] < 10:
                                    st.session_state['multi_row_count'] += 1
                        with c_btn2:
                            if st.button("➖ 減少一列", use_container_width=True):
                                if st.session_state['multi_row_count'] > 1:
                                    st.session_state['multi_row_count'] -= 1
                        st.caption(f"目前將顯示 **{st.session_state['multi_row_count']}** 列供填寫 (上限 10 列)")
                        st.markdown('</div>', unsafe_allow_html=True)

                    with st.form("entry_form", clear_on_submit=True):
                        col_p1, col_p2 = st.columns(2)
                        p_name = col_p1.text_input("👤 填報人姓名 (必填)")
                        p_ext = col_p2.text_input("📞 聯絡分機 (必填)")
                        
                        fuel_card_id = ""
                        data_entries = []
                        f_files = None
                        is_shared = False
                        note_input = ""
                        
                        if report_mode == "用油量申報 (含單筆/多筆/油卡)":
                            fuel_card_id = st.text_input("💳 油卡編號 (選填)")
                            
                            st.divider()
                            st.markdown("⛽ **加油明細區 (必填)**")
                            
                            rows = st.session_state['multi_row_count']
                            for i in range(rows):
                                c_d, c_v = st.columns(2)
                                _date = c_d.date_input(f"📅 第 {i+1} 筆 - 加油日期", datetime.today(), key=f"md_{i}")
                                _vol = c_v.number_input(f"💧 第 {i+1} 筆 - 加油量(公升)", min_value=0.0, step=0.01, format="%.2f", key=f"mv_{i}")
                                data_entries.append({"date": _date, "vol": _vol})
                            
                            st.markdown("---")
                            is_shared = st.checkbox("與其他設備共用加油單")
                            
                            st.markdown("**🧾 備註 (選填)**")
                            st.caption("A. 若一張發票加多台設備，請填寫相同發票號碼以便核對。")
                            st.caption("B. 若有資料誤繕情形，請您重新登錄，並於備註欄註記「請刪除前筆資料，以本筆資料為準」，以利管理單位協助刪除。")
                            note_input = st.text_input("備註內容")
                            
                            st.markdown("**📂 上傳佐證資料 (必填)**")
                            f_files = st.file_uploader("支援 png, jpg, jpeg, pdf (最多 5 個，單檔限 10MB)", type=['png', 'jpg', 'jpeg', 'pdf'], accept_multiple_files=True)
                        
                        else:
                            st.info("ℹ️ 您選擇了「本季無使用」，系統將自動記錄油量為 0，無需上傳佐證資料。")
                            data_entries.append({"date": datetime.today(), "vol": 0.0})
                            note_input = "本季無使用"

                        st.markdown("---")
                        st.markdown("""
                        <div class="privacy-box">
                            <div class="privacy-title">📜 個人資料蒐集、處理及利用告知聲明</div>
                            1. <strong>蒐集機關</strong>：國立嘉義大學。<br>
                            2. <strong>蒐集目的</strong>：進行本校公務車輛/機具之加油紀錄管理、校園溫室氣體（碳）盤查統計、稽核佐證資料蒐集及後續能源使用分析。<br>
                            3. <strong>個資類別</strong>：填報人姓名。<br>
                            4. <strong>利用期間</strong>：姓名保留至填報年度後第二年1月1日，期滿即進行「去識別化」刪除，其餘數據永久保存。<br>
                            5. <strong>利用對象</strong>：本校教師、行政人員及碳盤查查驗人員。<br>
                            6. <strong>您有權依個資法請求查詢、更正或刪除您的個資。如不提供，將無法完成填報。</strong><br>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        agree_privacy = st.checkbox("我已閱讀並同意上述聲明，且確認所填資料無誤。", value=False)
                        
                        # V53.0: 調整送出按鈕佈局 (強制置中 + 放大)
                        c_sub1, c_sub2, c_sub3 = st.columns([1, 2, 1])
                        with c_sub2:
                            submitted = st.form_submit_button("🚀 確認送出資料", use_container_width=True)
                        
                        if submitted:
                            if not agree_privacy:
                                st.error("❌ 請務必勾選「我已閱讀並同意上述聲明」，才能送出資料！")
                            elif not p_name or not p_ext:
                                st.warning("⚠️ 「填報人姓名」與「聯絡分機」為必填欄位！")
                            elif report_mode == "用油量申報 (含單筆/多筆/油卡)":
                                if not f_files:
                                    st.error("⚠️ 請務必上傳佐證資料！")
                                elif len(f_files) > 5:
                                    st.error("❌ 檔案數量過多！最多 5 個。")
                                else:
                                    if data_entries[0]['vol'] <= 0:
                                        st.warning("⚠️ 第一筆加油量不能為 0，請確實填寫。")
                                    else:
                                        valid_logic = True
                                        file_links = []
                                        
                                        progress_text = "資料處理中..."
                                        my_bar = st.progress(0, text=progress_text)
                                        
                                        for idx, f_file in enumerate(f_files):
                                            try:
                                                f_file.seek(0)
                                                file_ext = f_file.name.split('.')[-1]
                                                first_date = data_entries[0]['date']
                                                clean_name = f"{selected_dept}_{selected_device}_{first_date}_{idx+1}.{file_ext}".replace("/", "_")
                                                file_meta = {'name': clean_name, 'parents': [DRIVE_FOLDER_ID]}
                                                media = MediaIoBaseUpload(f_file, mimetype=f_file.type, resumable=True)
                                                file = drive_service.files().create(body=file_meta, media_body=media, fields='webViewLink').execute()
                                                file_links.append(file.get('webViewLink'))
                                            except Exception as e:
                                                st.error(f"上傳失敗: {e}")
                                                valid_logic = False
                                                break
                                        
                                        if valid_logic:
                                            final_links = "\n".join(file_links)
                                            my_bar.progress(50, text="寫入資料庫...")
                                            
                                            rows_to_append = []
                                            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                            shared_str = "是" if is_shared else "-"
                                            card_str = fuel_card_id if fuel_card_id else "-"
                                            
                                            for entry in data_entries:
                                                if entry['vol'] > 0:
                                                    row_data = [
                                                        current_time, selected_dept, p_name, p_ext,
                                                        selected_device, str(row.get('校內財產編號', '-')), str(row.get('原燃物料名稱', '-')),
                                                        card_str, str(entry["date"]), entry["vol"], 
                                                        shared_str, note_input, final_links
                                                    ]
                                                    rows_to_append.append(row_data)
                                            
                                            if rows_to_append:
                                                ws_record.append_rows(rows_to_append)
                                                my_bar.progress(100, text="完成！")
                                                time.sleep(0.5)
                                                my_bar.empty()
                                                st.success("✅ 申報成功！")
                                                st.balloons()
                                                st.session_state['reset_counter'] += 1
                                                st.cache_data.clear()
                                                st.rerun()
                                            else:
                                                st.warning("⚠️ 沒有有效的加油資料可寫入 (油量需大於 0)。")

                            else: # 本季無使用
                                my_bar = st.progress(50, text="寫入紀錄...")
                                current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                row_data = [
                                    current_time, selected_dept, p_name, p_ext,
                                    selected_device, str(row.get('校內財產編號', '-')), str(row.get('原燃物料名稱', '-')),
                                    "-", str(data_entries[0]["date"]), 0, 
                                    "-", note_input, "無"
                                ]
                                ws_record.append_row(row_data)
                                my_bar.progress(100, text="完成！")
                                time.sleep(0.5)
                                my_bar.empty()
                                st.success("✅ 已回報本季無使用")
                                st.session_state['reset_counter'] += 1
                                st.cache_data.clear()
                                st.rerun()
        
        st.markdown("""
            <div class="contact-footer">
            如有填報疑問，請電洽環安中心林小姐(分機 7137)，謝謝
            </div>
        """, unsafe_allow_html=True)

    # --- Tab 2: 動態查詢看板 (年度檢視) ---
    with tabs[1]:
        st.markdown("### 📊 動態查詢看板 (年度檢視)")
        st.info("請選擇「單位」與「年份」，檢視該年度的用油統計與詳細紀錄。")
        
        col_r1, col_r2 = st.columns([4, 1])
        with col_r2:
            if st.button("🔄 刷新數據", use_container_width=True, key="refresh_all"): 
                st.cache_data.clear()
                st.rerun()
        
        available_years = []
        if not df_records.empty and '加油量' in df_records.columns and '加油日期' in df_records.columns:
            df_records['加油量'] = pd.to_numeric(df_records['加油量'], errors='coerce').fillna(0)
            df_records['日期格式'] = pd.to_datetime(df_records['加油日期'], errors='coerce')
            
            available_years = sorted(df_records['日期格式'].dt.year.dropna().astype(int).unique(), reverse=True)
            if not available_years:
                available_years = [datetime.now().year]
            
            record_units = sorted([str(x) for x in df_records['填報單位'].unique() if str(x) != 'nan'])
            
            c_dept, c_year = st.columns([2, 1])
            query_dept = c_dept.selectbox("🏢 選擇查詢單位", record_units, index=None, placeholder="請選擇...")
            query_year = c_year.selectbox("📅 選擇統計年度", available_years, index=0) 
            
            if query_dept and query_year:
                df_dept = df_records[df_records['填報單位'] == query_dept].copy()
                df_final = df_dept[df_dept['日期格式'].dt.year == query_year]
                
                if not df_final.empty:
                    # 1. KPI 數據
                    if '原燃物料名稱' in df_final.columns:
                        df_final['原燃物料名稱'] = df_final['原燃物料名稱'].fillna('').astype(str)
                        gas_mask = df_final['原燃物料名稱'].str.contains('汽油', na=False)
                        diesel_mask = df_final['原燃物料名稱'].str.contains('柴油', na=False)
                        gasoline_sum = df_final.loc[gas_mask, '加油量'].sum()
                        diesel_sum = df_final.loc[diesel_mask, '加油量'].sum()
                    else:
                        gasoline_sum = 0
                        diesel_sum = 0
                    
                    total_sum = df_final['加油量'].sum()
                    
                    gas_pct = (gasoline_sum / total_sum * 100) if total_sum > 0 else 0
                    diesel_pct = (diesel_sum / total_sum * 100) if total_sum > 0 else 0
                    
                    st.markdown(f"<div class='kpi-header'>{query_dept} - {query_year}年度 用油統計</div>", unsafe_allow_html=True)
                    
                    kpi_html = f"""
                    <div class="kpi-container">
                        <div class="kpi-card kpi-card-gas">
                            <div class="kpi-title">⛽ 汽油使用量</div>
                            <div class="kpi-value">{gasoline_sum:,.2f}<span class="kpi-unit"> 公升</span></div>
                            <div class="kpi-sub">佔比 {gas_pct:.2f}%</div>
                        </div>
                        <div class="kpi-card kpi-card-diesel">
                            <div class="kpi-title">🚛 柴油使用量</div>
                            <div class="kpi-value">{diesel_sum:,.2f}<span class="kpi-unit"> 公升</span></div>
                            <div class="kpi-sub">佔比 {diesel_pct:.2f}%</div>
                        </div>
                        <div class="kpi-card kpi-card-total">
                            <div class="kpi-title">💧 總用油量</div>
                            <div class="kpi-value">{total_sum:,.2f}<span class="kpi-unit"> 公升</span></div>
                            <div class="kpi-sub">100%</div>
                        </div>
                    </div>
                    """
                    st.markdown(kpi_html, unsafe_allow_html=True)
                    
                    # 2. 趨勢圖
                    st.subheader(f"📊 {query_year}年度 每月加油趨勢")
                    all_months = pd.DataFrame({'月份': list(range(1, 13))})
                    df_final['月份'] = df_final['日期格式'].dt.month
                    real_data = df_final.groupby(['月份', '設備名稱備註'])['加油量'].sum().reset_index()
                    unique_devices = df_final['設備名稱備註'].unique()
                    
                    if len(unique_devices) > 0:
                        month_device_grid = pd.MultiIndex.from_product([range(1, 13), unique_devices], names=['月份', '設備名稱備註']).to_frame(index=False)
                        merged_data = pd.merge(month_device_grid, real_data, on=['月份', '設備名稱備註'], how='left').fillna(0)
                        
                        total_data = df_final.groupby(['設備名稱備註'])['加油量'].sum().reset_index()
                        total_data['月份'] = 13
                        
                        final_chart_data = pd.concat([merged_data, total_data])
                        
                        def map_month(x):
                            return "全年度累計" if x == 13 else f"{x}月"
                        
                        final_chart_data['月份標籤'] = final_chart_data['月份'].apply(map_month)
                        
                        morandi_colors = ['#88B04B', '#92A8D1', '#F7CAC9', '#B565A7', '#009B77', '#DD4124', '#D65076', '#45B8AC', '#EFC050', '#5B5EA6']
                        
                        fig = px.bar(
                            final_chart_data, 
                            x='月份標籤', 
                            y='加油量', 
                            color='設備名稱備註', 
                            labels={'加油量': '加油量 (公升)', '月份標籤': '統計月份', '設備名稱備註': '設備名稱'},
                            color_discrete_sequence=morandi_colors,
                            template="plotly_white"
                        )
                        
                        fig.update_xaxes(type='category', tickfont=dict(size=14))
                        fig.update_yaxes(title_font=dict(size=16), tickfont=dict(size=14))
                        fig.update_traces(width=0.6, texttemplate='%{y:.2f}', textposition='inside', textfont=dict(size=14))
                        fig.update_layout(barmode='stack', font=dict(size=14), legend=dict(font=dict(size=12)))
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.info("尚無足夠資料產生圖表")
                    
                    # 3. 圓餅圖
                    st.markdown("---")
                    st.subheader("🥧 油品設備佔比分析")
                    c_pie1, c_pie2 = st.columns(2)
                    
                    with c_pie1:
                        st.markdown('<div class="pie-chart-box">', unsafe_allow_html=True) 
                        gas_df = df_final[df_final['原燃物料名稱'].str.contains('汽油', na=False)]
                        if not gas_df.empty:
                            fig_gas = px.pie(gas_df, values='加油量', names='設備名稱備註', title='⛽ 汽油設備用油佔比', color_discrete_sequence=px.colors.sequential.Teal, hole=0.4)
                            fig_gas.update_traces(textinfo='percent+label', textfont_size=14)
                            st.plotly_chart(fig_gas, use_container_width=True)
                        else:
                            st.info("無汽油使用紀錄")
                        st.markdown('</div>', unsafe_allow_html=True) 

                    with c_pie2:
                        st.markdown('<div class="pie-chart-box">', unsafe_allow_html=True) 
                        diesel_df = df_final[df_final['原燃物料名稱'].str.contains('柴油', na=False)]
                        if not diesel_df.empty:
                            fig_diesel = px.pie(diesel_df, values='加油量', names='設備名稱備註', title='🚛 柴油設備用油佔比', color_discrete_sequence=px.colors.sequential.Oranges, hole=0.4)
                            fig_diesel.update_traces(textinfo='percent+label', textfont_size=14)
                            st.plotly_chart(fig_diesel, use_container_width=True)
                        else:
                            st.info("無柴油使用紀錄")
                        st.markdown('</div>', unsafe_allow_html=True) 

                    st.markdown("---")
                    
                    # 4. 明細表
                    st.subheader(f"📋 {query_year}年度 填報明細")
                    target_cols = ["加油日期", "設備名稱備註", "原燃物料名稱", "油卡編號", "加油量", "填報人", "備註", "與其他設備共用加油單"]
                    available_cols = [c for c in target_cols if c in df_final.columns]
                    
                    df_display = df_final[available_cols].sort_values(by='加油日期', ascending=False)
                    df_display = df_display.rename(columns={'加油量': '加油量(公升)'})
                    
                    st.dataframe(df_display.style.format({"加油量(公升)": "{:.2f}"}), use_container_width=True)
                    
                else:
                    st.warning(f"⚠️ {query_dept} 在 {query_year} 年度尚無填報紀錄。")
        else:
            st.warning("📭 目前資料庫尚無有效資料，請先至「新增填報」分頁填寫。")

        st.markdown("""
            <div class="contact-footer">
            如有填報疑問，請電洽環安中心林小姐(分機 7137)，謝謝
            </div>
        """, unsafe_allow_html=True)

    # --- Tab 3: 管理 (Admin Only) ---
    if username == 'admin':
        with tabs[2]:
            st.header("🛠️ 設備資料庫管理")
            edited_df = st.data_editor(df_equip, num_rows="dynamic", use_container_width=True, key="editor")
            if st.button("💾 儲存變更", type="primary"):
                try:
                    ws_equip.clear()
                    updated_data = [edited_df.columns.tolist()] + edited_df.values.tolist()
                    ws_equip.update(updated_data)
                    st.success("✅ 更新成功！")
                    st.rerun()
                except Exception as e:
                    st.error(f"儲存失敗: {e}")