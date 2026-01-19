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

def get_taiwan_time(): return datetime.utcnow() + timedelta(hours=8)

st.markdown("""
<style>
    /* =========================================
       🎨 V96.0 穩定核心修復版 (Core Data Fix)
       ========================================= */
    :root { color-scheme: light; }
    :root {
        --btn-bg: #B0BEC5; --btn-border: #2C3E50; --btn-text: #17202A;      
        --orange-bg: #E67E22; --orange-dark: #D35400; --orange-text: #FFFFFF;
        --bg-color: #EAEDED; --card-bg: #FFFFFF; --text-main: #2C3E50;
        --border-color: #BDC3C7; --morandi-red: #A93226; 
        --kpi-gas-border: #52BE80; --kpi-diesel-border: #F4D03F;
        --kpi-total-border: #5DADE2; --kpi-co2-border: #AF7AC5;
    }
    [data-testid="stAppViewContainer"] { background-color: var(--bg-color); color: var(--text-main); }
    [data-testid="stHeader"] { background-color: rgba(0,0,0,0); }
    [data-testid="stSidebar"] { background-color: var(--card-bg); border-right: 1px solid var(--border-color); }
    h1, h2, h3, h4, h5, h6, p, span, label, div { color: var(--text-main); }
    div[data-baseweb="input"] > div, div[data-baseweb="base-input"] > input, textarea, input { background-color: #FFFFFF !important; border-color: #BDC3C7 !important; color: #000000 !important; font-size: 1.15rem !important; }
    div[data-baseweb="select"] > div { border-color: #BDC3C7 !important; background-color: #FFFFFF !important; }
    div[data-baseweb="select"] span { color: #000000 !important; }
    ul[data-baseweb="menu"] { background-color: #FFFFFF !important; }
    ul[data-baseweb="menu"] li { color: #000000 !important; }
    div.stButton > button, button[kind="primary"], [data-testid="stFormSubmitButton"] > button {
        background-color: var(--orange-bg) !important; color: var(--orange-text) !important;
        border: 2px solid var(--orange-dark) !important; border-radius: 12px !important;
        font-size: 1.3rem !important; font-weight: 800 !important; padding: 0.7rem 1.5rem !important;
        transition: all 0.2s ease !important; box-shadow: 0 4px 6px rgba(230, 126, 34, 0.3) !important;
    }
    div.stButton > button:hover, button[kind="primary"]:hover, [data-testid="stFormSubmitButton"] > button:hover { background-color: var(--orange-dark) !important; border-color: #A04000 !important; transform: translateY(-2px) !important; }
    
    /* Cards */
    .batch-card, .equip-card { background-color: #FFFFFF; border: 1px solid #BDC3C7; border-radius: 10px; margin-bottom: 10px; box-shadow: 0 2px 5px rgba(0,0,0,0.05); overflow: hidden; height: 100%; }
    .batch-card { border-left: 5px solid #E67E22; display: flex; flex-direction: column; justify-content: center; padding: 10px 15px; }
    .batch-title { font-size: 1.1rem; font-weight: bold; color: #2C3E50; margin-bottom: 3px; }
    .batch-sub { font-size: 0.9rem; color: #566573; }
    
    .equip-header { padding: 10px 15px; border-bottom: 1px solid #BDC3C7; display: flex; justify-content: space-between; align-items: center; }
    .equip-code { font-size: 1.05rem; font-weight: 800; color: #2C3E50; }
    .equip-name { font-size: 0.95rem; color: #455A64; font-weight: 600;}
    .equip-vol { font-size: 1.6rem; font-weight: 900; color: var(--morandi-red); line-height: 1.2;} 
    .equip-fuel-type { font-size: 0.85rem; color: #566573; font-weight: bold; background: rgba(255,255,255,0.6); padding: 2px 6px; border-radius: 4px; margin-left: 5px;}
    .equip-body { padding: 15px; display: flex; flex-direction: column; gap: 8px; }
    .equip-info { font-size: 0.9rem; line-height: 1.5; color: #34495E; }
    .equip-footer { margin-top: 10px; padding-top: 10px; border-top: 1px dashed #D7DBDD; display: flex; justify-content: space-between; align-items: center; }
    .status-badge { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 0.8rem; font-weight: bold; }
    .status-warn { background-color: #FADBD8; color: #943126; border: 1px solid #F1948A;}
    .count-text { font-size: 0.85rem; color: #2C3E50; font-weight: 800; margin-right: 8px; }
    
    .kpi-card { padding: 20px; border-radius: 15px; text-align: center; background-color: var(--card-bg); box-shadow: 0 4px 10px rgba(0,0,0,0.1); border: 1px solid var(--border-color); height: 100%; }
    .kpi-gas { border-bottom: 6px solid var(--kpi-gas-border); } .kpi-diesel { border-bottom: 6px solid var(--kpi-diesel-border); }
    .kpi-total { border-bottom: 6px solid var(--kpi-total-border); } .kpi-co2 { border-bottom: 6px solid var(--kpi-co2-border); }
    .kpi-title { font-size: 1.1rem; font-weight: bold; opacity: 0.8; color: #566573 !important; }
    .kpi-value { font-size: 2.5rem; font-weight: 800; color: #2C3E50 !important; margin: 5px 0;}
    .device-info-box { background-color: var(--card-bg); border: 2px solid #5DADE2; border-radius: 10px; padding: 20px; margin-bottom: 20px; }
    .alert-box { background-color: #FCF3CF; border: 2px solid #F1C40F; padding: 15px; border-radius: 10px; margin-bottom: 20px; color: #9A7D0A !important; font-weight: bold; text-align: center; }
    .login-header { font-size: 2.2rem; font-weight: 800; color: var(--text-main) !important; text-align: center; margin-bottom: 20px; padding: 25px; background-color: var(--card-bg); border: 2px solid var(--border-color); border-radius: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    .privacy-box { background-color: #F8F9F9; border: 1px solid #BDC3C7; padding: 15px; border-radius: 10px; font-size: 0.9rem; color: #566573; margin-bottom: 10px; }
    .privacy-title { font-weight: bold; color: #2C3E50; margin-bottom: 5px; font-size: 1rem; }
    button[data-baseweb="tab"] div p { font-size: 1.6rem !important; font-weight: 900 !important; color: #566573; }
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

# V96: 色彩調和修正 (GV-2/GV-3 差異化)
MORANDI_COLORS = {
    "公務車輛(GV-1-)": "#B0C4DE", "乘坐式割草機(GV-2-)": "#F5CBA7", "乘坐式農用機具(GV-3-)": "#D7BDE2",
    "鍋爐(GS-1-)": "#E6B0AA", "發電機(GS-2-)": "#A9CCE3",
    "肩背或手持式割草機、吹葉機(GS-3-)": "#A3E4D7", "肩背或手持式農用機具(GS-4-)": "#F9E79F"
}

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
    
    if st.session_state["authentication_status"] is False:
        st.error('❌ 帳號或密碼錯誤'); st.stop()
    elif st.session_state["authentication_status"] is None:
        st.info('🔒 請輸入帳號密碼登入'); st.stop()
        
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
# 頁面邏輯
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
                        
                        # V96: 文字更新
                        st.markdown("⛽ **請填入各設備該月份之加油總量(公升)，若該月份無使用請填0：**")
                        batch_inputs = {}
                        for idx, row in filtered_equip.iterrows():
                            c_card, c_val = st.columns([7, 3]) 
                            with c_card:
                                st.markdown(f"""<div class="batch-card"><div class="batch-title">⛽ {row['設備名稱備註']}</div><div class="batch-sub">{row.get('原燃物料名稱')} | 財產編號: {row.get('校內財產編號','-')} | 部門: {row.get('設備所屬單位/部門','-')}</div></div>""", unsafe_allow_html=True)
                            with c_val:
                                st.write("") 
                                vol = st.number_input(f"加油量", min_value=0.0, step=0.1, key=f"b_v_{row['校內財產編號']}_{idx}", label_visibility="collapsed")
                                batch_inputs[idx] = vol
                                
                        st.markdown("---")
                        st.markdown("**📂 上傳中油加油明細 (只需一份)**")
                        f_file = st.file_uploader("支援 PDF/JPG/PNG", type=['pdf', 'jpg', 'png', 'jpeg'])
                        st.markdown("---")
                        st.markdown(privacy_html, unsafe_allow_html=True)
                        agree_privacy = st.checkbox("我已閱讀並同意個資聲明，且確認所填資料無誤。", value=False)
                        submitted = st.form_submit_button("🚀 批次確認送出", use_container_width=True)
                        
                        if submitted:
                            total_vol = sum(batch_inputs.values())
                            if not agree_privacy: st.error("❌ 請勾選同意聲明")
                            elif not p_name or not p_ext: st.warning("⚠️ 姓名與分機為必填")
                            elif not f_file: st.error("⚠️ 請上傳加油明細佐證")
                            elif total_vol <= 0: st.warning("⚠️ 總加油量為 0，請至少填寫一台設備的油量")
                            else:
                                try:
                                    f_file.seek(0)
                                    file_ext = f_file.name.split('.')[-1]
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
                                        if vol > 0:
                                            row = filtered_equip.loc[idx]
                                            rows_to_append.append([current_time, selected_dept, p_name, p_ext, row['設備名稱備註'], str(row.get('校內財產編號','-')), row['原燃物料名稱'], fleet_id, str(batch_date), vol, "是", f"批次申報-{target_sub_cat}", file_link])
                                    if rows_to_append:
                                        ws_record.append_rows(rows_to_append)
                                        st.success(f"✅ 批次申報成功！已寫入 {len(rows_to_append)} 筆紀錄。")
                                        st.balloons(); st.session_state['reset_counter'] += 1; time.sleep(1.5); st.rerun()
                                    else: st.warning("沒有可寫入的紀錄 (數值需大於 0)。")
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
                        st.markdown('<div class="setting-box">**🔧 設定明細筆數** (請先調整好筆數，再進行填寫)</div>', unsafe_allow_html=True)
                        c_btn1, c_btn2, c_dummy = st.columns([1, 1, 3])
                        with c_btn1:
                            if st.button("➕ 增加一列", use_container_width=True): st.session_state['multi_row_count'] += 1
                        with c_btn2:
                            if st.button("➖ 減少一列", use_container_width=True) and st.session_state['multi_row_count'] > 1: st.session_state['multi_row_count'] -= 1
                        st.caption(f"目前將顯示 **{st.session_state['multi_row_count']}** 列供填寫 (上限 10 列)")

                    with st.form("entry_form", clear_on_submit=True):
                        col_p1, col_p2 = st.columns(2)
                        p_name = col_p1.text_input("👤 填報人姓名 (必填)")
                        p_ext = col_p2.text_input("📞 聯絡分機 (必填)")
                        fuel_card_id = ""; data_entries = []; f_files = None; note_input = ""
                        
                        if report_mode == "用油量申報 (含單筆/多筆/油卡)":
                            fuel_card_id = st.text_input("💳 油卡編號 (選填)")
                            st.divider()
                            st.markdown("⛽ **加油明細區 (必填)**")
                            for i in range(st.session_state['multi_row_count']):
                                c_d, c_v = st.columns(2)
                                _date = c_d.date_input(f"📅 日期 {i+1}", datetime.today(), key=f"d_{i}")
                                _vol = c_v.number_input(f"💧 油量 {i+1}", min_value=0.0, step=0.1, key=f"v_{i}")
                                data_entries.append({"date": _date, "vol": _vol})
                            st.markdown("---")
                            is_shared = st.checkbox("與其他設備共用加油單")
                            st.markdown("**🧾 備註 (選填)**")
                            st.caption("若一張發票加多台設備，請填寫相同發票號碼以便核對。")
                            note_input = st.text_input("備註內容")
                            st.markdown("**📂 上傳佐證資料 (必填)**")
                            f_files = st.file_uploader("支援 png, jpg, jpeg, pdf (最多 5 個)", type=['png', 'jpg', 'jpeg', 'pdf'], accept_multiple_files=True)
                        else:
                            st.info("ℹ️ 您選擇了「無使用」，請選擇無使用的期間。")
                            c_start, c_end = st.columns(2)
                            d_start = c_start.date_input("📅 開始日期", datetime(datetime.now().year, 1, 1))
                            d_end = c_end.date_input("📅 結束日期", datetime.now())
                            data_entries.append({"date": d_end, "vol": 0.0})
                            note_input = f"無使用 (期間: {d_start} ~ {d_end})"
                            f_files = None

                        st.markdown("---"); st.markdown(privacy_html, unsafe_allow_html=True)
                        agree_privacy = st.checkbox("我已閱讀並同意上述聲明，且確認所填資料無誤。", value=False)
                        submitted = st.form_submit_button("🚀 確認送出", use_container_width=True)
                        
                        if submitted:
                            if not agree_privacy: st.error("❌ 請務必勾選同意聲明！")
                            elif not p_name or not p_ext: st.warning("⚠️ 姓名與分機為必填！")
                            elif report_mode == "用油量申報 (含單筆/多筆/油卡)" and not f_files: st.error("⚠️ 請上傳佐證資料！")
                            else:
                                if data_entries[0]['vol'] <= 0 and report_mode == "用油量申報 (含單筆/多筆/油卡)": st.warning("⚠️ 第一筆加油量不能為 0。")
                                else:
                                    valid_logic = True; file_links = []
                                    my_bar = st.progress(0, text="資料處理中...")
                                    if report_mode == "用油量申報 (含單筆/多筆/油卡)":
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
                                            except Exception as e: st.error(f"上傳失敗: {e}"); valid_logic = False; break
                                    
                                    if valid_logic:
                                        final_links = "\n".join(file_links) if file_links else "無"
                                        my_bar.progress(50, text="寫入資料庫...")
                                        rows_to_append = []
                                        current_time = get_taiwan_time().strftime("%Y-%m-%d %H:%M:%S")
                                        shared_str = "是" if is_shared else "-"
                                        card_str = fuel_card_id if fuel_card_id else "-"
                                        for entry in data_entries:
                                            if entry['vol'] > 0 or report_mode == "無使用":
                                                rows_to_append.append([current_time, selected_dept, p_name, p_ext, selected_device, str(row.get('校內財產編號', '-')), str(row.get('原燃物料名稱', '-')), card_str, str(entry["date"]), entry["vol"], shared_str, note_input, final_links])
                                        if rows_to_append:
                                            ws_record.append_rows(rows_to_append)
                                            my_bar.progress(100, text="完成！"); time.sleep(0.5); my_bar.empty()
                                            st.success("✅ 申報成功！"); st.balloons(); st.session_state['reset_counter'] += 1; st.rerun()
                                        else: st.warning("⚠️ 無有效資料可寫入。")

    with tabs[1]: st.info("📊 外部看板功能維持不變")

# ------------------------------------------
# 👑 超級管理員專區 (V96.0: 穩定核心修復)
# ------------------------------------------
elif st.session_state['current_page'] == 'admin_dashboard' and username == 'admin':
    st.title("👑 超級管理員後台")
    
    # 1. 核心資料預處理 (Core Data Pipeline) - 這是解決所有報錯的關鍵！
    # 將原始資料 (df_records) 轉換為乾淨的 (df_clean) 供所有分頁使用
    df_clean = df_records.copy()
    if not df_clean.empty:
        # 強制型別轉換
        df_clean['加油量'] = pd.to_numeric(df_clean['加油量'], errors='coerce').fillna(0)
        df_clean['日期格式'] = pd.to_datetime(df_clean['加油日期'], errors='coerce')
        df_clean['年份'] = df_clean['日期格式'].dt.year.fillna(0).astype(int)
        df_clean['月份'] = df_clean['日期格式'].dt.month.fillna(0).astype(int)
        # 綁定設備類別
        if not df_equip.empty:
            device_map = pd.Series(df_equip['統計類別'].values, index=df_equip['設備名稱備註']).to_dict()
            df_clean['統計類別'] = df_clean['設備名稱備註'].map(device_map).fillna("其他/未分類")

    # 年度選擇器 (使用處理過的年份)
    all_years = sorted(df_clean['年份'][df_clean['年份']>0].unique(), reverse=True)
    if not all_years: all_years = [datetime.now().year]
    c_year, _ = st.columns([1, 3])
    selected_admin_year = c_year.selectbox("📅 請選擇檢視年度", all_years, index=0)
    
    # 全域篩選 (篩選出該年度資料)
    df_year = df_clean[df_clean['年份'] == selected_admin_year]

    admin_tabs = st.tabs(["📝 全校燃油設備總覽", "🔍 申報資料異動", "📊 動態管理儀表板"])

    # === Tab A: 全校燃油設備總覽 (KeyError 修復) ===
    with admin_tabs[0]:
        # 未申報篩選 (現在讀取的是 df_clean，有日期格式，不會報錯)
        with st.expander("🔍 篩選未申報名單 (點擊展開)", expanded=False):
            c_f1, c_f2 = st.columns(2)
            d_filter_start = c_f1.date_input("查詢起始日", date(selected_admin_year, 1, 1))
            d_filter_end = c_f2.date_input("查詢結束日", date.today())
            
            if st.button("開始篩選未申報單位"):
                # 使用 datetime.date 進行比較
                mask_period = (df_clean['日期格式'].dt.date >= d_filter_start) & (df_clean['日期格式'].dt.date <= d_filter_end)
                reported_equip = set(df_clean[mask_period]['設備名稱備註'].unique())
                
                all_equip_df = df_equip.copy()
                all_equip_df['是否有報'] = all_equip_df['設備名稱備註'].apply(lambda x: x in reported_equip)
                unreported_df = all_equip_df[~all_equip_df['是否有報']]
                
                if not unreported_df.empty:
                    st.error(f"🚩 期間 [{d_filter_start} ~ {d_filter_end}] 共有 {len(unreported_df)} 台設備未申報！")
                    for unit, group in unreported_df.groupby('填報單位'):
                        with st.container():
                            st.markdown(f"**🏢 {unit}** (未申報數: {len(group)})")
                            st.dataframe(group[['設備名稱備註', '保管人', '校內財產編號']], use_container_width=True)
                            st.divider()
                else: st.success("🎉 太棒了！該期間所有設備皆有申報紀錄。")

        # 統計卡片
        if not df_year.empty and not df_equip.empty:
            annual_sum = df_year.groupby('設備名稱備註')['加油量'].sum().reset_index()
            annual_sum.rename(columns={'加油量': '年度用油量'}, inplace=True)
            last_report = df_year.groupby('設備名稱備註')['日期格式'].max().reset_index()
            report_count = df_year.groupby('設備名稱備註').size().reset_index(name='申報次數')
            
            target_cols = ['設備編號', '設備名稱備註', '原燃物料名稱', '設備數量', '設備所屬單位/部門', '保管人', '設備詳細位置/樓層', '統計類別']
            existing_cols = [c for c in target_cols if c in df_equip.columns]
            
            summary_df = pd.merge(df_equip[existing_cols], annual_sum, on='設備名稱備註', how='left')
            summary_df = pd.merge(summary_df, last_report, on='設備名稱備註', how='left')
            summary_df = pd.merge(summary_df, report_count, on='設備名稱備註', how='left')
            summary_df['年度用油量'] = summary_df['年度用油量'].fillna(0)
            summary_df['申報次數'] = summary_df['申報次數'].fillna(0).astype(int)
            
            c_dl, _ = st.columns([1,4])
            csv_sum = summary_df.to_csv(index=False).encode('utf-8-sig')
            c_dl.download_button("⬇️ 下載完整統計報表 (CSV)", csv_sum, f"summary_{selected_admin_year}.csv", "text/csv")
            
            for category in DEVICE_ORDER:
                cat_df = summary_df[summary_df['統計類別'] == category]
                if not cat_df.empty:
                    header_color = MORANDI_COLORS.get(category, "#CFD8DC")
                    st.markdown(f"### 📂 {category}")
                    cols = st.columns(2)
                    for idx, row in cat_df.reset_index().iterrows():
                        col = cols[idx % 2] 
                        with col:
                            last_date = row['日期格式']
                            if pd.isna(last_date):
                                status_html = '<span class="status-badge status-warn">⚠️ 尚無紀錄</span>'
                                last_date_str = "無"
                            else:
                                days_diff = (datetime.now() - last_date).days
                                if days_diff > 180: status_html = f'<span class="status-badge status-warn">⚠️ 逾期未填</span>'
                                else: status_html = '' 
                                last_date_str = last_date.strftime("%Y-%m-%d")
                            fuel_type = "⛽" if "汽油" in str(row['原燃物料名稱']) else "🚛"
                            
                            card_html = f"""<div class="equip-card"><div class="equip-header" style="background-color: {header_color};"><div class="equip-title-group"><div class="equip-code">{row.get('設備編號','-')}</div><div class="equip-name">{row.get('設備名稱備註','-')}</div></div><div class="equip-fuel-group"><div class="equip-vol">{row['年度用油量']:,.2f}</div><span class="equip-fuel-type">{fuel_type} {row.get('原燃物料名稱','')} (公升)</span></div></div><div class="equip-body"><div class="equip-info">🏢 部門: {row.get('設備所屬單位/部門','-')} | 👤 保管人: {row.get('保管人','-')}<br>📍 位置: {row.get('設備詳細位置/樓層','-')} | 📊 數量: {row.get('設備數量','-')}</div><div class="equip-footer"><div class="last-date">最後申報日期: {last_date_str}</div><div style="display:flex; align-items:center;"><span class="count-text">申報次數: <b>{row['申報次數']}</b></span>{status_html}</div></div></div></div>"""
                            st.markdown(card_html, unsafe_allow_html=True)
        else: st.warning("尚無資料可供統計。")

    # === Tab B: 申報資料異動 (StreamlitAPIException 修復) ===
    with admin_tabs[1]:
        st.subheader("🔍 申報資料異動")
        df_display = df_year.copy()
        if not df_display.empty:
            # 強制轉換日期為 date 物件，避免編輯器報錯
            df_display['加油日期'] = pd.to_datetime(df_display['加油日期']).dt.date
            
            edited_df = st.data_editor(
                df_display, 
                column_config={
                    "佐證資料": st.column_config.LinkColumn("佐證", display_text="🔗"), 
                    "加油日期": st.column_config.DateColumn("日期", format="YYYY-MM-DD"), 
                    "加油量": st.column_config.NumberColumn("油量", format="%.2f"), 
                    "填報時間": st.column_config.TextColumn("填報時間", disabled=True)
                }, 
                num_rows="dynamic", 
                use_container_width=True, 
                key="record_editor_v96"
            )
            if st.button("💾 儲存變更", type="primary"):
                try:
                    ws_record.clear()
                    export_df = edited_df.copy()
                    export_df['加油日期'] = export_df['加油日期'].astype(str)
                    ws_record.update([export_df.columns.tolist()] + export_df.astype(str).values.tolist())
                    st.success("✅ 更新成功！"); st.cache_data.clear(); time.sleep(1); st.rerun()
                except Exception as e: st.error(f"更新失敗: {e}")
        else: st.info(f"{selected_admin_year} 年度尚無資料。")

    # === Tab C: 儀表板 (ValueError 修復) ===
    with admin_tabs[2]:
        st.subheader("📊 動態管理儀表板")
        if not df_year.empty:
            c_ctrl1, c_ctrl2 = st.columns(2)
            stat_mode = c_ctrl1.radio("統計模式", ["依設備統計", "依單位統計"], horizontal=True)
            filter_options = ["全部燃油設備"] + DEVICE_ORDER if stat_mode == "依設備統計" else ["全部單位"] + sorted([x for x in df_equip['填報單位'].unique() if x!='-' and x!='填報單位'])
            selected_filter = c_ctrl2.selectbox("選擇細項", filter_options)
            
            df_dash = df_year.copy()
            if stat_mode == "依設備統計" and selected_filter != "全部燃油設備": df_dash = df_dash[df_dash['統計類別'] == selected_filter]
            elif stat_mode == "依單位統計" and selected_filter != "全部單位": df_dash = df_dash[df_dash['填報單位'] == selected_filter]

            if not df_dash.empty:
                gas_sum = df_dash[df_dash['原燃物料名稱'].str.contains('汽油', na=False)]['加油量'].sum()
                diesel_sum = df_dash[df_dash['原燃物料名稱'].str.contains('柴油', na=False)]['加油量'].sum()
                total_sum = df_dash['加油量'].sum()
                total_co2 = (gas_sum * 0.0022) + (diesel_sum * 0.0027)
                k1, k2, k3, k4 = st.columns(4)
                k1.markdown(f"""<div class="kpi-card kpi-gas"><div class="kpi-title">⛽ 汽油總量</div><div class="kpi-value">{gas_sum:,.1f}</div></div>""", unsafe_allow_html=True)
                k2.markdown(f"""<div class="kpi-card kpi-diesel"><div class="kpi-title">🚛 柴油總量</div><div class="kpi-value">{diesel_sum:,.1f}</div></div>""", unsafe_allow_html=True)
                k3.markdown(f"""<div class="kpi-card kpi-total"><div class="kpi-title">💧 總用油量</div><div class="kpi-value">{total_sum:,.1f}</div></div>""", unsafe_allow_html=True)
                k4.markdown(f"""<div class="kpi-card kpi-co2"><div class="kpi-title">☁️ 碳排放量</div><div class="kpi-value">{total_co2:,.4f}</div></div>""", unsafe_allow_html=True)
                st.markdown("---")
                
                st.subheader("🏗️ 設備數量統計")
                df_asset = df_equip.copy()
                df_asset['油品大類'] = df_asset['原燃物料名稱'].apply(lambda x: '汽油' if '汽油' in x else ('柴油' if '柴油' in x else '其他'))
                df_asset['設備數量'] = pd.to_numeric(df_asset['設備數量'], errors='coerce').fillna(1)
                
                if stat_mode == "依設備統計":
                    if selected_filter == "全部燃油設備": fig_cnt = px.bar(df_asset.groupby(['統計類別', '油品大類'])['設備數量'].sum().reset_index(), x='統計類別', y='設備數量', color='油品大類', title="各類設備數量 (依油品堆疊)", text_auto=True, color_discrete_map={'汽油':'#1ABC9C', '柴油':'#E67E22', '其他':'#95A5A6'})
                    else:
                        df_sub = df_asset[df_asset['統計類別'] == selected_filter]
                        fig_cnt = px.bar(df_sub.groupby(['填報單位', '油品大類'])['設備數量'].sum().reset_index(), x='填報單位', y='設備數量', color='油品大類', title=f"{selected_filter} - 各單位數量分布", text_auto=True, color_discrete_map={'汽油':'#1ABC9C', '柴油':'#E67E22', '其他':'#95A5A6'})
                else:
                    if selected_filter == "全部單位": fig_cnt = px.bar(df_asset.groupby(['填報單位', '統計類別'])['設備數量'].sum().reset_index(), y='填報單位', x='設備數量', color='統計類別', orientation='h', title="各單位設備數量 (依類別堆疊)", height=800)
                    else:
                        df_sub = df_asset[df_asset['填報單位'] == selected_filter]
                        fig_cnt = px.bar(df_sub.groupby(['統計類別', '油品大類'])['設備數量'].sum().reset_index(), y='統計類別', x='設備數量', color='油品大類', orientation='h', barmode='group', title=f"{selected_filter} - 設備持有狀況", text_auto=True, color_discrete_map={'汽油':'#1ABC9C', '柴油':'#E67E22', '其他':'#95A5A6'})
                st.plotly_chart(fig_cnt, use_container_width=True)

                st.subheader("📈 單位用油量統計")
                view_fuel = st.radio("檢視油品", ["全部", "僅汽油", "僅柴油"], horizontal=True)
                df_use = df_dash.copy()
                df_use['油品大類'] = df_use['原燃物料名稱'].apply(lambda x: '汽油' if '汽油' in x else ('柴油' if '柴油' in x else '其他'))
                if view_fuel == "僅汽油": df_use = df_use[df_use['油品大類'] == '汽油']
                elif view_fuel == "僅柴油": df_use = df_use[df_use['油品大類'] == '柴油']
                
                # 關鍵修復: 繪圖前再次確保型別
                if stat_mode == "依設備統計": stack_col = '設備名稱備註'
                else: stack_col = '填報單位'
                df_use[stack_col] = df_use[stack_col].astype(str)
                
                months = list(range(1, 13))
                chart_data = df_use.groupby(['月份', '油品類別', stack_col])['加油量'].sum().reset_index()
                fig_use = px.bar(chart_data, x=['月份', '油品類別'], y='加油量', color=stack_col, title=f"每月用油趨勢 ({stack_col}堆疊)", text_auto='.1f', color_discrete_sequence=px.colors.qualitative.Prism)
                usage_totals = df_use.groupby(['月份', '油品類別'])['加油量'].sum().reset_index()
                fig_use.add_trace(go.Scatter(x=[usage_totals['月份'], usage_totals['油品類別']], y=usage_totals['加油量'], text=usage_totals['加油量'].apply(lambda x:f"{x:.1f}"), mode='text', textposition='top center', showlegend=False, textfont=dict(size=14)))
                fig_use.update_layout(xaxis=dict(title="月份 / 油品", tickmode='array', tickvals=list(range(1, 13)), ticktext=[f"{m}月" for m in months], tickfont=dict(size=14)), yaxis=dict(title="加油量 (公升)"), height=600, margin=dict(t=50, b=120))
                st.plotly_chart(fig_use, use_container_width=True)

                st.markdown("---")
                c_tree, c_sun = st.columns([3, 2])
                with c_tree:
                    st.subheader("🌲 碳排熱點分析 (Treemap)")
                    df_dash['CO2e'] = df_dash.apply(lambda r: r['加油量']*0.0022 if '汽油' in r['原燃物料名稱'] else r['加油量']*0.0027, axis=1)
                    if not df_dash.empty:
                        fig_tree = px.treemap(df_dash, path=['填報單位', '設備名稱備註'], values='CO2e', color='CO2e', color_continuous_scale='RdBu_r')
                        st.plotly_chart(fig_tree, use_container_width=True)
                with c_sun:
                    st.subheader("🍩 油品結構 (Sunburst)")
                    sun_data = df_dash.groupby(['填報單位', '油品大類'])['加油量'].sum().reset_index()
                    if not sun_data.empty:
                        fig_sun = px.sunburst(sun_data, path=['油品大類', '填報單位'], values='加油量', color='油品大類', color_discrete_map={'汽油':'#1ABC9C', '柴油':'#E67E22', '其他':'#95A5A6'})
                        st.plotly_chart(fig_sun, use_container_width=True)
            else: st.warning("在此篩選條件下無資料。")
        else: st.info("尚無該年度資料，無法顯示儀表板。")

    st.markdown('<div class="contact-footer">管理員系統版本 V96.0 (Stable Core)</div>', unsafe_allow_html=True)