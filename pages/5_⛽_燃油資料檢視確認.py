import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import gspread
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import streamlit_authenticator as stauth
import time
import re
import io
import uuid
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from PIL import Image
from functools import wraps

# [PDF 截圖套件防護]
try:
    import fitz  # PyMuPDF
    HAS_FITZ = True
except ImportError:
    HAS_FITZ = False

# ==========================================
# 0. 系統設定 (查核專屬後台)
# ==========================================
st.set_page_config(page_title="燃油資料檢視確認", page_icon="🔍", layout="wide")

def get_taiwan_time():
    return datetime.utcnow() + timedelta(hours=8)

# ==========================================
# [防護機制] 徹底解決 _auth_request 憑證失效問題
# ==========================================
def with_retry(max_retries=5, base_delay=2.0, backoff_factor=2.0):
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            delay = base_delay
            for attempt in range(max_retries):
                try:
                    return func(*args, **kwargs)
                except Exception as e:
                    err_str = str(e).lower()
                    if "429" in err_str or "quota" in err_str or "rate limit" in err_str or "auth" in err_str:
                        if attempt < max_retries - 1:
                            time.sleep(delay + random.uniform(0.1, 1.0))
                            delay *= backoff_factor
                            continue
                    raise e
        return wrapper
    return decorator

# 每次寫入都重新獲取最新憑證，避免 Session 過期導致 'AuthorizedSession' 報錯
@with_retry(max_retries=3, base_delay=2.0)
def update_sheet_row_safe(worksheet_name, range_name, values):
    oauth = st.secrets["gcp_oauth"]
    creds = Credentials(token=None, refresh_token=oauth["refresh_token"], token_uri="https://oauth2.googleapis.com/token", client_id=oauth["client_id"], client_secret=oauth["client_secret"], scopes=["https://www.googleapis.com/auth/spreadsheets"])
    gc = gspread.authorize(creds)
    sh = gc.open_by_key("1gqDU21YJeBoBOd8rMYzwwZ45offXWPGEODKTF6B8k-Y")
    ws = sh.worksheet(worksheet_name)
    ws.batch_update([{'range': range_name, 'values': [values]}])

# ==========================================
# 1. CSS 樣式表 (莫蘭迪深色調)
# ==========================================
st.markdown("""
<style>
    :root {
        color-scheme: light;
        --orange-bg: #E67E22;     
        --orange-dark: #D35400;
        --text-main: #2C3E50;
        --morandi-dark: #2C3E50;
        --morandi-tab-bg: #34495E;
        --morandi-tab-hover: #1A252F;
    }
    [data-testid="stAppViewContainer"] { background-color: #F4F6F6; color: var(--text-main); }
    [data-testid="stHeader"] { background-color: rgba(0,0,0,0); }
    [data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 1px solid #BDC3C7; }
    
    div[data-baseweb="input"] > div, div[data-baseweb="base-input"] > input, textarea, input { background-color: #FFFFFF !important; border-color: #BDC3C7 !important; color: #000000 !important; font-size: 1.15rem !important; }
    div[data-baseweb="select"] > div { border-color: #BDC3C7 !important; background-color: #FFFFFF !important; }
    
    /* 莫蘭迪深色調 Tabs */
    div[data-baseweb="tab-list"] { background-color: transparent !important; gap: 5px; }
    div[data-baseweb="tab-list"] button { background-color: var(--morandi-tab-bg) !important; border-radius: 8px 8px 0 0 !important; padding: 12px 25px !important; border: 1px solid #2C3E50 !important; border-bottom: none !important; transition: all 0.3s; }
    div[data-baseweb="tab-list"] button div p { color: #ECF0F1 !important; font-size: 1.25rem !important; font-weight: 800 !important; margin: 0; }
    div[data-baseweb="tab-list"] button:hover { background-color: var(--morandi-tab-hover) !important; }
    div[data-baseweb="tab-list"] button[aria-selected="true"] { background-color: var(--morandi-tab-hover) !important; border-top: 4px solid #E67E22 !important; }
    div[data-baseweb="tab-list"] button[aria-selected="true"] div p { color: #F39C12 !important; }
    div[data-baseweb="tab-highlight"] { display: none; }

    button[kind="primary"] { background-color: var(--orange-bg) !important; color: #FFFFFF !important; border: 2px solid var(--orange-dark) !important; border-radius: 12px !important; font-size: 1.2rem !important; font-weight: 800 !important; padding: 0.6rem 1.2rem !important; box-shadow: 0 4px 6px rgba(230, 126, 34, 0.3) !important; width: 100%; }
    button[kind="primary"] p { color: #FFFFFF !important; } 
    button[kind="primary"]:hover { background-color: var(--orange-dark) !important; transform: translateY(-2px) !important; }
    
    /* 大標題：靠左對齊同一行 */
    .dashboard-main-title { font-size: 2.2rem; font-weight: 900; text-align: left; color: #1A5276; margin-bottom: 25px; margin-top: 10px; letter-spacing: 1px; }
    
    .stRadio div[role="radiogroup"] label { background-color: #D6EAF8 !important; border: 1px solid #AED6F1 !important; border-radius: 8px !important; padding: 8px 15px !important; margin-right: 10px !important; }
    .stRadio div[role="radiogroup"] label p { font-size: 1.15rem !important; font-weight: 800 !important; color: #1A5276 !important; }
    .stRadio div[role="radiogroup"] label[data-checked="true"] { background-color: #1A5276 !important; border-color: #1A5276 !important; }
    .stRadio div[role="radiogroup"] label[data-checked="true"] p { color: #FFFFFF !important; }
    
    /* 視覺化卡片樣式 */
    .db-card { background-color: #FFFFFF; border: 1px solid #BDC3C7; border-radius: 12px; padding: 0; box-shadow: 0 4px 10px rgba(0,0,0,0.05); margin-bottom: 25px; overflow: hidden; }
    .db-card-header { background-color: #2C3E50; color: #FFFFFF; padding: 12px 20px; font-size: 1.25rem; font-weight: 900; display: flex; justify-content: space-between; }
    .db-card-body { padding: 20px; background-color: #FDFEFE; }
    
    /* HTML表格樣式 */
    .visual-table { width: 100%; border-collapse: collapse; margin-top: 10px; background-color: #FFFFFF; font-size: 1.05rem; }
    .visual-table th { background-color: #EBF5FB; color: #2C3E50; padding: 10px; text-align: left; border: 1px solid #BDC3C7; }
    .visual-table td { padding: 10px; border: 1px solid #EAEDED; color: #34495E; }
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

# [嚴格限制] 僅開放管理者
if username != 'admin':
    st.error("🚫 權限不足：此頁面僅供系統管理員 (admin) 檢視與操作。")
    st.stop()

try:
    _raw_creds = st.secrets["credentials"]
    credentials_login = clean_secrets(_raw_creds)
    cookie_cfg = st.secrets["cookie"]
    authenticator = stauth.Authenticate(credentials_login, cookie_cfg["name"], cookie_cfg["key"], cookie_cfg["expiry_days"])
except: pass

with st.sidebar:
    st.header(f"👤 {name} (最高管理員)")
    st.success("☁️ 查核後台連線正常")
    authenticator.logout('登出系統', 'sidebar')

# ==========================================
# 3. 資料庫連線與資料載入
# ==========================================
SHEET_ID = "1gqDU21YJeBoBOd8rMYzwwZ45offXWPGEODKTF6B8k-Y" 
DEVICE_ORDER = ["公務車輛(GV-1-)", "乘坐式割草機(GV-2-)", "乘坐式農用機具(GV-3-)", "鍋爐(GS-1-)", "發電機(GS-2-)", "肩背或手持式割草機、吹葉機(GS-3-)", "肩背或手持式農用機具(GS-4-)"]
DEVICE_CODE_MAP = {"GV-1": "公務車輛(GV-1-)", "GV-2": "乘坐式割草機(GV-2-)", "GV-3": "乘坐式農用機具(GV-3-)", "GS-1": "鍋爐(GS-1-)", "GS-2": "發電機(GS-2-)", "GS-3": "肩背或手持式割草機、吹葉機(GS-3-)", "GS-4": "肩背或手持式農用機具(GS-4-)"}

# 僅獲取 Drive Service 用於下載圖片
@st.cache_resource
def init_google_drive():
    oauth = st.secrets["gcp_oauth"]
    creds = Credentials(token=None, refresh_token=oauth["refresh_token"], token_uri="https://oauth2.googleapis.com/token", client_id=oauth["client_id"], client_secret=oauth["client_secret"], scopes=["https://www.googleapis.com/auth/drive"])
    drive = build('drive', 'v3', credentials=creds)
    return drive

try:
    drive_service = init_google_drive()
except Exception as e: 
    st.error(f"雲端硬碟連線失敗: {e}")
    st.stop()

# 資料讀取 (獨立連線避免過期)
@st.cache_data(ttl=60) 
def load_fuel_data():
    oauth = st.secrets["gcp_oauth"]
    creds = Credentials(token=None, refresh_token=oauth["refresh_token"], token_uri="https://oauth2.googleapis.com/token", client_id=oauth["client_id"], client_secret=oauth["client_secret"], scopes=["https://www.googleapis.com/auth/spreadsheets"])
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SHEET_ID)
    
    try: ws_equip = sh.worksheet("設備清單") 
    except: ws_equip = sh.sheet1 
    try: ws_record = sh.worksheet("油料填報紀錄")
    except: ws_record = sh.worksheet("填報紀錄")
    
    eq_data = ws_equip.get_all_values()
    df_e = pd.DataFrame(eq_data[1:], columns=eq_data[0]) if len(eq_data) > 1 else pd.DataFrame(columns=eq_data[0])
    df_e['_row_index'] = range(2, len(df_e) + 2)
    if '設備編號' in df_e.columns: df_e['統計類別'] = df_e['設備編號'].apply(lambda c: next((v for k, v in DEVICE_CODE_MAP.items() if str(c).startswith(k)), "其他/未分類"))
    
    rec_data = ws_record.get_all_values()
    df_r = pd.DataFrame(rec_data[1:], columns=rec_data[0]) if len(rec_data) > 1 else pd.DataFrame(columns=rec_data[0])
    df_r['_row_index'] = range(2, len(df_r) + 2)
    
    return df_e, df_r, eq_data[0], rec_data[0]

df_equip_full, df_records, eq_cols, rec_cols = load_fuel_data()

# ==========================================
# 4. 輔助功能 (Email & 圖片下載)
# ==========================================
def send_system_email(to_email, subject, body):
    try:
        smtp_cfg = st.secrets.get("smtp", {})
        if not smtp_cfg: return False, "系統尚未設定 SMTP 參數"
        if not str(to_email).strip() or str(to_email).strip().lower() in ['nan', 'none', '']: return False, "無效的信箱"
        
        msg = MIMEMultipart()
        msg['From'] = smtp_cfg.get("email", "noreply@system.com")
        msg['To'] = str(to_email).strip()
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'html'))
        
        server = smtplib.SMTP(smtp_cfg["server"], smtp_cfg["port"])
        server.starttls()
        server.login(smtp_cfg["email"], smtp_cfg["password"])
        server.send_message(msg)
        server.quit()
        return True, "發送成功"
    except Exception as e: return False, f"發信失敗: {str(e)}"

@st.cache_data(ttl=86400, show_spinner=False, max_entries=100)
def get_cached_file_from_drive(url):
    try:
        d_svc = init_google_drive()
        match = re.search(r'/d/([a-zA-Z0-9_-]+)', url)
        if not match: match = re.search(r'id=([a-zA-Z0-9_-]+)', url)
        if not match: return None, False, None, None, False

        file_id = match.group(1)
        try:
            meta = d_svc.files().get(fileId=file_id, fields="name").execute()
            filename = meta.get('name', f'佐證資料_{file_id}')
        except: filename = f"佐證資料_{file_id}.pdf"
            
        is_pdf = filename.lower().endswith('.pdf')
        request = d_svc.files().get_media(fileId=file_id)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while done is False: status, done = downloader.next_chunk()
        
        fh.seek(0); file_bytes = fh.read()
        
        try:
            if is_pdf and HAS_FITZ:
                doc = fitz.open("pdf", file_bytes)
                images = []
                for page_num in range(len(doc)):
                    page = doc.load_page(page_num)
                    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                    images.append(Image.open(io.BytesIO(pix.tobytes("png"))))
                is_landscape = images[0].width > images[0].height if images else False
                return images, is_landscape, file_bytes, filename, True
            else:
                img = Image.open(io.BytesIO(file_bytes))
                return [img], img.width > img.height, file_bytes, filename, False
        except: return None, False, file_bytes, filename, is_pdf
    except Exception: return None, False, None, None, False

def render_image_gallery(links, base_key):
    if not links:
        st.info("無佐證圖片連結")
        return
    
    unique_links = list(dict.fromkeys([l.strip() for l_str in links for l in str(l_str).split('\n') if str(l).strip() not in ["", "無", "佐證如總務處事務組中油明細", "-"]]))
    if not unique_links:
        st.info("無佐證圖片或採用共用/統一明細。")
        return
        
    with st.spinner("🚀 載入圖片中..."):
        images_data = []
        for url in unique_links:
            images_list, is_landscape, file_bytes, filename, is_pdf = get_cached_file_from_drive(url)
            if images_list:
                for i, img in enumerate(images_list):
                    ps = f" (第 {i+1} 頁)" if is_pdf and len(images_list) > 1 else ""
                    images_data.append({ "url": url, "img": img, "is_landscape": is_landscape, "bytes": file_bytes, "filename": filename, "is_pdf": is_pdf, "ps": ps })
            else: images_data.append({ "url": url, "img": None, "is_landscape": False, "bytes": file_bytes, "filename": filename, "is_pdf": is_pdf, "ps": "" })
        
        idx = 0
        while idx < len(images_data):
            item = images_data[idx]
            unique_key = f"{base_key}_{idx}_{uuid.uuid4().hex[:6]}"
            if item['bytes'] is None:
                st.error("⚠️ 無法讀取檔案"); idx += 1
            elif item['img'] is None:
                st.download_button(label=f"📥 下載 {item['filename']}", data=item['bytes'], file_name=item['filename'], key=unique_key); idx += 1
            else:
                cap = f"📄 {item['filename']}{item['ps']}"
                if item['is_landscape']:
                    st.image(item['img'], use_container_width=True, caption=cap); idx += 1
                else:
                    c1, c2 = st.columns(2)
                    c1.image(item['img'], use_container_width=True, caption=cap); idx += 1
                    if idx < len(images_data) and images_data[idx]['img'] is not None and not images_data[idx]['is_landscape']:
                        c2.image(images_data[idx]['img'], use_container_width=True, caption=f"📄 {images_data[idx]['filename']}{images_data[idx]['ps']}"); idx += 1
            if idx < len(images_data): st.markdown("<hr style='margin:10px 0; border:1px dashed #E5E7E9;'>", unsafe_allow_html=True)

# ==========================================
# 5. 各分頁邏輯 (Fragment)
# ==========================================

@st.fragment
def render_tab1_db_manage(df_equip_full, eq_cols):
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("### 🗄️ 燃油設備單位資料庫管理")
    st.info("請於下方視覺化卡片中直接修改資料，系統會「精準更新」該筆設備，確保其他資料庫內容安全無虞。")
    
    all_years = sorted([y for y in df_equip_full['設備檢視年度'].unique() if str(y).strip() != ''], reverse=True) if '設備檢視年度' in df_equip_full.columns else [str(datetime.now().year)]
    all_units = sorted([u for u in df_equip_full['填報單位'].unique() if str(u).strip() != ''])
    
    c1, c2 = st.columns(2)
    sel_year = c1.selectbox("📅 設備年度", all_years, index=0)
    sel_unit = c2.selectbox("🏢 填報單位", all_units, index=0)
    
    if '設備檢視年度' in df_equip_full.columns:
        df_target = df_equip_full[(df_equip_full['設備檢視年度'].astype(str) == str(sel_year)) & (df_equip_full['填報單位'] == sel_unit)].copy()
    else:
        df_target = df_equip_full[df_equip_full['填報單位'] == sel_unit].copy()

    if not df_target.empty:
        st.markdown("---")
        for idx, row in df_target.iterrows():
            r_idx = row['_row_index']
            eq_name = row.get('設備名稱備註', '未命名設備')
            
            with st.container():
                st.markdown(f"<div class='db-card'><div class='db-card-header'><span>🚜 {eq_name}</span><span>編號: {row.get('設備編號', '-')}</span></div><div class='db-card-body'>", unsafe_allow_html=True)
                
                with st.form(key=f"db_form_{r_idx}"):
                    c_a1, c_a2, c_a3 = st.columns(3)
                    new_id = c_a1.text_input("設備編號", value=row.get('設備編號', ''), key=f"eid_{r_idx}")
                    new_name = c_a2.text_input("設備名稱備註", value=eq_name, key=f"enm_{r_idx}")
                    new_fuel = c_a3.text_input("原燃物料名稱", value=row.get('原燃物料名稱', ''), key=f"efl_{r_idx}")
                    
                    c_b1, c_b2, c_b3 = st.columns(3)
                    new_keeper = c_b1.text_input("保管人", value=row.get('保管人', ''), key=f"ekp_{r_idx}")
                    new_sub = c_b2.text_input("設備所屬單位/部門", value=row.get('設備所屬單位/部門', ''), key=f"esb_{r_idx}")
                    new_loc = c_b3.text_input("詳細位置/樓層", value=row.get('設備詳細位置/樓層', ''), key=f"elc_{r_idx}")
                    
                    c_c1, c_c2, c_c3 = st.columns(3)
                    new_qty = c_c1.text_input("設備數量", value=row.get('設備數量', ''), key=f"eqt_{r_idx}")
                    new_mail = c_c2.text_input("電子郵件", value=row.get('電子郵件', ''), key=f"eml_{r_idx}")
                    new_ast = c_c3.text_input("校內財產編號", value=row.get('校內財產編號', ''), key=f"eas_{r_idx}")
                    
                    if st.form_submit_button("💾 精準儲存此設備變更", type="primary"):
                        updated_vals = []
                        for col in eq_cols:
                            if col == "設備編號": updated_vals.append(new_id)
                            elif col == "設備名稱備註": updated_vals.append(new_name)
                            elif col == "原燃物料名稱": updated_vals.append(new_fuel)
                            elif col == "保管人": updated_vals.append(new_keeper)
                            elif col == "設備所屬單位/部門": updated_vals.append(new_sub)
                            elif col == "設備詳細位置/樓層": updated_vals.append(new_loc)
                            elif col == "設備數量": updated_vals.append(new_qty)
                            elif col == "電子郵件": updated_vals.append(new_mail)
                            elif col == "校內財產編號": updated_vals.append(new_ast)
                            elif col in row: updated_vals.append(str(row[col]))
                            else: updated_vals.append("")
                        
                        try:
                            range_str = f"A{r_idx}:{chr(65 + len(updated_vals) - 1)}{r_idx}"
                            with st.spinner("🔄 寫入資料庫中..."):
                                update_sheet_row_safe("設備清單", range_str, updated_vals)
                            st.success(f"✅ {new_name} 資料更新成功！")
                            st.cache_data.clear()
                            time.sleep(1)
                            st.rerun()
                        except Exception as e:
                            st.error(f"寫入失敗：{e}")
                
                st.markdown("</div></div>", unsafe_allow_html=True)
    else:
        st.warning("查無符合條件之設備資料。")

@st.fragment
def render_tab2_notify(df_records, df_equip_full):
    st.markdown("<br>", unsafe_allow_html=True)
    
    mode = st.radio("請選擇作業模式：", ["📅 每月申報提醒通知", "🔍 篩選未申報名單催報"], horizontal=True, label_visibility="collapsed")
    st.markdown("---")
    
    df_clean = df_records.copy()
    df_clean['加油量'] = pd.to_numeric(df_clean['加油量'], errors='coerce').fillna(0)
    df_clean['日期格式'] = pd.to_datetime(df_clean['加油日期'], errors='coerce')
    
    test_mode = st.radio("寄件模式", ["🧪 測試模式 (僅寄送至測試信箱)", "🚀 正式發送 (寄送至各單位真實信箱)"], horizontal=True)
    test_email = st.text_input("測試用收件信箱", "test@example.com")
    st.markdown("<br>", unsafe_allow_html=True)

    if "每月" in mode:
        c_y, c_m = st.columns(2)
        cur_year = datetime.now().year
        cur_month = datetime.now().month
        sel_year = c_y.number_input("檢視年度", value=cur_year, min_value=2020, max_value=2100, step=1)
        sel_month = c_m.selectbox("檢視月份", range(1, 13), index=cur_month-1)
        
        if '設備檢視年度' in df_equip_full.columns:
            df_eq = df_equip_full[df_equip_full['設備檢視年度'].astype(str) == str(sel_year)].copy()
        else: df_eq = df_equip_full.copy()
        
        df_target_rec = df_clean[(df_clean['日期格式'].dt.year == sel_year) & (df_clean['日期格式'].dt.month == sel_month)]
        reported_keys = set(df_target_rec['填報單位'].astype(str) + "|||" + df_target_rec['設備名稱備註'].astype(str))
        
        def get_report_info(r):
            key = str(r.get('填報單位', '')) + "|||" + str(r.get('設備名稱備註', ''))
            if key in reported_keys:
                recs = df_target_rec[(df_target_rec['填報單位'] == r['填報單位']) & (df_target_rec['設備名稱備註'] == r['設備名稱備註'])]
                dts = ", ".join(recs['加油日期'].astype(str).unique())
                vol = recs['加油量'].sum()
                return pd.Series(["✅ 已申報", dts, vol])
            else:
                return pd.Series(["❌ 未申報", "-", 0.0])
                
        df_eq[['申報狀態', '最近加油日期', '當月加油量']] = df_eq.apply(get_report_info, axis=1)
        
        st.markdown(f"#### 📅 【{sel_year}年 {sel_month}月】各單位設備填報情形視覺化總覽")
        
        for unit, group in df_eq.groupby('填報單位'):
            html_table = f"""<div style="background:#FDFEFE; border:1px solid #BDC3C7; border-radius:12px; padding:15px; margin-bottom:15px; box-shadow:0 2px 5px rgba(0,0,0,0.05);">
                <h4 style="color:#2C3E50; margin-top:0; border-bottom:2px solid #E67E22; padding-bottom:5px;">🏢 {unit}</h4>
                <table class="visual-table">
                <tr><th>狀態</th><th>設備編號</th><th>設備名稱</th><th>燃料</th><th>數量</th><th>日期</th><th>加油量</th></tr>"""
            
            for _, row in group.iterrows():
                color = "#148F77" if "✅" in row['申報狀態'] else "#C0392B"
                html_table += f"<tr><td style='color:{color}; font-weight:bold;'>{row['申報狀態']}</td><td>{row.get('設備編號','-')}</td><td>{row['設備名稱備註']}</td><td>{row.get('原燃物料名稱','-')}</td><td>{row.get('設備數量','1')}</td><td>{row['最近加油日期']}</td><td>{row['當月加油量']:.1f} L</td></tr>"
            html_table += "</table></div>"
            st.markdown(html_table, unsafe_allow_html=True)
        
        unreported_units = sorted(list(df_eq[df_eq['申報狀態'] == '❌ 未申報']['填報單位'].unique()))
        all_units = sorted(list(df_eq['填報單位'].unique()))
        
        st.markdown("---")
        st.markdown("#### 📧 自訂信件與寄送")
        send_target = st.radio("寄送對象選擇：", ["過濾寄送 (僅寄給「尚未完成申報」的單位)", "全面寄送 (強制寄給「所有」單位)"], horizontal=True)
        
        target_units = unreported_units if "過濾寄送" in send_target else all_units
        st.info(f"👉 預計發送單位數量：**{len(target_units)}** 個")
        
        default_subject = f"【提醒】{sel_year}年{sel_month}月份燃油設備油料使用申報通知，請於次月5日前完成申報"
        default_body = f"""您好：\n\n依據環境部「事業應盤查登錄溫室氣體排放量之排放源」及「溫室氣體排放量盤查登錄及查驗管理辦法」，本校須依辦理溫室氣體排放量盤查登錄作業。\n\n提醒您，請於每月5日前至校內溫室氣體盤查填報系統申報貴單位前一月份之燃油設備油料使用情形。\n\n若有疑問，請洽環境保護及安全管理中心林小姐 (分機7137)\n\n感謝您的配合！\n環境保護及安全管理中心 敬上"""
        
        sub_input = st.text_input("信件主旨", value=default_subject)
        body_input = st.text_area("信件內容 (不含動態設備清單，系統會自動附於尾部)", value=default_body, height=200)
        
        if st.button("🚀 確認並寄出通知信", type="primary"):
            if not target_units:
                st.success("所有單位皆已申報，無須寄信。")
            else:
                with st.spinner("正在發送信件..."):
                    success_count, fail_count = 0, 0
                    for unit in target_units:
                        unit_eqs = df_eq[(df_eq['填報單位'] == unit) & (df_eq['申報狀態'] == '❌ 未申報' if "過濾寄送" in send_target else True)]
                        emails = [test_email] if "測試模式" in test_mode else unit_eqs['電子郵件'].unique()
                        
                        for email in set(emails):
                            if str(email).strip() and str(email).strip() != 'nan':
                                eq_li = "".join([f"<li>{r['設備名稱備註']} ：{r.get('設備數量','1')}台</li>" for _, r in unit_eqs.iterrows()])
                                final_body = body_input.replace('\n', '<br>') + f"<br><br><p>貴單位 ({unit}) 負責之設備清單如下，請撥冗確認：</p><ul>{eq_li}</ul><br><span style='color:#7F8C8D;font-size:12px;'>(系統自動發送)</span>"
                                ok, msg = send_system_email(email, sub_input, final_body)
                                if ok: success_count += 1
                                else: fail_count += 1
                    
                    st.success(f"寄送完畢！成功: {success_count} 封，失敗/無效: {fail_count} 封。")

    else:
        c_f1, c_f2 = st.columns(2)
        d_start = c_f1.date_input("查詢起始日", date(datetime.now().year, 1, 1))
        d_end = c_f2.date_input("查詢結束日", date.today())
        
        target_year = d_end.year
        if '設備檢視年度' in df_equip_full.columns:
            df_eq = df_equip_full[df_equip_full['設備檢視年度'].astype(str) == str(target_year)].copy()
        else: df_eq = df_equip_full.copy()
        
        df_target_rec = df_clean[(df_clean['日期格式'].dt.date >= d_start) & (df_clean['日期格式'].dt.date <= d_end)]
        reported_keys = set(df_target_rec['填報單位'].astype(str) + "|||" + df_target_rec['設備名稱備註'].astype(str))
        
        df_eq['申報狀態'] = df_eq.apply(lambda r: "✅ 已申報" if (str(r.get('填報單位', '')) + "|||" + str(r.get('設備名稱備註', ''))) in reported_keys else "❌ 未申報", axis=1)
        df_eq['未申報期間'] = f"{d_start} ~ {d_end}"
        
        st.markdown(f"#### 🔍 【{d_start} ~ {d_end}】期間未申報篩選視覺化")
        
        for unit, group in df_eq.groupby('填報單位'):
            html_table = f"""<div style="background:#FDFEFE; border:1px solid #BDC3C7; border-radius:12px; padding:15px; margin-bottom:15px; box-shadow:0 2px 5px rgba(0,0,0,0.05);">
                <h4 style="color:#2C3E50; margin-top:0; border-bottom:2px solid #E67E22; padding-bottom:5px;">🏢 {unit}</h4>
                <table class="visual-table">
                <tr><th>狀態</th><th>設備編號</th><th>設備名稱</th><th>燃料</th><th>數量</th><th>查核期間</th></tr>"""
            
            for _, row in group.iterrows():
                color = "#148F77" if "✅" in row['申報狀態'] else "#C0392B"
                html_table += f"<tr><td style='color:{color}; font-weight:bold;'>{row['申報狀態']}</td><td>{row.get('設備編號','-')}</td><td>{row['設備名稱備註']}</td><td>{row.get('原燃物料名稱','-')}</td><td>{row.get('設備數量','1')}</td><td>{row['未申報期間']}</td></tr>"
            html_table += "</table></div>"
            st.markdown(html_table, unsafe_allow_html=True)
        
        unreported_units = sorted(list(df_eq[df_eq['申報狀態'] == '❌ 未申報']['填報單位'].unique()))
        
        st.markdown("---")
        st.markdown("#### 📧 自訂信件與寄送")
        st.info(f"👉 系統將自動發送給未申報的 **{len(unreported_units)}** 個單位。")
        
        default_subject = f"【催報提醒】查貴單位尚未申報({d_start} ~ {d_end})之燃油設備油料使用情形，請盡速申報"
        default_body = f"""您好：\n\n依據規定，本校須依辦理溫室氣體排放量盤查登錄作業。\n\n經系統比對，貴單位於 {d_start} 至 {d_end} 期間有設備尚未完成申報，請於通知日起3日內至系統完成補登。\n\n若有疑問，請洽環境保護及安全管理中心林小姐 (分機7137)\n\n感謝配合！"""
        
        sub_input = st.text_input("信件主旨", value=default_subject, key="u_sub")
        body_input = st.text_area("信件內容 (不含設備清單)", value=default_body, height=180, key="u_body")
        
        if st.button("🚀 針對未申報單位寄出催報信", type="primary"):
            if not unreported_units: st.success("🎉 太棒了！該期間全數設備皆已完成申報。")
            else:
                with st.spinner("正在發送催報信件..."):
                    success_count, fail_count = 0, 0
                    for unit in unreported_units:
                        unit_eqs = df_eq[(df_eq['填報單位'] == unit) & (df_eq['申報狀態'] == '❌ 未申報')]
                        emails = [test_email] if "測試模式" in test_mode else unit_eqs['電子郵件'].unique()
                        
                        for email in set(emails):
                            if str(email).strip() and str(email).strip() != 'nan':
                                eq_li = "".join([f"<li>{r['設備名稱備註']}</li>" for _, r in unit_eqs.iterrows()])
                                final_body = body_input.replace('\n', '<br>') + f"<br><br><p>貴單位 ({unit}) 尚未申報之設備：</p><ul>{eq_li}</ul><br><span style='color:#7F8C8D;font-size:12px;'>(系統自動發送)</span>"
                                ok, msg = send_system_email(email, sub_input, final_body)
                                if ok: success_count += 1
                                else: fail_count += 1
                    
                    st.success(f"寄送完畢！成功: {success_count} 封，失敗: {fail_count} 封。")


@st.fragment
def render_tab3_edit_records(df_records, ws_rec):
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("### 🔍 申報資料異動管理")
    
    df_clean = df_records.copy()
    df_clean['日期格式'] = pd.to_datetime(df_clean['加油日期'], errors='coerce')
    df_clean['年份'] = df_clean['日期格式'].dt.year.fillna(0).astype(int)
    df_clean['月份'] = df_clean['日期格式'].dt.month.fillna(0).astype(int)
    
    # [需求 4.2] 對齊統計類別 (如 GV-1)
    df_clean['統計類別'] = df_clean['設備編號'].apply(lambda c: next((v for k, v in DEVICE_CODE_MAP.items() if str(c).startswith(k)), "其他/未分類"))
    
    all_years = sorted(df_clean['年份'][df_clean['年份']>0].unique(), reverse=True) if not df_clean.empty else [datetime.now().year]
    
    c_y, c_m, c_c = st.columns(3)
    sel_yr = c_y.selectbox("📅 篩選年份", all_years, index=0)
    
    df_yr = df_clean[df_clean['年份'] == sel_yr]
    all_months = sorted(df_yr['月份'][df_yr['月份']>0].unique()) if not df_yr.empty else [1]
    sel_mo = c_m.selectbox("📆 篩選月份", all_months, index=0)
    
    df_mo = df_yr[df_yr['月份'] == sel_mo]
    
    # [需求 4.2] 設備類別下拉選單
    cats = ["全部"] + DEVICE_ORDER
    sel_cat = c_c.selectbox("🚜 設備類別", cats, index=0)
    
    st.markdown("---")
    
    if sel_cat != "全部":
        df_target = df_mo[df_mo['統計類別'] == sel_cat]
    else:
        df_target = df_mo
        
    if not df_target.empty:
        # [需求 4.3] 按照設備分組，左側列出該設備所有申報單，右側去重顯示圖片
        grouped = df_target.groupby('設備名稱備註')
        
        for eq_name, group in grouped:
            col_left, col_right = st.columns([4, 6], gap="large")
            
            with col_left:
                st.markdown(f"<div class='edit-card-header'>🚜 設備: {eq_name}</div>", unsafe_allow_html=True)
                
                for idx, row in group.iterrows():
                    row_idx_in_sheet = int(row['_row_index'])
                    
                    with st.expander(f"✏️ 編輯紀錄: {row.get('加油日期', '')} ({row.get('加油量', 0)} L)"):
                        with st.form(key=f"rec_form_{row_idx_in_sheet}"):
                            eq_id = st.text_input("設備編號 / 財產編號", value=row.get('校內財產編號', ''), key=f"id_{row_idx_in_sheet}")
                            eq_name_input = st.text_input("設備名稱備註", value=row.get('設備名稱備註', ''), key=f"nm_{row_idx_in_sheet}")
                            eq_qty = st.text_input("設備數量", value=row.get('設備數量', '1'), key=f"qty_{row_idx_in_sheet}")
                            eq_fuel = st.text_input("原燃物料名稱", value=row.get('原燃物料名稱', ''), key=f"fl_{row_idx_in_sheet}")
                            
                            try: d_val = pd.to_datetime(row['加油日期']).date()
                            except: d_val = date.today()
                            eq_date = st.date_input("加油日期", value=d_val, key=f"dt_{row_idx_in_sheet}")
                            
                            eq_vol = st.number_input("加油量", value=float(row.get('加油量', 0)), step=0.1, key=f"vl_{row_idx_in_sheet}")
                            eq_share = st.selectbox("與其他設備共用加油單", ["是", "否", "-"], index=0 if row.get('與其他設備共用加油單')=='是' else (1 if row.get('與其他設備共用加油單')=='否' else 2), key=f"sh_{row_idx_in_sheet}")
                            
                            eq_dept = st.text_input("填報單位", value=row.get('填報單位', ''), key=f"dp_{row_idx_in_sheet}")
                            c_u1, c_u2 = st.columns(2)
                            eq_user = c_u1.text_input("填報人", value=row.get('填報人', ''), key=f"ur_{row_idx_in_sheet}")
                            eq_ext = c_u2.text_input("分機", value=row.get('填報人分機', ''), key=f"ex_{row_idx_in_sheet}")
                            
                            eq_note = st.text_area("備註資訊", value=row.get('備註', ''), height=80, key=f"nt_{row_idx_in_sheet}")
                            
                            submit_btn = st.form_submit_button("💾 精準更新此筆紀錄", type="primary")
                            if submit_btn:
                                orig_cols = rec_cols
                                updated_vals = []
                                for col in orig_cols:
                                    if col == "設備名稱備註": updated_vals.append(eq_name_input)
                                    elif col == "校內財產編號": updated_vals.append(eq_id)
                                    elif col == "設備數量": updated_vals.append(eq_qty)
                                    elif col == "原燃物料名稱": updated_vals.append(eq_fuel)
                                    elif col == "加油日期": updated_vals.append(str(eq_date))
                                    elif col == "加油量": updated_vals.append(str(eq_vol))
                                    elif col == "與其他設備共用加油單": updated_vals.append(eq_share)
                                    elif col == "填報單位": updated_vals.append(eq_dept)
                                    elif col == "填報人": updated_vals.append(eq_user)
                                    elif col == "填報人分機": updated_vals.append(eq_ext)
                                    elif col == "備註": updated_vals.append(eq_note)
                                    elif col in row: updated_vals.append(str(row[col]))
                                    else: updated_vals.append("")
                                
                                try:
                                    range_str = f"A{row_idx_in_sheet}:{chr(65 + len(updated_vals) - 1)}{row_idx_in_sheet}"
                                    with st.spinner("🔄 寫入資料庫中..."):
                                        update_sheet_row_safe("油料填報紀錄", range_str, updated_vals)
                                    st.success("✅ 更新成功！")
                                    st.cache_data.clear()
                                    time.sleep(1)
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"寫入失敗：{e}")

            with col_right:
                # 統一收集該設備所有明細的圖片，並智慧去重
                raw_links = []
                for links_str in group['佐證資料'].dropna():
                    if links_str and str(links_str).strip() not in ["無", "-"]:
                        raw_links.extend([l.strip() for l in re.split(r'[\n,]', str(links_str)) if l.strip()])
                        
                unique_links = list(dict.fromkeys(raw_links))
                render_image_gallery(unique_links, base_key=f"img_grp_{eq_name}")
            
            st.markdown("<hr style='border: 1px solid #BDC3C7; margin: 30px 0;'>", unsafe_allow_html=True)
    else:
        st.warning("該條件下無申報紀錄。")


# ==========================================
# 6. 主程式入口
# ==========================================
def main():
    st.markdown('<div class="dashboard-main-title">🔍 燃油資料檢視與確認專區 <span style="font-size: 1.5rem; color: #5D6D7E; font-weight: 600;">(Data Verification & Audit Area)</span></div>', unsafe_allow_html=True)
    
    admin_tabs = st.tabs([
        "🗄️ 燃油設備單位資料庫管理",
        "⚠️ 寄送通知與未申報篩選", 
        "🔍 申報資料異動"
    ])
    
    with admin_tabs[0]: render_tab1_db_manage(df_equip_full, eq_cols)
    with admin_tabs[1]: render_tab2_notify(df_records, df_equip_full)
    with admin_tabs[2]: render_tab3_edit_records(df_records, ws_record)

if __name__ == "__main__":
    main()