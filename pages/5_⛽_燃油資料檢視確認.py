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

# [防護機制] 智慧重試裝甲
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
                    if "429" in err_str or "quota" in err_str or "rate limit" in err_str:
                        if attempt < max_retries - 1:
                            time.sleep(delay + random.uniform(0.1, 1.0))
                            delay *= backoff_factor
                            continue
                    raise e
        return wrapper
    return decorator

@with_retry(max_retries=3, base_delay=2.0)
def update_sheet_row(worksheet, range_name, values):
    """精準更新單一列，避免覆蓋整張表"""
    worksheet.update(range_name, [values])

# ==========================================
# 1. CSS 樣式表 (莫蘭迪深色 Tab 系統)
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
    div[data-baseweb="tab-list"] button { background-color: var(--morandi-tab-bg) !important; border-radius: 8px 8px 0 0 !important; padding: 10px 25px !important; border: 1px solid #2C3E50 !important; border-bottom: none !important; transition: all 0.3s; }
    div[data-baseweb="tab-list"] button div p { color: #ECF0F1 !important; font-size: 1.25rem !important; font-weight: 800 !important; margin: 0; }
    div[data-baseweb="tab-list"] button:hover { background-color: var(--morandi-tab-hover) !important; }
    div[data-baseweb="tab-list"] button[aria-selected="true"] { background-color: var(--morandi-tab-hover) !important; border-top: 4px solid #E67E22 !important; }
    div[data-baseweb="tab-list"] button[aria-selected="true"] div p { color: #F39C12 !important; }
    div[data-baseweb="tab-highlight"] { display: none; }

    button[kind="primary"] { background-color: var(--orange-bg) !important; color: #FFFFFF !important; border: 2px solid var(--orange-dark) !important; border-radius: 12px !important; font-size: 1.2rem !important; font-weight: 800 !important; padding: 0.6rem 1.2rem !important; box-shadow: 0 4px 6px rgba(230, 126, 34, 0.3) !important; width: 100%; }
    button[kind="primary"] p { color: #FFFFFF !important; } 
    button[kind="primary"]:hover { background-color: var(--orange-dark) !important; transform: translateY(-2px) !important; }
    
    .dashboard-main-title { font-size: 2.4rem; font-weight: 900; text-align: center; color: #1A5276; margin-bottom: 25px; margin-top: 10px; letter-spacing: 2px; }
    .stRadio div[role="radiogroup"] label { background-color: #D6EAF8 !important; border: 1px solid #AED6F1 !important; border-radius: 8px !important; padding: 8px 15px !important; margin-right: 10px !important; }
    .stRadio div[role="radiogroup"] label p { font-size: 1.15rem !important; font-weight: 800 !important; color: #1A5276 !important; }
    .stRadio div[role="radiogroup"] label[data-checked="true"] { background-color: #1A5276 !important; border-color: #1A5276 !important; }
    .stRadio div[role="radiogroup"] label[data-checked="true"] p { color: #FFFFFF !important; }
    
    /* 編輯卡片樣式 */
    .edit-card { background-color: #FFFFFF; border: 1px solid #BDC3C7; border-radius: 12px; padding: 20px; box-shadow: 0 4px 10px rgba(0,0,0,0.05); margin-bottom: 25px; }
    .edit-card-header { font-size: 1.3rem; font-weight: 900; color: #2C3E50; border-bottom: 2px solid #E67E22; padding-bottom: 10px; margin-bottom: 15px; }
    
    /* 狀態標籤 */
    .status-ok { background-color: #D4EFDF; color: #148F77; padding: 4px 12px; border-radius: 12px; font-weight: bold; font-size: 0.95rem; }
    .status-ng { background-color: #FADBD8; color: #C0392B; padding: 4px 12px; border-radius: 12px; font-weight: bold; font-size: 0.95rem; }
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
DEVICE_CODE_MAP = {"GV-1": "公務車輛(GV-1-)", "GV-2": "乘坐式割草機(GV-2-)", "GV-3": "乘坐式農用機具(GV-3-)", "GS-1": "鍋爐(GS-1-)", "GS-2": "發電機(GS-2-)", "GS-3": "肩背或手持式割草機、吹葉機(GS-3-)", "GS-4": "肩背或手持式農用機具(GS-4-)"}

@st.cache_resource
def init_google_fuel():
    oauth = st.secrets["gcp_oauth"]
    creds = Credentials(token=None, refresh_token=oauth["refresh_token"], token_uri="https://oauth2.googleapis.com/token", client_id=oauth["client_id"], client_secret=oauth["client_secret"], scopes=["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/spreadsheets"])
    gc = gspread.authorize(creds); drive = build('drive', 'v3', credentials=creds)
    return gc, drive

try:
    gc_conn, drive_service = init_google_fuel()
except Exception as e: 
    st.error(f"雲端連線失敗: {e}")
    st.stop()

@st.cache_data(ttl=120) # 極短快取，確保資料隨時最新
def load_fuel_data():
    sh = gc_conn.open_by_key(SHEET_ID)
    try: ws_equip = sh.worksheet("設備清單") 
    except: ws_equip = sh.sheet1 
    try: ws_record = sh.worksheet("油料填報紀錄")
    except: ws_record = sh.worksheet("填報紀錄")
    
    eq_data = ws_equip.get_all_values()
    df_e = pd.DataFrame(eq_data[1:], columns=eq_data[0]) if len(eq_data) > 1 else pd.DataFrame(columns=eq_data[0])
    # 加入 row_index 以利後續精準寫入 (Gsheet 第一行為標題，資料從第二行開始)
    df_e['_row_index'] = range(2, len(df_e) + 2)
    if '設備編號' in df_e.columns: df_e['統計類別'] = df_e['設備編號'].apply(lambda c: next((v for k, v in DEVICE_CODE_MAP.items() if str(c).startswith(k)), "其他/未分類"))
    
    rec_data = ws_record.get_all_values()
    df_r = pd.DataFrame(rec_data[1:], columns=rec_data[0]) if len(rec_data) > 1 else pd.DataFrame(columns=rec_data[0])
    df_r['_row_index'] = range(2, len(df_r) + 2)
    
    return df_e, df_r, ws_equip, ws_record

df_equip_full, df_records, ws_equip, ws_record = load_fuel_data()

# ==========================================
# 4. 輔助功能 (Email & 圖片下載)
# ==========================================
def send_system_email(to_email, subject, body):
    try:
        smtp_cfg = st.secrets.get("smtp", {})
        if not smtp_cfg: return False, "系統尚未設定 SMTP 參數 (st.secrets['smtp'])"
        if not str(to_email).strip() or str(to_email).strip().lower() in ['nan', 'none', '']: return False, "無效的電子郵件地址"
        
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
    except Exception as e: return False, f"失敗: {str(e)}"

@st.cache_data(ttl=86400, show_spinner=False, max_entries=100)
def get_cached_file_from_drive(url):
    try:
        _, d_svc = init_google_fuel()
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
    except Exception as e: return None, False, None, None, False

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
def render_tab1_db_manage(df_equip_full, ws_eq):
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("### 🗄️ 燃油設備單位資料庫管理")
    st.info("請篩選條件後，直接於下方表格雙擊欄位進行修改。系統將會「精準更新」有異動的列，不影響其他資料。")
    
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
        # 指定欲顯示/編輯的欄位
        edit_cols = ['設備編號', '設備名稱備註', '原燃物料名稱', '保管人', '設備所屬單位/部門', '設備詳細位置/樓層', '設備數量', '電子郵件', '校內財產編號']
        
        # 確保原本的 row_index 在 df_target 裡面，但不顯示給使用者修改
        disp_cols = [c for c in edit_cols if c in df_target.columns]
        df_show = df_target[['_row_index'] + disp_cols].reset_index(drop=True)
        
        st.markdown(f"**篩選結果：共 {len(df_show)} 筆設備**")
        edited_df = st.data_editor(
            df_show, 
            column_config={"_row_index": None}, # 隱藏內部列編號
            use_container_width=True, 
            num_rows="fixed",
            key="db_editor"
        )
        
        if st.button("💾 儲存資料庫異動", type="primary"):
            try:
                # 比對找出有差異的列
                changed_rows = []
                orig_cols = ws_eq.row_values(1) # 取得原本 Sheet 所有的欄位標題順序
                
                for idx in range(len(df_show)):
                    orig_row = df_show.iloc[idx]
                    new_row = edited_df.iloc[idx]
                    if not orig_row[disp_cols].equals(new_row[disp_cols]):
                        row_idx = int(orig_row['_row_index'])
                        
                        # 構建該列完整的更新資料 (只替換被編輯的欄位，其他欄位保留原 df_equip_full 內的舊值)
                        full_orig_data = df_equip_full[df_equip_full['_row_index'] == row_idx].iloc[0]
                        updated_values = []
                        for col_name in orig_cols:
                            if col_name in disp_cols:
                                updated_values.append(str(new_row[col_name]))
                            elif col_name in full_orig_data:
                                updated_values.append(str(full_orig_data[col_name]))
                            else:
                                updated_values.append("")
                        
                        changed_rows.append((row_idx, updated_values))
                
                if changed_rows:
                    with st.spinner("🔄 正在精準寫入異動資料..."):
                        for r_idx, vals in changed_rows:
                            # 進行 API 寫入，例如寫入 A2:Z2
                            range_str = f"A{r_idx}:{chr(65 + len(vals) - 1)}{r_idx}"
                            update_sheet_row(ws_eq, range_str, vals)
                    
                    st.success(f"✅ 成功精準更新 {len(changed_rows)} 筆設備資料！")
                    st.cache_data.clear()
                    time.sleep(1)
                    st.rerun()
                else:
                    st.info("無任何資料被修改。")
            except Exception as e:
                st.error(f"寫入失敗：{e}")
    else:
        st.warning("查無符合條件之設備資料。")

@st.fragment
def render_tab2_notify(df_records, df_equip_full):
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("### ⚠️ 寄送通知與未申報篩選")
    
    mode = st.radio("請選擇作業模式：", ["📅 每月申報提醒通知 (依月份檢視)", "🔍 篩選未申報名單催報 (依日期區間)"], horizontal=True, label_visibility="collapsed")
    st.markdown("---")
    
    df_clean = df_records.copy()
    df_clean['加油量'] = pd.to_numeric(df_clean['加油量'], errors='coerce').fillna(0)
    df_clean['日期格式'] = pd.to_datetime(df_clean['加油日期'], errors='coerce')
    
    if "每月" in mode:
        st.markdown("#### 📅 每月申報填報情形總覽與發信")
        c_y, c_m = st.columns(2)
        cur_year = datetime.now().year
        cur_month = datetime.now().month
        sel_year = c_y.number_input("檢視年度", value=cur_year, min_value=2020, max_value=2100, step=1)
        sel_month = c_m.selectbox("檢視月份", range(1, 13), index=cur_month-1)
        
        # 篩選當月應申報的設備 (依照年份對應設備庫)
        if '設備檢視年度' in df_equip_full.columns:
            df_eq = df_equip_full[df_equip_full['設備檢視年度'].astype(str) == str(sel_year)].copy()
        else: df_eq = df_equip_full.copy()
        
        # 抓取該月申報紀錄
        df_target_rec = df_clean[(df_clean['日期格式'].dt.year == sel_year) & (df_clean['日期格式'].dt.month == sel_month)]
        reported_keys = set(df_target_rec['填報單位'].astype(str) + "|||" + df_target_rec['設備名稱備註'].astype(str))
        
        # 標註申報狀態與加油量資訊
        def get_report_info(r):
            key = str(r.get('填報單位', '')) + "|||" + str(r.get('設備名稱備註', ''))
            if key in reported_keys:
                recs = df_target_rec[(df_target_rec['填報單位'] == r['填報單位']) & (df_target_rec['設備名稱備註'] == r['設備名稱備註'])]
                dts = ", ".join(recs['加油日期'].astype(str).unique())
                vol = recs['加油量'].sum()
                return pd.Series(["✅ 已申報", dts, vol])
            else:
                return pd.Series(["❌ 未申報", "-", 0.0])
                
        df_eq[['申報狀態', '最近加油日期', '當月累計加油量']] = df_eq.apply(get_report_info, axis=1)
        
        st.markdown(f"**【{sel_year}年 {sel_month}月】各單位設備填報情形**")
        show_cols = ['申報狀態', '設備編號', '設備名稱備註', '原燃物料名稱', '填報單位', '設備數量', '最近加油日期', '當月累計加油量']
        df_show = df_eq[[c for c in show_cols if c in df_eq.columns]].sort_values(['申報狀態', '填報單位'])
        
        st.dataframe(df_show, use_container_width=True)
        
        unreported_units = sorted(list(df_eq[df_eq['申報狀態'] == '❌ 未申報']['填報單位'].unique()))
        all_units = sorted(list(df_eq['填報單位'].unique()))
        
        st.markdown("---")
        st.markdown("#### 📧 編輯與寄送通知信")
        send_target = st.radio("寄送對象選擇：", ["僅寄給「尚未完成申報」的單位", "強制寄給「所有」單位"], horizontal=True)
        
        target_units = unreported_units if "尚未完成" in send_target else all_units
        st.info(f"👉 預計發送單位數量：**{len(target_units)}** 個")
        
        default_subject = f"【提醒】{sel_year}年{sel_month}月份燃油設備油料使用申報通知，請於次月5日前完成申報"
        default_body = f"""您好：\n\n依據環境部「事業應盤查登錄溫室氣體排放量之排放源」及「溫室氣體排放量盤查登錄及查驗管理辦法」，本校須依辦理溫室氣體排放量盤查登錄作業。\n\n提醒您，請於每月5日前至校內溫室氣體盤查填報系統申報貴單位前一月份之燃油設備油料使用情形。\n\n若有疑問，請洽環境保護及安全管理中心林小姐 (分機7137)\n\n感謝您的配合！\n環境保護及安全管理中心 敬上"""
        
        sub_input = st.text_input("信件主旨", value=default_subject)
        body_input = st.text_area("信件內容 (不含動態設備清單，系統會自動將清單附在內容後方)", value=default_body, height=200)
        
        if st.button("🚀 確認並寄出通知信", type="primary"):
            if not target_units:
                st.success("所有單位皆已申報，無須寄信。")
            else:
                with st.spinner("正在發送信件..."):
                    success_count, fail_count = 0, 0
                    for unit in target_units:
                        unit_eqs = df_eq[(df_eq['填報單位'] == unit) & (df_eq['申報狀態'] == '❌ 未申報' if "尚未完成" in send_target else True)]
                        emails = unit_eqs['電子郵件'].unique()
                        for email in emails:
                            if str(email).strip() and str(email).strip() != 'nan':
                                eq_li = "".join([f"<li>{r['設備名稱備註']} ：{r.get('設備數量','1')}台</li>" for _, r in unit_eqs.iterrows()])
                                final_body = body_input.replace('\n', '<br>') + f"<br><br><p>貴單位 ({unit}) 負責之設備清單如下，請撥冗確認：</p><ul>{eq_li}</ul><br><span style='color:#7F8C8D;font-size:12px;'>(系統自動發送)</span>"
                                ok, msg = send_system_email(email, sub_input, final_body)
                                if ok: success_count += 1
                                else: fail_count += 1
                    
                    st.success(f"寄送完畢！成功: {success_count} 封，失敗/無效信箱: {fail_count} 封。")

    else:
        st.markdown("#### 🔍 依起訖日篩選未申報名單與催報")
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
        
        st.markdown(f"**【{d_start} ~ {d_end}】期間設備填報情形**")
        show_cols = ['申報狀態', '設備編號', '設備名稱備註', '原燃物料名稱', '填報單位', '設備數量', '未申報期間']
        df_show = df_eq[[c for c in show_cols if c in df_eq.columns]].sort_values(['申報狀態', '填報單位'])
        st.dataframe(df_show, use_container_width=True)
        
        unreported_units = sorted(list(df_eq[df_eq['申報狀態'] == '❌ 未申報']['填報單位'].unique()))
        
        st.markdown("---")
        st.markdown("#### 📧 編輯與寄送催報信")
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
                        emails = unit_eqs['電子郵件'].unique()
                        for email in emails:
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
    st.markdown("### 🔍 申報資料異動 (視覺化 4:6 比對區)")
    
    df_clean = df_records.copy()
    df_clean['日期格式'] = pd.to_datetime(df_clean['加油日期'], errors='coerce')
    df_clean['年份'] = df_clean['日期格式'].dt.year.fillna(0).astype(int)
    df_clean['月份'] = df_clean['日期格式'].dt.month.fillna(0).astype(int)
    
    all_years = sorted(df_clean['年份'][df_clean['年份']>0].unique(), reverse=True) if not df_clean.empty else [datetime.now().year]
    
    c_y, c_m, c_c = st.columns(3)
    sel_yr = c_y.selectbox("篩選年份", all_years, index=0)
    
    df_yr = df_clean[df_clean['年份'] == sel_yr]
    all_months = sorted(df_yr['月份'][df_yr['月份']>0].unique()) if not df_yr.empty else [1]
    sel_mo = c_m.selectbox("篩選月份", all_months, index=0)
    
    df_mo = df_yr[df_yr['月份'] == sel_mo]
    if '設備名稱備註' in df_mo.columns:
        cats = ["全部"] + sorted(df_mo['設備名稱備註'].unique().tolist())
    else: cats = ["全部"]
    sel_cat = c_c.selectbox("設備名稱備註", cats, index=0)
    
    st.markdown("---")
    
    if sel_cat != "全部":
        df_target = df_mo[df_mo['設備名稱備註'] == sel_cat]
    else:
        df_target = df_mo
        
    if not df_target.empty:
        # 按照設備編號 (或名稱) 依序顯示
        df_target = df_target.sort_values(by=['設備名稱備註', '加油日期'], ascending=[True, False])
        
        for idx, row in df_target.iterrows():
            row_idx_in_sheet = int(row['_row_index'])
            
            with st.container():
                col_left, col_right = st.columns([4, 6], gap="large")
                
                with col_left:
                    st.markdown(f"<div class='edit-card-header'>⚙️ 編輯: {row.get('設備名稱備註', '未命名設備')}</div>", unsafe_allow_html=True)
                    
                    with st.form(key=f"edit_form_{row_idx_in_sheet}"):
                        eq_id = st.text_input("設備編號 / 財產編號", value=row.get('校內財產編號', ''), key=f"id_{row_idx_in_sheet}")
                        eq_name = st.text_input("設備名稱備註", value=row.get('設備名稱備註', ''), key=f"nm_{row_idx_in_sheet}")
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
                        
                        submit_btn = st.form_submit_button("💾 精準更新此筆資料", type="primary")
                        if submit_btn:
                            orig_cols = ws_rec.row_values(1)
                            updated_vals = []
                            # Mapping based on typical layout. Must strictly match columns.
                            for col in orig_cols:
                                if col == "設備名稱備註": updated_vals.append(eq_name)
                                elif col == "校內財產編號": updated_vals.append(eq_id)
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
                                update_sheet_row(ws_rec, range_str, updated_vals)
                                st.success("✅ 更新成功！")
                                st.cache_data.clear()
                                time.sleep(1)
                                st.rerun()
                            except Exception as e:
                                st.error(f"寫入失敗：{e}")

                with col_right:
                    links_str = row.get('佐證資料', '')
                    render_image_gallery([links_str], base_key=f"img_{row_idx_in_sheet}")
                
                st.markdown("<hr style='border: 1px solid #BDC3C7; margin: 30px 0;'>", unsafe_allow_html=True)
    else:
        st.warning("該條件下無申報紀錄。")


# ==========================================
# 6. 主程式入口
# ==========================================
def main():
    st.markdown('<div class="dashboard-main-title">🔍 燃油資料檢視與確認專區<br><span style="font-size: 1.3rem; color: #5D6D7E; font-weight: 600;">Data Verification & Audit Area</span></div>', unsafe_allow_html=True)
    
    admin_tabs = st.tabs([
        "🗄️ 燃油設備單位資料庫管理",
        "⚠️ 寄送通知與未申報篩選", 
        "🔍 申報資料異動 (4:6視覺比對)"
    ])
    
    # 執行三個 Fragment
    with admin_tabs[0]: render_tab1_db_manage(df_equip_full, ws_equip)
    with admin_tabs[1]: render_tab2_notify(df_records, df_equip_full)
    with admin_tabs[2]: render_tab3_edit_records(df_records, ws_record)

if __name__ == "__main__":
    main()