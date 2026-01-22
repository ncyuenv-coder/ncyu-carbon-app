import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import gspread
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import unicodedata

# ==========================================
# 0. 系統設定
# ==========================================
st.set_page_config(page_title="冷媒填報 - 嘉義大學", page_icon="❄️", layout="wide")

def get_taiwan_time():
    return datetime.utcnow() + timedelta(hours=8)

# ==========================================
# 1. CSS 樣式
# ==========================================
st.markdown("""
<style>
    [data-testid="stFileUploaderDropzone"] {
        background-color: #D6EAF8; border: 2px dashed #2E86C1; padding: 20px;
    }
    .note-text {color: #566573; font-weight: bold; font-size: 0.9rem;}
    .section-header {
        font-size: 1.15rem; font-weight: 800; color: #2C3E50; 
        border-left: 5px solid #E67E22; padding-left: 10px; margin-top: 20px; margin-bottom: 10px;
    }
    .debug-success {
        background-color: #D4EFDF; border: 1px solid #27AE60; 
        padding: 10px; border-radius: 5px; color: #196F3D; font-weight: bold; margin-bottom: 10px;
    }
</style>
""", unsafe_allow_html=True)

# 2. 身份驗證
if st.session_state.get("authentication_status") is not True:
    st.warning("🔒 請先至首頁 (Hello) 登入系統")
    st.stop()

# 3. 資料庫連線
# ✅ 已更新為您提供的最新 ID
REF_SHEET_ID = "1p7GsW-nrjerXhnn3pNgZzu_CdIh1Yxsm-fLJDqQ6MqA"
REF_FOLDER_ID = "1o0S56OyStDjvC5tgBWiUNqNjrpXuCQMI"

@st.cache_resource
def init_google_ref():
    oauth = st.secrets["gcp_oauth"]
    creds = Credentials(token=None, refresh_token=oauth["refresh_token"], token_uri="https://oauth2.googleapis.com/token", client_id=oauth["client_id"], client_secret=oauth["client_secret"], scopes=["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/spreadsheets"])
    gc = gspread.authorize(creds); drive = build('drive', 'v3', credentials=creds)
    return gc, drive

try:
    gc, drive_service = init_google_ref()
    sh_ref = gc.open_by_key(REF_SHEET_ID)
    
    # 嘗試讀取分頁 (使用標準名稱)
    ws_units = sh_ref.worksheet("單位資訊")
    ws_buildings = sh_ref.worksheet("建築物清單")
    ws_types = sh_ref.worksheet("設備類型")
    ws_coef = sh_ref.worksheet("冷媒係數表")
    
    try: ws_records = sh_ref.worksheet("冷媒填報紀錄")
    except: 
        ws_records = sh_ref.add_worksheet(title="冷媒填報紀錄", rows="1000", cols="15")
        ws_records.append_row(["填報時間","填報人","填報人分機","校區","所屬單位","填報單位名稱","建築物名稱","辦公室編號","維修日期","設備類型","設備品牌型號","冷媒種類","冷媒填充量","備註","佐證資料"])

except Exception as e:
    st.error(f"❌ 資料庫連線失敗: {e}。請檢查 REF_SHEET_ID 是否正確，或分頁名稱是否為 '單位資訊' (無空格)。")
    st.stop()

# 4. 資料讀取 (V225: 絕對位置強制讀取 - 無視標題)
@st.cache_data(ttl=0)
def load_ref_data_v225():
    def clean_text(text):
        if pd.isna(text): return ""
        text = str(text)
        return unicodedata.normalize('NFKC', text).strip()

    def get_df_by_position(ws):
        # 抓取所有資料
        data = ws.get_all_values()
        if len(data) > 1:
            # 直接把資料轉成 DataFrame，跳過第一列 (假設是標題，但我們不用標題來索引)
            # 強制只取前兩欄，並命名為 0 和 1 (整數索引)
            # 這樣不管標題叫什麼，df[0] 就是第一欄，df[1] 就是第二欄
            
            # 先確認是否有資料
            rows = data[1:]
            
            # 建立暫存 list 來確保每行都有 2 欄 (補齊空值)
            normalized_rows = []
            for row in rows:
                if len(row) >= 2:
                    normalized_rows.append([row[0], row[1]])
                elif len(row) == 1:
                    normalized_rows.append([row[0], ""])
                else:
                    normalized_rows.append(["", ""])
            
            df = pd.DataFrame(normalized_rows, columns=[0, 1])
            
            # 清洗內容
            for col in df.columns:
                df[col] = df[col].apply(clean_text)
                
            return df
        return pd.DataFrame()
    
    # 針對設備類型與係數表，我們也用同樣邏輯
    # 設備類型通常只有 1 欄
    def get_df_single_col(ws):
        data = ws.get_all_values()
        if len(data) > 1:
            # 取第一欄
            rows = [row[0] for row in data[1:] if row]
            df = pd.DataFrame(rows, columns=[0])
            df[0] = df[0].apply(clean_text)
            return df
        return pd.DataFrame()

    return get_df_by_position(ws_units), get_df_by_position(ws_buildings), get_df_single_col(ws_types), get_df_by_position(ws_coef)

df_units, df_buildings, df_types, df_coef = load_ref_data_v225()

# 5. 頁面內容
st.title("❄️ 冷媒填報專區")

# 強制刷新按鈕
if st.button("🔄 刷新資料庫 (更新後請點此)", type="primary"):
    st.cache_data.clear()
    st.rerun()

# --- 簡單診斷 (確認有讀到資料) ---
if not df_units.empty:
    # 檢查第一筆資料是否為空，若正常則不顯示紅字
    pass
else:
    st.error("⚠️ 【單位資訊】讀取為空！請檢查 Google Sheet。")

tabs = st.tabs(["📝 新增填報", "📊 動態查詢看板"])

with tabs[0]:
    with st.form("ref_form", clear_on_submit=True):
        
        # === 區塊 1: 填報人基本資訊區 ===
        st.markdown('<div class="section-header">1. 填報人基本資訊區</div>', unsafe_allow_html=True)
        
        c1, c2 = st.columns(2)
        
        # 1-1. 所屬單位 (強制讀取第 1 欄 / Index 0)
        unit_depts = []
        if not df_units.empty:
            # 0 代表第一欄
            unit_depts = sorted([x for x in df_units[0].unique() if x])
        sel_dept = c1.selectbox("所屬單位", unit_depts, index=None, placeholder="請選擇單位...")
        
        # 1-2. 填報單位名稱 (強制讀取第 2 欄 / Index 1，依第 1 欄篩選)
        unit_names = []
        if sel_dept and not df_units.empty:
            # 篩選邏輯：第1欄 == 選中的單位
            mask = df_units[0] == sel_dept
            unit_names = sorted([x for x in df_units[mask][1].unique() if x])
            
        sel_unit_name = c2.selectbox("填報單位名稱", unit_names, index=None, placeholder="請先選擇所屬單位...")
        
        # 1-3. 開放欄位
        c3, c4 = st.columns(2)
        name = c3.text_input("填報人")
        ext = c4.text_input("填報人分機")
        
        st.markdown("---")
        
        # === 區塊 2: 詳細位置資訊區 ===
        st.markdown('<div class="section-header">2. 詳細位置資訊區</div>', unsafe_allow_html=True)
        
        c6, c7 = st.columns(2)
        
        # 2-1. 填報單位所在校區 (強制讀取建築物清單 第 1 欄 / Index 0)
        loc_campuses = []
        if not df_buildings.empty:
            loc_campuses = sorted([x for x in df_buildings[0].unique() if x])
        sel_loc_campus = c6.selectbox("填報單位所在校區", loc_campuses, index=None, placeholder="請選擇校區...")
        
        # 2-2. 建築物名稱 (強制讀取建築物清單 第 2 欄 / Index 1)
        buildings = []
        if sel_loc_campus and not df_buildings.empty:
            mask_b = df_buildings[0] == sel_loc_campus
            buildings = sorted([x for x in df_buildings[mask_b][1].unique() if x])
        sel_build = c6.selectbox("建築物名稱", buildings, index=None, placeholder="請先選擇校區...")
        
        # 2-3. 辦公室
        office = c7.text_input("辦公室編號", placeholder="例如：202辦公室、306研究室")
        
        st.markdown("---")
        
        # === 區塊 3: 設備修繕資訊 ===
        st.markdown('<div class="section-header">3. 設備修繕冷媒填充資訊區</div>', unsafe_allow_html=True)
        c8, c9 = st.columns(2)
        r_date = c8.date_input("維修日期 (統一填寫發票日期)", datetime.today())
        
        # 設備類型 (強制讀取第 1 欄)
        e_types = []
        if not df_types.empty:
            e_types = sorted([x for x in df_types[0].unique() if x])
        sel_etype = c9.selectbox("設備類型", e_types, index=None, placeholder="請選擇...")
        
        c10, c11 = st.columns(2)
        e_model = c10.text_input("設備品牌型號", placeholder="例如：國際 CS-100FL+CU-100FLC")
        
        # 冷媒種類 (強制讀取係數表 第 2 欄 / Index 1 - 依據您提供的係數表 CSV，名稱在第 2 欄)
        r_types = []
        if not df_coef.empty:
            # 如果係數表有 2 欄以上，取第 2 欄；否則取第 1 欄
            target_idx = 1 if df_coef.shape[1] > 1 else 0
            r_types = sorted([x for x in df_coef[target_idx].unique() if x])
            
        sel_rtype = c11.selectbox("冷媒種類", r_types, index=None, placeholder="請選擇...")
        
        amount = st.number_input("冷媒填充量 (公斤)", min_value=0.0, step=0.1, format="%.2f")
        
        st.markdown("請上傳冷媒填充單據佐證資料")
        f_file = st.file_uploader("上傳佐證 (必填)", type=['pdf', 'jpg', 'png'], label_visibility="collapsed")
        
        st.markdown("---")
        note = st.text_input("備註內容", placeholder="備註 (選填)")
        
        st.markdown('<div style="background-color:#F8F9F9; padding:10px; font-size:0.9rem;"><strong>📜 個資聲明</strong>：蒐集目的為設備管理與碳盤查，保存至申報後第二年。</div>', unsafe_allow_html=True)
        agree = st.checkbox("我已閱讀並同意個資聲明")
        
        submitted = st.form_submit_button("🚀 確認送出", use_container_width=True)
        
        if submitted:
            # 必填檢查
            if not agree: st.error("❌ 請勾選同意聲明")
            elif not sel_dept or not sel_unit_name: st.warning("⚠️ 請完整選擇【基本資訊】中的單位資訊")
            elif not name or not ext: st.warning("⚠️ 請填寫填報人與分機")
            elif not sel_loc_campus or not sel_build: st.warning("⚠️ 請完整選擇【位置資訊】中的校區與建築物")
            elif not sel_etype or not sel_rtype: st.warning("⚠️ 請選擇設備類型與冷媒種類")
            elif not f_file: st.error("⚠️ 請上傳佐證資料")
            else:
                try:
                    f_file.seek(0); f_ext = f_file.name.split('.')[-1]
                    # 檔名邏輯
                    clean_name = f"{sel_loc_campus}_{sel_dept}_{sel_unit_name}_{r_date}_{sel_etype}_{sel_rtype}.{f_ext}"
                    
                    meta = {'name': clean_name, 'parents': [REF_FOLDER_ID]}
                    media = MediaIoBaseUpload(f_file, mimetype=f_file.type, resumable=True)
                    file = drive_service.files().create(body=meta, media_body=media, fields='webViewLink').execute()
                    link = file.get('webViewLink')
                    
                    current_time = get_taiwan_time().strftime("%Y-%m-%d %H:%M:%S")
                    
                    # 寫入資料
                    row_data = [
                        current_time, name, ext, sel_loc_campus, sel_dept, sel_unit_name, 
                        sel_build, office, str(r_date), sel_etype, e_model, 
                        sel_rtype, amount, note, link
                    ]
                    ws_records.append_row(row_data)
                    st.success("✅ 冷媒填報成功！")
                    st.balloons()
                except Exception as e:
                    st.error(f"上傳或寫入失敗: {e}")

with tabs[1]:
    st.info("🚧 動態查詢看板開發中...")