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
</style>
""", unsafe_allow_html=True)

# 2. 身份驗證
if st.session_state.get("authentication_status") is not True:
    st.warning("🔒 請先至首頁 (Hello) 登入系統")
    st.stop()

# 3. 資料庫連線設定
REF_SHEET_ID = "1ZdvMBkprsN9w6EUKeGU_KYC8UKeS0rmX1Nq0yXzESIc"
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
    
    # 讀取分頁 (使用新定義的 Sheet 名稱)
    ws_units = sh_ref.worksheet("單位資訊")
    ws_buildings = sh_ref.worksheet("建築物清單")
    ws_types = sh_ref.worksheet("設備類型")
    ws_coef = sh_ref.worksheet("冷媒係數表")
    
    try: ws_records = sh_ref.worksheet("冷媒填報紀錄")
    except: 
        ws_records = sh_ref.add_worksheet(title="冷媒填報紀錄", rows="1000", cols="15")
        ws_records.append_row(["填報時間","填報人","填報人分機","校區","所屬單位","填報單位名稱","建築物名稱","辦公室編號","維修日期","設備類型","設備品牌型號","冷媒種類","冷媒填充量","備註","佐證資料"])

except Exception as e:
    st.error(f"❌ 資料庫連線失敗: {e}")
    st.stop()

# 4. 資料讀取 (V219: 強制 2 欄位鎖定)
@st.cache_data(ttl=60)
def load_ref_data_v219():
    def clean_text(text):
        if pd.isna(text): return ""
        text = str(text)
        # 強力清洗：轉半形、去頭尾空白
        text = unicodedata.normalize('NFKC', text).strip()
        return text

    def get_df_by_position(ws):
        data = ws.get_all_values()
        if len(data) > 1:
            # 不管標題叫什麼，我們只認位置
            # 使用第一列當作 DataFrame 的欄位名稱
            df = pd.DataFrame(data[1:], columns=data[0])
            # 清洗所有內容
            for col in df.columns:
                df[col] = df[col].apply(clean_text)
            return df
        return pd.DataFrame()
    
    return get_df_by_position(ws_units), get_df_by_position(ws_buildings), get_df_by_position(ws_types), get_df_by_position(ws_coef)

df_units, df_buildings, df_types, df_coef = load_ref_data_v219()

# 5. 頁面內容
st.title("❄️ 冷媒填報專區")

if st.button("🔄 刷新選單資料", type="secondary"):
    st.cache_data.clear()
    st.rerun()

tabs = st.tabs(["📝 新增填報", "📊 動態查詢看板"])

with tabs[0]:
    with st.form("ref_form", clear_on_submit=True):
        
        # === 區塊 1: 填報人基本資訊區 (2層連動：單位 -> 名稱) ===
        st.markdown('<div class="section-header">1. 填報人基本資訊區</div>', unsafe_allow_html=True)
        
        c1, c2 = st.columns(2)
        
        # 1-1. 所屬單位 (鎖定第 1 欄)
        unit_depts = []
        if not df_units.empty:
            # iloc[:, 0] = A欄
            unit_depts = sorted([x for x in df_units.iloc[:, 0].unique() if x])
        sel_dept = c1.selectbox("所屬單位", unit_depts, index=None, placeholder="請選擇單位...")
        
        # 1-2. 填報單位名稱 (鎖定第 2 欄, 依 A 欄篩選)
        unit_names = []
        if sel_dept and not df_units.empty:
            # 篩選 A 欄 == 選中單位
            mask = df_units.iloc[:, 0] == sel_dept
            # 取 B 欄 (iloc[:, 1])
            if df_units.shape[1] >= 2:
                unit_names = sorted([x for x in df_units[mask].iloc[:, 1].unique() if x])
        sel_unit_name = c2.selectbox("填報單位名稱", unit_names, index=None, placeholder="請先選擇所屬單位...")
        
        # 1-3. 開放欄位
        c3, c4 = st.columns(2)
        name = c3.text_input("填報人")
        ext = c4.text_input("填報人分機")
        
        st.markdown("---")
        
        # === 區塊 2: 詳細位置資訊區 (2層連動：校區 -> 建築物) ===
        st.markdown('<div class="section-header">2. 詳細位置資訊區</div>', unsafe_allow_html=True)
        c6, c7 = st.columns(2)
        
        # 2-1. 填報單位所在校區 (鎖定建築物清單第 1 欄)
        loc_campuses = []
        if not df_buildings.empty:
            loc_campuses = sorted([x for x in df_buildings.iloc[:, 0].unique() if x])
        sel_loc_campus = c6.selectbox("填報單位所在校區", loc_campuses, index=None, placeholder="請選擇校區...")
        
        # 2-2. 建築物名稱 (鎖定建築物清單第 2 欄, 依 A 欄篩選)
        buildings = []
        if sel_loc_campus and not df_buildings.empty:
            mask_b = df_buildings.iloc[:, 0] == sel_loc_campus
            if df_buildings.shape[1] >= 2:
                buildings = sorted([x for x in df_buildings[mask_b].iloc[:, 1].unique() if x])
        sel_build = c6.selectbox("建築物名稱", buildings, index=None, placeholder="請先選擇校區...")
        
        # 2-3. 辦公室編號
        office = c7.text_input("辦公室編號", placeholder="例如：202辦公室、306研究室")
        
        st.markdown("---")
        
        # === 區塊 3: 設備修繕資訊 ===
        st.markdown('<div class="section-header">3. 設備修繕冷媒填充資訊區</div>', unsafe_allow_html=True)
        c8, c9 = st.columns(2)
        r_date = c8.date_input("維修日期 (統一填寫發票日期)", datetime.today())
        
        # 設備類型 (A欄)
        e_types = []
        if not df_types.empty:
            e_types = sorted([x for x in df_types.iloc[:, 0].unique() if x])
        sel_etype = c9.selectbox("設備類型", e_types, index=None, placeholder="請選擇...")
        
        c10, c11 = st.columns(2)
        e_model = c10.text_input("設備品牌型號", placeholder="例如：國際 CS-100FL+CU-100FLC")
        
        # 冷媒種類 (B欄 - 因為您提供的係數表 csv A欄是代碼，B欄是名稱)
        r_types = []
        if not df_coef.empty and df_coef.shape[1] > 1:
            r_types = sorted([x for x in df_coef.iloc[:, 1].unique() if x])
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
                    # 檔名邏輯：使用位置校區
                    clean_name = f"{sel_loc_campus}_{sel_dept}_{sel_unit_name}_{r_date}_{sel_etype}_{sel_rtype}.{f_ext}"
                    
                    meta = {'name': clean_name, 'parents': [REF_FOLDER_ID]}
                    media = MediaIoBaseUpload(f_file, mimetype=f_file.type, resumable=True)
                    file = drive_service.files().create(body=meta, media_body=media, fields='webViewLink').execute()
                    link = file.get('webViewLink')
                    
                    current_time = get_taiwan_time().strftime("%Y-%m-%d %H:%M:%S")
                    
                    # 寫入資料庫
                    # 注意：校區欄位填入的是 sel_loc_campus (詳細位置的校區)
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