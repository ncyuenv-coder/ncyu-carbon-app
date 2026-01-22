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
    /* 統一上傳區樣式 */
    [data-testid="stFileUploaderDropzone"] {
        background-color: #D6EAF8; border: 2px dashed #2E86C1; padding: 20px;
    }
    .note-text {color: #566573; font-weight: bold; font-size: 0.9rem;}
    
    /* 強制刷新按鈕樣式 */
    div.stButton > button[kind="secondary"] {
        color: #E74C3C; border-color: #E74C3C;
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
    
    # 讀取分頁 (請確認 Sheet 名稱已改為 "單位資訊")
    ws_units = sh_ref.worksheet("單位資訊")
    ws_buildings = sh_ref.worksheet("建築物清單")
    ws_types = sh_ref.worksheet("設備類型")
    ws_coef = sh_ref.worksheet("冷媒係數表")
    
    try: ws_records = sh_ref.worksheet("冷媒填報紀錄")
    except: 
        ws_records = sh_ref.add_worksheet(title="冷媒填報紀錄", rows="1000", cols="15")
        ws_records.append_row(["填報時間","填報人","填報人分機","校區","所屬單位","填報單位名稱","建築物名稱","辦公室編號","維修日期","設備類型","設備品牌型號","冷媒種類","冷媒填充量","備註","佐證資料"])

except Exception as e:
    st.error(f"❌ 資料庫連線失敗，請檢查 Google Sheet 是否存在 '單位資訊' 分頁。錯誤訊息: {e}")
    st.stop()

# 4. 資料讀取 (V212: 絕對位置鎖定 + 強制清洗)
@st.cache_data(ttl=60)
def load_ref_data_v212():
    def clean_text(text):
        if pd.isna(text): return ""
        text = str(text)
        # 強力清洗：轉半形、去頭尾空白
        text = unicodedata.normalize('NFKC', text).strip()
        return text

    def get_df_by_position(ws):
        # 讀取所有資料
        data = ws.get_all_values()
        if len(data) > 1:
            # 使用第 1 列當標題，但我們後續會用 iloc (位置) 來操作，不依賴標題名稱
            df = pd.DataFrame(data[1:], columns=data[0])
            # 清洗所有格子
            for col in df.columns:
                df[col] = df[col].apply(clean_text)
            return df
        return pd.DataFrame()
    
    return get_df_by_position(ws_units), get_df_by_position(ws_buildings), get_df_by_position(ws_types), get_df_by_position(ws_coef)

df_units, df_buildings, df_types, df_coef = load_ref_data_v212()

# 5. 頁面內容
st.title("❄️ 冷媒填報專區")

# 強制刷新按鈕 (如果 Sheet 改了但網頁沒變，按這個)
if st.button("🔄 刷新資料庫 (若選單沒更新請點此)", type="secondary"):
    st.cache_data.clear()
    st.rerun()

tabs = st.tabs(["📝 新增填報", "📊 動態查詢看板"])

with tabs[0]:
    with st.form("ref_form", clear_on_submit=True):
        st.subheader("填報人基本資料區")
        c1, c2, c3 = st.columns(3)
        
        # === V212 核心連動邏輯 (絕對位置 A->B->C) ===
        
        # 1. 校區 (鎖定 Column A / index 0)
        campuses = []
        if not df_units.empty:
            # iloc[:, 0] 代表第一欄
            campuses = sorted([x for x in df_units.iloc[:, 0].unique() if x])
        sel_campus = c1.selectbox("校區", campuses, index=None, placeholder="請選擇...")
        
        # 2. 所屬單位 (鎖定 Column B / index 1)
        depts = []
        if sel_campus and not df_units.empty:
            # 篩選條件：第1欄 == 選中的校區
            mask = df_units.iloc[:, 0] == sel_campus
            depts = sorted([x for x in df_units[mask].iloc[:, 1].unique() if x])
        sel_dept = c2.selectbox("所屬單位", depts, index=None, placeholder="請先選擇校區...")
        
        # 3. 填報單位名稱 (鎖定 Column C / index 2)
        units = []
        if sel_dept and not df_units.empty:
            # 篩選條件：第1欄 == 校區 AND 第2欄 == 單位
            mask = (df_units.iloc[:, 0] == sel_campus) & (df_units.iloc[:, 1] == sel_dept)
            units = sorted([x for x in df_units[mask].iloc[:, 2].unique() if x])
        sel_unit = c3.selectbox("填報單位名稱", units, index=None, placeholder="請先選擇所屬單位...")
        
        # ---------------------------------------------
        
        c4, c5 = st.columns(2)
        name = c4.text_input("填報人")
        ext = c5.text_input("填報人分機")
        
        st.markdown("---")
        st.subheader("詳細位置資訊區")
        c6, c7 = st.columns(2)
        
        # 建築物 (鎖定前兩欄：A=校區, B=建築物)
        builds = []
        if sel_campus and not df_buildings.empty:
            # 假設建築物清單也是 A欄=校區, B欄=建築物名稱
            mask_b = df_buildings.iloc[:, 0] == sel_campus
            builds = sorted([x for x in df_buildings[mask_b].iloc[:, 1].unique() if x])
            
        sel_build = c6.selectbox("建築物名稱", builds, index=None, placeholder="請先選擇上方校區...")
        office = c7.text_input("辦公室編號", placeholder="例如：404辦公室")
        
        st.markdown("---")
        st.subheader("設備修繕冷媒填充資訊區")
        c8, c9 = st.columns(2)
        r_date = c8.date_input("維修日期 (統一填寫發票日期)", datetime.today())
        
        # 設備類型 (A欄)
        e_types = []
        if not df_types.empty:
            e_types = sorted([x for x in df_types.iloc[:, 0].unique() if x])
        sel_etype = c9.selectbox("設備類型", e_types, index=None, placeholder="請選擇...")
        
        c10, c11 = st.columns(2)
        e_model = c10.text_input("設備品牌型號", placeholder="例如：國際 CS-100FL+CU-100FLC")
        
        # [cite_start]冷媒種類 (B欄, 冷媒係數表的第2欄) [cite: 1]
        r_types = []
        if not df_coef.empty and df_coef.shape[1] > 1:
            r_types = sorted([x for x in df_coef.iloc[:, 1].unique() if x])
        sel_rtype = c11.selectbox("冷媒種類", r_types, index=None, placeholder="請選擇...")
        
        amount = st.number_input("冷媒填充量 (公斤)", min_value=0.0, step=0.1, format="%.2f")
        
        st.markdown("請上傳冷媒填充單據佐證資料")
        f_file = st.file_uploader("上傳佐證 (必填)", type=['pdf', 'jpg', 'png'], label_visibility="collapsed")
        
        st.markdown("---")
        note = st.text_input("備註內容", placeholder="備註 (選填)")
        st.markdown('<div class="note-text">如有資料誤繕情形，請重新登錄1次資訊，並於備註欄填寫：「前筆資料誤繕，請刪除。」</div>', unsafe_allow_html=True)
        
        st.markdown("""
        <div style="background-color:#F8F9F9; padding:10px; border-radius:5px; font-size:0.9rem; margin-bottom:10px;">
        <strong>📜 個人資料蒐集聲明</strong><br>
        1. 蒐集目的：冷媒設備維修管理與碳盤查統計。<br>
        2. 利用期間：保存至填報年度後第二年1月1日。<br>
        3. 您有權依個資法請求查詢或刪除。
        </div>
        """, unsafe_allow_html=True)
        
        agree = st.checkbox("我已閱讀並同意個資聲明")
        
        submitted = st.form_submit_button("🚀 確認送出", use_container_width=True)
        
        if submitted:
            if not agree: st.error("❌ 請勾選同意聲明")
            elif not sel_campus or not sel_dept or not sel_unit: st.warning("⚠️ 請完整選擇填報單位資訊")
            elif not name or not ext: st.warning("⚠️ 請填寫填報人與分機")
            elif not sel_build: st.warning("⚠️ 請選擇建築物")
            elif not sel_etype or not sel_rtype: st.warning("⚠️ 請選擇設備類型與冷媒種類")
            elif not f_file: st.error("⚠️ 請上傳佐證資料")
            else:
                try:
                    f_file.seek(0); f_ext = f_file.name.split('.')[-1]
                    clean_name = f"{sel_campus}_{sel_dept}_{sel_unit}_{r_date}_{sel_etype}_{sel_rtype}.{f_ext}"
                    meta = {'name': clean_name, 'parents': [REF_FOLDER_ID]}
                    media = MediaIoBaseUpload(f_file, mimetype=f_file.type, resumable=True)
                    file = drive_service.files().create(body=meta, media_body=media, fields='webViewLink').execute()
                    link = file.get('webViewLink')
                    
                    current_time = get_taiwan_time().strftime("%Y-%m-%d %H:%M:%S")
                    row_data = [
                        current_time, name, ext, sel_campus, sel_dept, sel_unit, 
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