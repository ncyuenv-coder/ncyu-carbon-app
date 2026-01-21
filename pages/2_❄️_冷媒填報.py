import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import gspread
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import numpy as np

# 0. 系統設定
st.set_page_config(page_title="冷媒填報 - 嘉義大學", page_icon="❄️", layout="wide")

def get_taiwan_time():
    return datetime.utcnow() + timedelta(hours=8)

# 1. 樣式 (獨立定義，不影響燃油)
st.markdown("""
<style>
    [data-testid="stFileUploaderDropzone"] {background-color: #D6EAF8; border: 2px dashed #2E86C1; padding: 20px;}
    .note-text {color: #566573; font-weight: bold; font-size: 0.9rem;}
</style>
""", unsafe_allow_html=True)

# 2. 身份驗證
if st.session_state.get("authentication_status") is not True:
    st.warning("請先至首頁登入系統")
    st.stop()

# 3. 資料庫連線 (冷媒專用)
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
    
    # 讀取必要分頁 (純讀取，不建立)
    ws_units = sh_ref.worksheet("單位資訊") # V208 更名確認
    ws_buildings = sh_ref.worksheet("建築物清單")
    ws_types = sh_ref.worksheet("設備類型")
    ws_coef = sh_ref.worksheet("冷媒係數表")
    
    # 寫入分頁 (可自動建立)
    try: ws_records = sh_ref.worksheet("冷媒填報紀錄")
    except: 
        ws_records = sh_ref.add_worksheet(title="冷媒填報紀錄", rows="1000", cols="15")
        ws_records.append_row(["填報時間","填報人","填報人分機","校區","所屬單位","填報單位名稱","建築物名稱","辦公室編號","維修日期","設備類型","設備品牌型號","冷媒種類","冷媒填充量","備註","佐證資料"])

except Exception as e:
    st.error(f"❌ 冷媒資料庫連線失敗: {e}")
    st.stop()

# 4. 資料讀取 (除錯模式)
@st.cache_data(ttl=60)
def load_ref_data_debug():
    def get_df(ws):
        data = ws.get_all_values()
        if len(data) > 1:
            # 強制將所有資料轉為字串並去除前後空白
            df = pd.DataFrame(data[1:], columns=data[0]).astype(str)
            df.columns = df.columns.str.strip() # 清洗標題
            df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x) # 清洗內容
            return df
        return pd.DataFrame()
    
    return get_df(ws_units), get_df(ws_buildings), get_df(ws_types), get_df(ws_coef)

df_units, df_buildings, df_types, df_coef = load_ref_data_debug()

# 5. 介面開始
st.title("❄️ 冷媒填報專區")

# --- 除錯擴充功能 (請務必打開檢查) ---
with st.expander("🛠️ 開發者除錯區 (若選單空白請點此檢查資料)", expanded=False):
    st.write("目前讀取到的【單位資訊】表單前 5 筆：")
    if not df_units.empty:
        st.dataframe(df_units.head())
        st.write(f"欄位名稱: {df_units.columns.tolist()}")
    else:
        st.error("⚠️ 讀不到資料！請確認 Google Sheet 是否有資料。")

tabs = st.tabs(["📝 新增填報", "📊 動態查詢看板"])

with tabs[0]:
    with st.form("ref_form", clear_on_submit=True):
        st.subheader("填報人基本資料區")
        c1, c2, c3 = st.columns(3)
        
        # 連動邏輯 (基於除錯後的資料結構)
        # 假設 A欄=校區, B欄=所屬單位, C欄=填報單位名稱
        
        # 1. 校區 (取所有不重複值)
        # 為了保險，我們直接用欄位索引 (iloc)
        campuses = sorted(df_units.iloc[:, 0].unique()) if not df_units.empty else []
        sel_campus = c1.selectbox("校區", campuses, index=None, placeholder="請選擇...")
        
        # 2. 所屬單位
        depts = []
        if sel_campus and not df_units.empty:
            # 篩選 A欄 == 選中校區，取 B欄 (iloc[:, 1])
            depts = sorted(df_units[df_units.iloc[:, 0] == sel_campus].iloc[:, 1].unique())
        sel_dept = c2.selectbox("所屬單位", depts, index=None, placeholder="先選校區...")
        
        # 3. 填報單位名稱
        units = []
        if sel_dept and not df_units.empty:
            # 篩選 A欄==校區 & B欄==部門，取 C欄 (iloc[:, 2])
            units = sorted(df_units[
                (df_ref_units.iloc[:, 0] == sel_campus) & 
                (df_units.iloc[:, 1] == sel_dept)
            ].iloc[:, 2].unique())
        sel_unit = c3.selectbox("填報單位名稱", units, index=None, placeholder="先選單位...")
        
        c4, c5 = st.columns(2)
        name = c4.text_input("填報人")
        ext = c5.text_input("填報人分機")
        
        st.markdown("---")
        st.subheader("詳細位置資訊區")
        c6, c7 = st.columns(2)
        
        # 建築物連動
        builds = []
        if sel_campus and not df_buildings.empty:
            # 假設 A欄=校區, B欄=建築物
            builds = sorted(df_buildings[df_buildings.iloc[:, 0] == sel_campus].iloc[:, 1].unique())
        sel_build = c6.selectbox("建築物名稱", builds, index=None, placeholder="先選校區...")
        office = c7.text_input("辦公室編號", placeholder="例：404辦公室")
        
        st.markdown("---")
        st.subheader("設備修繕冷媒填充資訊區")
        c8, c9 = st.columns(2)
        r_date = c8.date_input("維修日期", datetime.today())
        
        e_types = sorted(df_types.iloc[:, 0].unique()) if not df_types.empty else []
        sel_etype = c9.selectbox("設備類型", e_types, index=None)
        
        c10, c11 = st.columns(2)
        e_model = c10.text_input("設備品牌型號")
        
        # 冷媒種類 (B欄, index 1)
        r_types = []
        if not df_coef.empty and df_coef.shape[1] > 1:
            r_types = sorted(df_coef.iloc[:, 1].unique())
        sel_rtype = c11.selectbox("冷媒種類", r_types, index=None)
        
        amount = st.number_input("冷媒填充量 (kg)", min_value=0.0, step=0.1)
        
        st.markdown("請上傳冷媒填充單據佐證資料")
        f_file = st.file_uploader("上傳佐證", type=['pdf', 'jpg', 'png'], label_visibility="collapsed")
        
        st.markdown("---")
        note = st.text_input("備註 (選填)")
        
        agree = st.checkbox("我已閱讀並同意個資聲明")
        
        submitted = st.form_submit_button("🚀 確認送出", use_container_width=True)
        
        if submitted:
            if not agree: st.error("請勾選同意聲明")
            elif not sel_campus or not sel_dept or not sel_unit: st.warning("請完整填寫單位資訊")
            elif not f_file: st.error("請上傳佐證")
            else:
                # 簡單寫入測試
                st.success("填報功能測試中... (資料正確抓取)")
                # 在此處貼上實際寫入邏輯
