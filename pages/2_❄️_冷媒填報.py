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

# 1. 樣式
st.markdown("""
<style>
    [data-testid="stFileUploaderDropzone"] {background-color: #D6EAF8; border: 2px dashed #2E86C1; padding: 20px;}
    .note-text {color: #566573; font-weight: bold; font-size: 0.9rem;}
</style>
""", unsafe_allow_html=True)

# 2. 身份驗證
if st.session_state.get("authentication_status") is not True:
    st.warning("🔒 請先至首頁 (Hello) 登入系統")
    st.stop()

# 3. 資料庫連線
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
    
    ws_units = sh_ref.worksheet("單位資訊")
    ws_buildings = sh_ref.worksheet("建築物清單")
    ws_types = sh_ref.worksheet("設備類型")
    ws_coef = sh_ref.worksheet("冷媒係數表")
    
    try: ws_records = sh_ref.worksheet("冷媒填報紀錄")
    except: 
        ws_records = sh_ref.add_worksheet(title="冷媒填報紀錄", rows="1000", cols="15")
        ws_records.append_row(["填報時間","填報人","填報人分機","校區","所屬單位","填報單位名稱","建築物名稱","辦公室編號","維修日期","設備類型","設備品牌型號","冷媒種類","冷媒填充量","備註","佐證資料"])

except Exception as e:
    st.error(f"❌ 冷媒資料庫連線失敗: {e}")
    st.stop()

# 4. 資料讀取 (V210: 改用欄位名稱篩選)
@st.cache_data(ttl=60)
def load_ref_data_v210():
    def get_clean_df(ws):
        data = ws.get_all_values()
        if len(data) > 1:
            # 建立 DataFrame 並將所有欄位轉為字串
            df = pd.DataFrame(data[1:], columns=data[0]).astype(str)
            
            # 清洗欄位名稱 (移除前後空白)
            df.columns = df.columns.str.strip()
            
            # 清洗內容 (移除前後空白)
            df = df.apply(lambda x: x.str.strip())
            
            return df
        return pd.DataFrame()
    
    return get_clean_df(ws_units), get_clean_df(ws_buildings), get_clean_df(ws_types), get_clean_df(ws_coef)

df_units, df_buildings, df_types, df_coef = load_ref_data_v210()

# 5. 介面
st.title("❄️ 冷媒填報專區")

tabs = st.tabs(["📝 新增填報", "📊 動態查詢看板"])

with tabs[0]:
    with st.form("ref_form", clear_on_submit=True):
        st.subheader("填報人基本資料區")
        c1, c2, c3 = st.columns(3)
        
        # --- V210 連動核心邏輯 (使用欄位名稱) ---
        
        # 1. 校區 (從 '校區' 欄位取值)
        campuses = []
        if not df_units.empty and '校區' in df_units.columns:
            campuses = sorted(df_units['校區'].unique())
            # 過濾空字串
            campuses = [x for x in campuses if x]
            
        sel_campus = c1.selectbox("校區", campuses, index=None, placeholder="請選擇...")
        
        # 2. 所屬單位 (篩選 '校區' == sel_campus，取 '所屬單位')
        depts = []
        if sel_campus and not df_units.empty:
            # 使用 query 語法或 boolean indexing (確保型態一致)
            mask = df_units['校區'] == str(sel_campus)
            depts = sorted(df_units[mask]['所屬單位'].unique())
            depts = [x for x in depts if x]
            
        sel_dept = c2.selectbox("所屬單位", depts, index=None, placeholder="請先選擇校區...")
        
        # 3. 填報單位名稱 (篩選 '校區' & '所屬單位'，取 '填報單位名稱')
        units = []
        if sel_dept and not df_units.empty:
            mask = (df_units['校區'] == str(sel_campus)) & (df_units['所屬單位'] == str(sel_dept))
            units = sorted(df_units[mask]['填報單位名稱'].unique())
            units = [x for x in units if x]
            
        sel_unit = c3.selectbox("填報單位名稱", units, index=None, placeholder="請先選擇所屬單位...")
        
        # ------------------------------------
        
        c4, c5 = st.columns(2)
        name = c4.text_input("填報人")
        ext = c5.text_input("填報人分機")
        
        st.markdown("---")
        st.subheader("詳細位置資訊區")
        c6, c7 = st.columns(2)
        
        # 建築物 (假設欄位名稱為 '校區' 和 '建築物名稱' - 請依實際Sheet調整)
        # 根據您的除錯圖，建築物清單應該也有類似結構
        builds = []
        if sel_campus and not df_buildings.empty:
            # 嘗試自動尋找對應欄位
            col_campus = df_buildings.columns[0] # 假設第1欄是校區
            col_build = df_buildings.columns[1]  # 假設第2欄是建築物
            
            mask = df_buildings[col_campus] == str(sel_campus)
            builds = sorted(df_buildings[mask][col_build].unique())
            builds = [x for x in builds if x]
            
        sel_build = c6.selectbox("建築物名稱", builds, index=None, placeholder="請先選擇上方校區...")
        office = c7.text_input("辦公室編號", placeholder="例如：404辦公室")
        
        st.markdown("---")
        st.subheader("設備修繕冷媒填充資訊區")
        c8, c9 = st.columns(2)
        r_date = c8.date_input("維修日期 (統一填寫發票日期)", datetime.today())
        
        e_types = sorted(df_types.iloc[:, 0].unique()) if not df_types.empty else []
        sel_etype = c9.selectbox("設備類型", e_types, index=None, placeholder="請選擇...")
        
        c10, c11 = st.columns(2)
        e_model = c10.text_input("設備品牌型號", placeholder="例如：國際 CS-100FL+CU-100FLC")
        
        # 冷媒種類 (B欄, index 1)
        r_types = []
        if not df_coef.empty and df_coef.shape[1] > 1:
            r_types = sorted(df_coef.iloc[:, 1].unique())
            r_types = [x for x in r_types if x]
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
            elif not reporter_name or not reporter_ext: st.warning("⚠️ 請填寫填報人與分機") # V210 補上變數
            elif not sel_build: st.warning("⚠️ 請選擇建築物")
            elif not sel_etype or not sel_rtype: st.warning("⚠️ 請選擇設備類型與冷媒種類")
            elif not f_file: st.error("⚠️ 請上傳佐證資料")
            else:
                try:
                    f_file.seek(0); f_ext = f_file.name.split('.')[-1]
                    # V132 邏輯應用於冷媒
                    clean_name = f"{sel_campus}_{sel_dept}_{sel_unit}_{r_date}_{sel_etype}_{sel_rtype}.{f_ext}"
                    
                    meta = {'name': clean_name, 'parents': [REF_FOLDER_ID]}
                    media = MediaIoBaseUpload(f_file, mimetype=f_file.type, resumable=True)
                    file = drive_service.files().create(body=meta, media_body=media, fields='webViewLink').execute()
                    link = file.get('webViewLink')
                    
                    # 寫入時間
                    current_time = get_taiwan_time().strftime("%Y-%m-%d %H:%M:%S")
                    
                    # V210: 修正變數名稱對應
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