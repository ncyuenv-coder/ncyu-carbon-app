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
st.set_page_config(page_title="冷媒填報(診斷中)", page_icon="🚑", layout="wide")

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
    .debug-box {
        background-color: #FADBD8; border: 2px solid #E74C3C; 
        padding: 15px; border-radius: 10px; margin-bottom: 20px;
        color: #C0392B; font-family: monospace;
    }
    .raw-data {
        background-color: #EAEDED; padding: 10px; font-size: 0.8rem;
    }
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
    st.error(f"❌ 資料庫連線失敗: {e}")
    st.stop()

# 4. 資料讀取 (V222: 極限原始讀取)
@st.cache_data(ttl=60)
def load_ref_data_v222():
    def get_raw_df(ws):
        # 不做任何清洗，直接抓所有資料
        data = ws.get_all_values()
        if len(data) > 1:
            # 為了怕標題重複，我們自己命名 Col_0, Col_1
            cols = [f"Col_{i}" for i in range(len(data[0]))]
            df = pd.DataFrame(data[1:], columns=cols)
            # 順便把原始標題存下來當參考
            df.attrs['original_headers'] = data[0]
            return df
        return pd.DataFrame()
    
    return get_raw_df(ws_units), get_raw_df(ws_buildings), get_raw_df(ws_types), get_raw_df(ws_coef)

df_units, df_buildings, df_types, df_coef = load_ref_data_v222()

# 5. 頁面內容
st.title("🚑 冷媒填報 (極限診斷模式)")

if st.button("🔄 刷新資料庫", type="secondary"):
    st.cache_data.clear()
    st.rerun()

# --- 極限診斷區 ---
with st.expander("🛠️ 點此查看 B 欄到底有沒有被讀進來", expanded=True):
    st.markdown("#### 1. 【單位資訊】資料表結構檢查")
    if not df_units.empty:
        st.write(f"資料表大小 (Rows, Cols): {df_units.shape}")
        st.write(f"原始標題列: {df_units.attrs.get('original_headers', 'Unknown')}")
        
        st.markdown("#### 2. 前 5 筆原始資料 (Raw Data)")
        st.dataframe(df_units.head())
        
        st.markdown("#### 3. B 欄 (Col_1) 唯一值預覽")
        if df_units.shape[1] > 1:
            unique_b = df_units['Col_1'].unique()
            st.write(f"找到 {len(unique_b)} 個填報單位，前 10 個如下：")
            st.write(unique_b[:10])
        else:
            st.error("❌ 慘！程式只讀到 1 個欄位，B 欄完全消失！")
    else:
        st.error("❌ 連 A 欄都沒讀到，資料表是空的！")

# --- 填報區 (測試用) ---
st.markdown("---")
st.markdown("### 🧪 測試區：不做連動，直接列出所有選項")

c1, c2 = st.columns(2)

# 1-1. 所屬單位 (Col_0)
if not df_units.empty:
    # 簡單清洗一下空白
    dept_list = sorted([str(x).strip() for x in df_units['Col_0'].unique() if str(x).strip()])
    sel_dept = c1.selectbox("A欄 (所屬單位)", dept_list, index=None)
else:
    c1.warning("無 A 欄資料")
    sel_dept = None

# 1-2. 填報單位名稱 (Col_1) - 這裡我不做連動，直接列出全部，看看有沒有東西
if not df_units.empty and df_units.shape[1] > 1:
    unit_list = sorted([str(x).strip() for x in df_units['Col_1'].unique() if str(x).strip()])
    c2.selectbox("B欄 (所有填報單位 - 不連動測試)", unit_list, index=None)
else:
    c2.error("無 B 欄資料")

# 如果 A 欄選了，我們試著手動篩選一次給你看
if sel_dept and not df_units.empty and df_units.shape[1] > 1:
    st.info(f"正在嘗試篩選 A欄 = '{sel_dept}' 的資料...")
    # 這裡用最笨的方法比對：字串包含
    mask = df_units['Col_0'].astype(str).str.contains(sel_dept, na=False)
    filtered_units = df_units[mask]['Col_1'].unique()
    
    if len(filtered_units) > 0:
        st.success(f"✅ 成功篩選到 {len(filtered_units)} 筆資料！")
        st.write(filtered_units)
    else:
        st.error(f"❌ 篩選結果為空！代表 A 欄的值 '{sel_dept}' 跟資料庫裡的值長得不一樣 (可能有隱形空白)。")