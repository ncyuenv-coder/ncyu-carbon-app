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
import os

# ==========================================
# 0. 系統設定
# ==========================================
st.set_page_config(page_title="冷媒填報 - 嘉義大學", page_icon="❄️", layout="wide")

def get_taiwan_time():
    return datetime.utcnow() + timedelta(hours=8)

# 取得使用者資訊
username = st.session_state.get("username")
name = st.session_state.get("name")

# ==========================================
# 1. CSS 樣式表
# ==========================================
st.markdown("""
<style>
    /* --- 全域設定 --- */
    :root {
        color-scheme: light;
        --orange-bg: #E67E22;     
        --orange-dark: #D35400;
        --text-main: #2C3E50;
        --text-sub: #566573;
        --morandi-red: #C0392B; 
        --kpi-gas: #52BE80;
        --kpi-diesel: #F4D03F;
        --kpi-total: #5DADE2;
        --kpi-co2: #AF7AC5;
        --morandi-blue: #34495E;
    }
    [data-testid="stAppViewContainer"] { background-color: #EAEDED; color: var(--text-main); }
    [data-testid="stHeader"] { background-color: rgba(0,0,0,0); }
    [data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 1px solid #BDC3C7; }
    div[data-baseweb="input"] > div, div[data-baseweb="base-input"] > input, textarea, input {
        background-color: #FFFFFF !important; border-color: #BDC3C7 !important; color: #000000 !important; font-size: 1.15rem !important;
    }
    div[data-baseweb="select"] > div { border-color: #BDC3C7 !important; background-color: #FFFFFF !important; }
    ul[data-baseweb="menu"] { background-color: #FFFFFF !important; }
    div.stButton > button, button[kind="primary"], [data-testid="stFormSubmitButton"] > button {
        background-color: var(--orange-bg) !important; 
        color: #FFFFFF !important; border: 2px solid var(--orange-dark) !important; border-radius: 12px !important;
        font-size: 1.3rem !important; font-weight: 800 !important; padding: 0.7rem 1.5rem !important;
        box-shadow: 0 4px 6px rgba(230, 126, 34, 0.3) !important; width: 100%; 
    }
    div.stButton > button p { color: #FFFFFF !important; } 
    div.stButton > button:hover, [data-testid="stFormSubmitButton"] > button:hover { 
        background-color: var(--orange-dark) !important; transform: translateY(-2px) !important; color: #FFFFFF !important;
    }
    button[data-baseweb="tab"] div p { font-size: 1.3rem !important; font-weight: 900 !important; color: var(--text-sub); }
    button[data-baseweb="tab"][aria-selected="true"] div p { color: #E67E22 !important; border-bottom: 3px solid #E67E22; }
    div[data-testid="stCheckbox"] label p { font-size: 1.2rem !important; color: #1F618D !important; font-weight: 900 !important; }
    .morandi-header {
        background-color: #EBF5FB; color: #2E4053; padding: 15px; border-radius: 8px;
        border-left: 8px solid #5499C7; font-size: 1.35rem; font-weight: 700;
        margin-top: 25px; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .privacy-box { background-color: #F8F9F9; border: 1px solid #BDC3C7; padding: 15px; border-radius: 10px; font-size: 0.95rem; color: #566573; margin-bottom: 10px; }
    .privacy-title { font-weight: bold; color: #2C3E50; margin-bottom: 5px; font-size: 1.1rem; }
    .correction-note { color: #566573; font-size: 0.95rem; font-weight: bold; margin-top: 5px; margin-bottom: 20px; }
    .horizontal-card {
        display: flex; border: 1px solid #BDC3C7; border-radius: 12px; overflow: hidden;
        margin-bottom: 25px; box-shadow: 0 4px 8px rgba(0,0,0,0.08); background-color: #FFFFFF; min-height: 280px;
    }
    .card-left {
        flex: 3; background-color: var(--morandi-blue); color: #FFFFFF;
        display: flex; flex-direction: column; justify-content: center; align-items: center;
        padding: 20px; text-align: center; border-right: 1px solid #2C3E50;
    }
    .dept-text { font-size: 1.6rem; font-weight: 700; margin-bottom: 8px; line-height: 1.4; }
    .unit-text { font-size: 1.3rem; font-weight: 500; opacity: 0.9; }
    .card-right { flex: 7; padding: 20px 30px; display: flex; flex-direction: column; justify-content: center; }
    .info-row { display: flex; align-items: flex-start; padding: 10px 0; font-size: 1.05rem; color: #566573; border-bottom: 1px dashed #F2F3F4; }
    .info-row:last-child { border-bottom: none; }
    .info-icon { margin-right: 12px; font-size: 1.2rem; width: 30px; text-align: center; }
    .info-label { font-weight: 700; margin-right: 10px; min-width: 160px; color: #2E4053; }
    .info-value { font-weight: 500; color: #17202A; flex: 1; line-height: 1.6; }
    .admin-kpi-card {
        background-color: #FFFFFF; border: 1px solid #BDC3C7; border-radius: 12px; overflow: hidden;
        box-shadow: 0 4px 10px rgba(0,0,0,0.1); height: 100%; text-align: center; margin-bottom: 20px;
    }
    .admin-kpi-header { padding: 10px; font-size: 1.2rem; font-weight: bold; color: #2C3E50; border-bottom: 1px solid rgba(0,0,0,0.1); }
    .admin-kpi-body { padding: 20px; }
    .admin-kpi-value { font-size: 2.8rem; font-weight: 900; color: #2C3E50; margin-bottom: 5px; }
    .admin-kpi-unit { font-size: 1rem; color: #7F8C8D; font-weight: normal; margin-left: 5px; }
    .dashboard-main-title {
        font-size: 1.8rem; font-weight: 900; text-align: center; color: #2C3E50; margin-bottom: 20px;
        background-color: #F8F9F9; padding: 10px; border-radius: 10px; border: 1px solid #BDC3C7;
    }
    .stRadio div[role="radiogroup"] label {
        background-color: #D6EAF8 !important; border: 1px solid #AED6F1 !important;
        border-radius: 8px !important; padding: 8px 15px !important; margin-right: 10px !important;
    }
    .stRadio div[role="radiogroup"] label p { font-size: 1.0rem !important; font-weight: 800 !important; color: #154360 !important; }
    [data-testid="stFileUploaderDropzone"] { background-color: #D6EAF8 !important; border: 2px dashed #2E86C1 !important; padding: 20px; border-radius: 12px; }
    [data-testid="stFileUploaderDropzone"] div, span, small { color: #154360 !important; font-weight: bold !important; }
</style>
""", unsafe_allow_html=True)

# 2. 身份驗證
if st.session_state.get("authentication_status") is not True:
    st.warning("🔒 請先至首頁 (Hello) 登入系統")
    st.stop()

# 3. 資料庫連線 (保持 Google Sheets 連線以供寫入)
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
    try: ws_records = sh_ref.worksheet("冷媒填報紀錄")
    except: 
        ws_records = sh_ref.add_worksheet(title="冷媒填報紀錄", rows="1000", cols="15")
        ws_records.append_row(["填報時間","填報人","填報人分機","校區","所屬單位","填報單位名稱","建築物名稱","辦公室編號","維修日期","設備類型","設備品牌型號","冷媒種類","冷媒填充量","備註","佐證資料"])
except Exception as e:
    st.error(f"❌ 資料庫連線失敗: {e}")
    st.stop()

# 4. 資料讀取 (V280.0 - 內建資料庫 + 雲端更新)

# 內建預設資料 (Fallback)
DEFAULT_UNITS = {'教務處': ['教務長室/副教務長室/專門委員室', '註冊與課務組', '教學發展組', '招生與出版組', '綜合行政組', '通識教育中心', '民雄教務'], '學生事務處': ['學務長室/副學務長室', '住宿服務組', '生活輔導組', '課外活動組', '學生輔導中心', '學生職涯發展中心', '衛生保健組', '原住民族學生資源中心', '特殊教育學生資源中心', '民雄學務'], '總務處': ['總務長室/副總務長室/簡任秘書室', '事務組', '出納組', '文書組', '資產經營管理組', '營繕組', '民雄總務', '新民聯辦', '駐衛警察隊'], '研究發展處': ['研發長室/副研發長室', '綜合企劃組', '學術發展組', '校務研究組'], '產學營運及推廣處': ['產學營運及推廣處長室', '行政管理組', '產學創育推廣中心'], '國際事務處': ['國際事務長室', '境外生事務組', '國際合作組'], '圖書資訊處': ['圖資長室', '圖資管理組', '資訊網路組', '諮詢服務組', '系統資訊組', '民雄圖書資訊', '新民分館', '民雄分館'], '校長室': ['校長室'], '行政副校長室': ['行政副校長室'], '學術副校長室': ['學術副校長室'], '國際副校長室': ['國際副校長室'], '秘書室': ['綜合業務組', '公共關係組', '校友服務組'], '體育室': ['蘭潭場館', '民雄場館', '林森場館', '新民場館'], '主計室': ['主計室'], '人事室': ['人事室'], '環境保護及安全管理中心': ['環境保護及安全管理中心'], '師資培育中心': ['師資培育中心主任室', '教育課程組', '實習輔導組', '綜合行政組'], '語言中心': ['主任室', '蘭潭語言中心', '民雄語言中心', '新民語言中心'], '理工學院': ['理工學院辦公室', '應用數學系', '電子物理學系', '應用化學系', '資訊工程學系', '生物機電工程學系', '土木與水資源工程學系', '水工與材料試驗場', '電機工程學系', '機械與能源工程學系'], '農學院': ['農學院辦公室', '農藝學系', '園藝學系', '森林暨自然資源學系', '木質材料與設計學系', '動物科學系', '農業經濟學系', '生物農業科技學系', '景觀學系', '植物醫學系', '農場管理進修學士學位學程'], '生命科學院': ['生命科學院辦公室', '食品科學系', '水生生物科學系', '生物資源學系', '生化科技學系', '微生物免疫與生物藥學系'], '管理學院': ['管理學院辦公室', '企業管理學系', '應用經濟學系', '生物事業管理學系', '資訊管理學系', '財務金融學系', '行銷與觀光管理學系', '全英文授課觀光暨管理學士學位學程'], '獸醫學院': ['獸醫學院辦公室', '獸醫學系', '雲嘉南動物疾病診斷中心', '動物醫院'], '師範學院': ['師範學院辦公室', '教育學系', '輔導與諮商學系', '體育與健康休閒學系', '特殊教育學系', '幼兒教育學系', '教育行政與政策發展研究所', '數理教育研究所'], '人文藝術學院': ['人文藝術學院辦公室', '中國文學系', '外國語言學系', '應用歷史學系', '視覺藝術學系', '音樂學系']}
DEFAULT_BUILDINGS = {'蘭潭校區': ['A01行政中心', 'A02森林館', 'A03動物科學館', 'A04農園館', 'A05工程館', 'A06食品科學館', 'A07嘉禾館', 'A08瑞穗館', 'A09游泳池', 'A10機械與能源工程學系創新育成大樓', 'A11木材利用工廠', 'A12動物試驗場', 'A13司令台', 'A14學生活動中心', 'A15電物一館', 'A16理工大樓', 'A17應化一館', 'A18A應化二館', 'A18B電物二館', 'A19農藝場管理室', 'A20國際交流學園', 'A21水工與材料試驗場', 'A22食品加工廠', 'A23機電館', 'A24生物資源館', 'A25生命科學館', 'A26農業科學館', 'A27植物醫學系館', 'A28水生生物科學館', 'A29園藝場管理室', 'A30園藝技藝中心', 'A31圖書資訊館', 'A32綜合教學大樓', 'A33生物農業科技二館', 'A34嘉大植物園', 'A35生技健康館', 'A36景觀學系大樓', 'A37森林生物多樣性館', 'A38動物產品研發推廣中心', 'A39學生活動廣場', 'A40焚化爐設備車倉庫', 'A41生物機械產業實驗室', 'A44有機蔬菜溫室', 'A45蝴蝶蘭溫室', 'A46魚類保育研究中心', 'A71員工單身宿舍', 'A72學苑餐廳', 'A73學一舍', 'A74學二舍', 'A75學三舍', 'A76學五舍', 'A77學六舍', 'A78農產品展售中心', 'A79綠建築', 'A80嘉大昆蟲館', 'A81蘭潭招待所', 'A82警衛室', '回收場'], '民雄校區': ['B01創意樓', 'B02大學館', 'B03教育館', 'B04新藝樓', 'B06警衛室', 'B07鍋爐間', 'B08司令台', 'B09加氯室', 'B10游泳池', 'B12工友室', 'BA行政大樓', 'BB初等教育館', 'BC圖書館', 'BD樂育堂', 'BE學人單身宿舍', 'BF綠園二舍', 'BG餐廳', 'BH綠園一舍', 'BI科學館', 'BJ人文館', 'BK音樂館', 'BL藝術館', 'BM文薈廳', 'BN社團教室'], '林森校區': ['C01警衛室', 'C02司令台', 'CA第一棟大樓', 'CB進修部大樓', 'CD國民輔導大樓', 'CE第二棟大樓', 'CF實輔室', 'CG圖書館', 'CH視聽教室', 'CI明德齋', 'CK餐廳', 'CL青雲齋', 'CN樂育堂', 'CP空大學習指導中心'], '新民校區': ['D01管理學院大樓A棟', 'D02管理學院大樓B棟', 'D03明德樓', 'D04獸醫館(獸醫學系、動物醫院、雲嘉南動物疾病診斷中心)', 'D05游泳池', 'D06溫室', 'D07司令台', 'D08警衛室', 'D09賢德樓'], '社口林場': ['E01林場實習館'], '林森校區-民國路': ['F01民國路進德樓']}
DEFAULT_TYPES = ['冰水主機', '冰箱', '冷凍櫃', '冷氣', '冷藏櫃', '飲水機']
DEFAULT_GWP = {'HFC-1234yf (R-1234yf)': 0.0, 'HFC-125 (R-125)': 3170.0, 'HFC-134a (R-134a)': 1300.0, 'HFC-143a (R-143a)': 4800.0, 'HFC-245fa (R-245fa)': 858.0, 'R404a': 3942.8, 'R407c': 1624.21, 'R-407D': 1487.05, 'R408a': 2429.9, 'R410a': 1923.5, 'R-507A': 3985.0, 'R-508A': 11607.0, 'R-508B': 11698.0, 'HFC-23 (R-23)': 12400.0, 'HFC-32 (R-32)': 677.0, 'R-411A': 0.0, 'R-402A': 0.0}

def load_static_data(source='local'):
    """
    讀取靜態資料。
    source='local': 直接回傳內建的字典 (速度最快，永不報錯)。
    source='cloud': 連線 Google Sheets 下載最新資料並更新。
    """
    if source == 'local':
        return DEFAULT_UNITS, DEFAULT_BUILDINGS, DEFAULT_TYPES, sorted(list(DEFAULT_GWP.keys())), DEFAULT_GWP
    else:
        # Cloud update
        try:
            ws_units = sh_ref.worksheet("單位資訊")
            ws_buildings = sh_ref.worksheet("建築物清單")
            ws_types = sh_ref.worksheet("設備類型")
            ws_coef = sh_ref.worksheet("冷媒係數表")
            
            df_units = pd.DataFrame(ws_units.get_all_records()).astype(str)
            df_build = pd.DataFrame(ws_buildings.get_all_records()).astype(str)
            df_types = pd.DataFrame(ws_types.get_all_records()).astype(str)
            df_coef = pd.DataFrame(ws_coef.get_all_records())
            
            # Process Units
            unit_dict = {}
            for _, row in df_units.iterrows():
                d = str(row.iloc[0]).strip()
                u = str(row.iloc[1]).strip()
                if d and u:
                    if d not in unit_dict: unit_dict[d] = []
                    if u not in unit_dict[d]: unit_dict[d].append(u)
            
            # Process Buildings
            build_dict = {}
            for _, row in df_build.iterrows():
                c = str(row.iloc[0]).strip()
                b = str(row.iloc[1]).strip()
                if c and b:
                    if c not in build_dict: build_dict[c] = []
                    if b not in build_dict[c]: build_dict[c].append(b)
            
            # Process Types
            e_types = sorted(df_types.iloc[:, 0].dropna().unique().tolist()) if not df_types.empty else []
            
            # Process Coef
            gwp_map = {}
            r_types = []
            if not df_coef.empty:
                for _, row in df_coef.iterrows():
                    r_name_full = str(row.iloc[1]).strip().replace('\u3000', ' ').replace('\xa0', ' ')
                    try: gwp = float(str(row.iloc[2]).replace(',', ''))
                    except: gwp = 0.0
                    if r_name_full:
                        r_types.append(r_name_full)
                        gwp_map[r_name_full] = gwp
            
            return unit_dict, build_dict, e_types, sorted(list(set(r_types))), gwp_map
            
        except Exception as e:
            st.error(f"雲端更新失敗: {e}")
            return DEFAULT_UNITS, DEFAULT_BUILDINGS, DEFAULT_TYPES, sorted(list(DEFAULT_GWP.keys())), DEFAULT_GWP

@st.cache_data(ttl=60)
def load_records_data():
    """動態填報紀錄 (嚴格只讀取 Google Sheets)"""
    try:
        data = ws_records.get_all_values()
        if len(data) > 1:
            raw_headers = data[0]
            col_mapping = {}
            for h in raw_headers:
                clean_h = str(h).strip()
                if "填充量" in clean_h or "重量" in clean_h: col_mapping[h] = "冷媒填充量"
                elif "種類" in clean_h or "品項" in clean_h: col_mapping[h] = "冷媒種類"
                elif "日期" in clean_h or "維修" in clean_h: col_mapping[h] = "維修日期"
                else: col_mapping[h] = clean_h
            
            df = pd.DataFrame(data[1:], columns=raw_headers)
            df.rename(columns=col_mapping, inplace=True)
            return df
        else:
            return pd.DataFrame(columns=["填報時間","填報人","填報人分機","校區","所屬單位","填報單位名稱","建築物名稱","辦公室編號","維修日期","設備類型","設備品牌型號","冷媒種類","冷媒填充量","備註","佐證資料"])
    except Exception as e:
        st.error(f"⚠️ 無法讀取雲端填報紀錄 (Google Sheet 連線錯誤): {e}。請檢查網路連線。")
        return pd.DataFrame()

# 初始化 (Session State)
if 'static_data_loaded' not in st.session_state:
    # 預設使用內建資料
    st.session_state['unit_dict'], st.session_state['build_dict'], st.session_state['e_types'], st.session_state['r_types'], st.session_state['gwp_map'] = load_static_data('local')
    st.session_state['static_data_loaded'] = True

unit_dict = st.session_state['unit_dict']
build_dict = st.session_state['build_dict']
e_types = st.session_state['e_types']
r_types = st.session_state['r_types']
gwp_map = st.session_state['gwp_map']

df_records = load_records_data()

# 清空欄位用
def reset_input_states():
    keys_to_reset = ['u_dept', 'u_unit', 'u_name', 'u_ext', 'u_campus', 'u_build', 'u_office', 'u_date', 'u_etype', 'u_model', 'u_rtype', 'u_amt', 'u_note', 'u_file', 'u_agree']
    for k in keys_to_reset:
        if k in st.session_state:
            del st.session_state[k]

# ==========================================
# 5. 功能模組：一般使用者介面
# ==========================================
def render_user_interface():
    st.markdown("### ❄️ 冷媒填報專區")
    tabs = st.tabs(["📝 新增填報", "📋 申報動態查詢"])

    # --- Tab 1: 新增填報 ---
    with tabs[0]:
        st.markdown('<div class="morandi-header">填報單位基本資訊區</div>', unsafe_allow_html=True)
        
        c1, c2 = st.columns(2)
        unit_depts = sorted(unit_dict.keys())
        sel_dept = c1.selectbox("所屬單位", unit_depts, index=None, placeholder="請選擇單位...", key="u_dept")
        unit_names = sorted(unit_dict.get(sel_dept, [])) if sel_dept else []
        sel_unit_name = c2.selectbox("填報單位名稱", unit_names, index=None, placeholder="請先選擇所屬單位...", key="u_unit")
        
        c3, c4 = st.columns(2)
        name = c3.text_input("填報人", key="u_name")
        ext = c4.text_input("填報人分機", key="u_ext")
        
        st.markdown('<div class="morandi-header">冷媒設備所在位置資訊區</div>', unsafe_allow_html=True)
        loc_campuses = sorted(build_dict.keys())
        sel_loc_campus = st.selectbox("填報單位所在校區", loc_campuses, index=None, placeholder="請選擇校區...", key="u_campus")
        c6, c7 = st.columns(2)
        buildings = sorted(build_dict.get(sel_loc_campus, [])) if sel_loc_campus else []
        sel_build = c6.selectbox("建築物名稱", buildings, index=None, placeholder="請先選擇校區...", key="u_build")
        office = c7.text_input("辦公室編號", placeholder="例如：202辦公室、306研究室", key="u_office")
        
        st.markdown('<div class="morandi-header">冷媒設備填充資訊區</div>', unsafe_allow_html=True)
        c8, c9 = st.columns(2)
        r_date = c8.date_input("維修日期 (統一填寫發票日期)", datetime.today(), key="u_date")
        sel_etype = c9.selectbox("設備類型", e_types, index=None, placeholder="請選擇...", key="u_etype")
        
        c10, c11 = st.columns(2)
        e_model = c10.text_input("設備品牌型號", placeholder="例如：國際 CS-100FL+CU-100FLC", key="u_model")
        sel_rtype = c11.selectbox("冷媒種類", r_types, index=None, placeholder="請選擇...", key="u_rtype")
        
        amount = st.number_input("冷媒填充量 (公斤)", min_value=0.0, step=0.1, format="%.2f", key="u_amt")
        st.markdown("請上傳冷媒填充單據佐證資料")
        f_file = st.file_uploader("上傳佐證 (必填)", type=['pdf', 'jpg', 'png'], label_visibility="collapsed", key="u_file")
        
        st.markdown("---")
        note = st.text_input("備註內容", placeholder="備註 (選填)", key="u_note")
        st.markdown('<div class="correction-note">如有資料誤繕情形，請重新登錄1次資訊，並於備註欄填寫：「前筆資料誤繕，請刪除。」，管理單位將協助刪除誤打資訊</div>', unsafe_allow_html=True)
        
        st.markdown("""<div class="privacy-box"><div class="privacy-title">📜 個人資料蒐集、處理及利用告知聲明</div>1. 蒐集機關：國立嘉義大學。<br>2. 蒐集目的：進行本校冷媒設備之冷媒填充紀錄管理、校園溫室氣體（碳）盤查統計、稽核佐證資料蒐集及後續能源使用分析。<br>3. 個資類別：填報人姓名。<br>4. 利用期間：姓名保留至填報年度後第二年1月1日，期滿即進行「去識別化」刪除，其餘數據永久保存。<br>5. 利用對象：本校教師、行政人員及碳盤查查驗人員。<br>6. 您有權依個資法請求查詢、更正或刪除您的個資。如不提供，將無法完成填報。</div>""", unsafe_allow_html=True)
        agree = st.checkbox("我已閱讀並同意個資聲明，且確認所填資料無誤。", key="u_agree")
        
        if st.button("🚀 確認送出", type="primary", use_container_width=True):
            if not agree: st.error("❌ 請勾選同意聲明")
            elif not sel_dept or not sel_unit_name: st.warning("⚠️ 請完整選擇單位資訊")
            elif not name or not ext: st.warning("⚠️ 請填寫填報人與分機")
            elif not sel_loc_campus or not sel_build: st.warning("⚠️ 請完整選擇位置資訊")
            elif not sel_etype or not sel_rtype: st.warning("⚠️ 請選擇設備類型與冷媒種類")
            elif not f_file: st.error("⚠️ 請上傳佐證資料")
            else:
                try:
                    f_file.seek(0); f_ext = f_file.name.split('.')[-1]
                    clean_name = f"{sel_loc_campus}_{sel_dept}_{sel_unit_name}_{r_date}_{sel_etype}_{sel_rtype}.{f_ext}"
                    meta = {'name': clean_name, 'parents': [REF_FOLDER_ID]}
                    media = MediaIoBaseUpload(f_file, mimetype=f_file.type, resumable=True)
                    file = drive_service.files().create(body=meta, media_body=media, fields='webViewLink').execute()
                    
                    row_data = [get_taiwan_time().strftime("%Y-%m-%d %H:%M:%S"), name, ext, sel_loc_campus, sel_dept, sel_unit_name, sel_build, office, str(r_date), sel_etype, e_model, sel_rtype, amount, note, file.get('webViewLink')]
                    ws_records.append_row(row_data)
                    
                    st.success("✅ 冷媒填報成功！欄位已清空，可繼續填寫下一筆。")
                    st.balloons()
                    
                    # 清空欄位並重新執行
                    reset_input_states()
                    st.cache_data.clear()
                    time.sleep(1)
                    st.rerun()
                    
                except Exception as e: st.error(f"上傳或寫入失敗: {e}")

    # --- Tab 2: 申報動態查詢 ---
    with tabs[1]:
        st.markdown('<div class="morandi-header">📋 申報動態查詢</div>', unsafe_allow_html=True)
        col_r1, col_r2 = st.columns([4, 1])
        with col_r2:
            if st.button("🔄 刷新數據", use_container_width=True, key="refresh_tab2"):
                st.cache_data.clear()
                st.rerun()

        if df_records.empty:
            st.info("目前尚無填報紀錄。")
        else:
            df_records['冷媒填充量'] = pd.to_numeric(df_records['冷媒填充量'], errors='coerce').fillna(0)
            df_records['維修日期'] = pd.to_datetime(df_records['維修日期'], errors='coerce')
            
            # 使用全名對應 GWP
            df_records['排放量(kgCO2e)'] = df_records.apply(lambda r: r['冷媒填充量'] * gwp_map.get(str(r['冷媒種類']).strip(), 0), axis=1)

            st.markdown("##### 🔍 查詢條件設定")
            c_f1, c_f2 = st.columns(2)
            sel_q_dept = c_f1.selectbox("所屬單位 (必選)", sorted(df_records['所屬單位'].dropna().unique()), index=None)
            sel_q_unit = c_f2.selectbox("填報單位名稱 (必選)", sorted(df_records[df_records['所屬單位']==sel_q_dept]['填報單位名稱'].dropna().unique()) if sel_q_dept else [], index=None)
            
            c_f3, c_f4 = st.columns(2)
            q_start_date = c_f3.date_input("查詢起始日期", value=date(datetime.now().year, 1, 1))
            q_end_date = c_f4.date_input("查詢結束日期", value=datetime.now().date())

            if sel_q_dept and sel_q_unit and q_start_date and q_end_date:
                mask = (df_records['所屬單位']==sel_q_dept) & (df_records['填報單位名稱']==sel_q_unit) & (df_records['維修日期']>=pd.Timestamp(q_start_date)) & (df_records['維修日期']<=pd.Timestamp(q_end_date))
                df_view = df_records[mask]
                
                if not df_view.empty:
                    left_html = f'<div class="dept-text">{sel_q_dept}</div>' if sel_q_dept==sel_q_unit else f'<div class="dept-text">{sel_q_dept}</div><div class="unit-text">{sel_q_unit}</div>'
                    campus_str = ", ".join(sorted(df_view['校區'].unique()))
                    build_str = ", ".join(sorted(df_view['建築物名稱'].unique())[:3])
                    
                    fill_info_list = []
                    for _, row in df_view.iterrows():
                        fill_info_list.append(f"<div>• {row['設備類型']}-{row['冷媒種類']}：{row['冷媒填充量']:.2f} kg</div>")
                    fill_info_str = "".join(fill_info_list)

                    total_kg = df_view['冷媒填充量'].sum()
                    type_sums = df_view.groupby('冷媒種類')['冷媒填充量'].sum().reset_index()
                    weight_str = f"<div style='font-weight: 900; margin-bottom: 5px; font-size: 1.05rem;'>總計：{total_kg:.2f} kg</div>"
                    for _, row in type_sums.iterrows():
                        weight_str += f"<div>• {row['冷媒種類']}：{row['冷媒填充量']:.2f} kg</div>"

                    total_emission = df_view['排放量(kgCO2e)'].sum()

                    st.markdown("---")
                    st.markdown(f"""
                    <div class="horizontal-card">
                        <div class="card-left">{left_html}</div>
                        <div class="card-right">
                            <div class="info-row"><span class="info-icon">📅</span><span class="info-label">查詢區間</span><span class="info-value">{q_start_date} ~ {q_end_date}</span></div>
                            <div class="info-row"><span class="info-icon">🏫</span><span class="info-label">所在校區</span><span class="info-value">{campus_str}</span></div>
                            <div class="info-row"><span class="info-icon">🏢</span><span class="info-label">建築物</span><span class="info-value">{build_str}</span></div>
                            <div class="info-row"><span class="info-icon">❄️</span><span class="info-label">冷媒填充資訊</span><span class="info-value">{fill_info_str}</span></div>
                            <div class="info-row"><span class="info-icon">⚖️</span><span class="info-label">重量統計</span><span class="info-value">{weight_str}</span></div>
                            <div class="info-row"><span class="info-icon">🌍</span><span class="info-label">碳排放量</span><span class="info-value" style="color:#C0392B;font-size:1.8rem;font-weight:900;">{total_emission:,.2f} <span style="font-size:1rem;">kgCO2e</span></span></div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    st.markdown("<br>", unsafe_allow_html=True)
                    st.subheader("📋 單位申報明細")
                    st.dataframe(df_view[["維修日期", "建築物名稱", "設備類型", "冷媒種類", "冷媒填充量", "排放量(kgCO2e)", "佐證資料"]], use_container_width=True)
                else: st.warning("查無資料")
            else: st.info("請選擇單位進行查詢")

# ==========================================
# 6. 功能模組：管理員後台
# ==========================================
def render_admin_dashboard():
    st.markdown("### 👑 冷媒管理後台")
    
    admin_tabs = st.tabs(["📊 全校冷媒填充儀表板", "📝 申報資料異動"])

    # V273: 按鈕放置於 Tab 2 內
    with admin_tabs[1]:
        st.subheader("📝 申報資料異動與下載")
        
        if st.button("🔄 更新背景資料庫 (從 Google Sheet 同步)", key="btn_update_db"):
            with st.spinner("正在從雲端下載最新資料..."):
                st.session_state['unit_dict'], st.session_state['build_dict'], st.session_state['e_types'], st.session_state['r_types'], st.session_state['gwp_map'] = load_static_data('cloud')
            st.success("✅ 資料庫已更新！")

        # 資料預處理
        df_clean = df_records.copy()
        if not df_clean.empty:
            campus_map = {"林森校區-民國路": "林森校區", "社口林場": "蘭潭校區"}
            df_clean['校區'] = df_clean['校區'].replace(campus_map)
            
            df_clean['冷媒顯示名稱'] = df_clean['冷媒種類']

            df_clean['冷媒填充量'] = pd.to_numeric(df_clean['冷媒填充量'], errors='coerce').fillna(0)
            df_clean['維修日期'] = pd.to_datetime(df_clean['維修日期'], errors='coerce')
            df_clean['年份'] = df_clean['維修日期'].dt.year.fillna(datetime.now().year).astype(int)
            df_clean['月份'] = df_clean['維修日期'].dt.month.fillna(0).astype(int)
            
            df_clean['排放量(kgCO2e)'] = df_clean.apply(lambda r: r['冷媒填充量'] * gwp_map.get(str(r['冷媒種類']).strip(), 0), axis=1)
            df_clean['排放量(公噸)'] = df_clean['排放量(kgCO2e)'] / 1000.0

            df_clean['維修日期'] = df_clean['維修日期'].dt.date
            edited = st.data_editor(
                df_clean, 
                column_config={
                    "佐證資料": st.column_config.LinkColumn("佐證"),
                    "冷媒填充量": st.column_config.NumberColumn("填充量", min_value=0.0, step=0.1),
                    "維修日期": st.column_config.DateColumn("日期", format="YYYY-MM-DD")
                },
                num_rows="dynamic",
                use_container_width=True,
                key="ref_editor"
            )
            
            csv_data = edited.to_csv(index=False).encode('utf-8-sig')
            st.download_button(label="📥 下載完整填報紀錄 (CSV)", data=csv_data, file_name=f"冷媒填報紀錄_full.csv", mime="text/csv", key="dl_ref_csv")
            
            if st.button("💾 儲存變更", type="primary"):
                try:
                    df_final = edited.copy()
                    
                    cols_to_remove = ['年份', '月份', '排放量(kgCO2e)', '排放量(公噸)', '冷媒顯示名稱']
                    for c in cols_to_remove:
                        if c in df_final.columns: del df_final[c]
                    
                    df_final['維修日期'] = df_final['維修日期'].astype(str)
                    
                    cols_to_write = df_records.columns.tolist()
                    cols_to_write = [c for c in cols_to_write if c in df_final.columns]
                    df_final = df_final[cols_to_write]
                    
                    ws_records.clear()
                    ws_records.update([df_final.columns.tolist()] + df_final.astype(str).values.tolist())
                    st.success("✅ 資料更新成功！")
                    st.cache_data.clear()
                    time.sleep(1)
                    st.rerun()
                except Exception as e:
                    st.error(f"更新失敗: {e}")
        else:
            st.info("無資料可編輯")

    # --- Admin Tab 1: 儀表板 ---
    MORANDI_PALETTE = ['#88B04B', '#92A8D1', '#F7CAC9', '#B565A7', '#009B77', '#DD4124', '#D65076', '#45B8AC', '#EFC050', '#5B5EA6', '#9B2335', '#DFCFBE']
    
    with admin_tabs[0]:
        df_clean_dash = df_records.copy()
        if not df_clean_dash.empty:
            campus_map = {"林森校區-民國路": "林森校區", "社口林場": "蘭潭校區"}
            df_clean_dash['校區'] = df_clean_dash['校區'].replace(campus_map)
            df_clean_dash['冷媒顯示名稱'] = df_clean_dash['冷媒種類']
            df_clean_dash['冷媒填充量'] = pd.to_numeric(df_clean_dash['冷媒填充量'], errors='coerce').fillna(0)
            df_clean_dash['維修日期'] = pd.to_datetime(df_clean_dash['維修日期'], errors='coerce')
            df_clean_dash['年份'] = df_clean_dash['維修日期'].dt.year.fillna(datetime.now().year).astype(int)
            df_clean_dash['月份'] = df_clean_dash['維修日期'].dt.month.fillna(0).astype(int)
            df_clean_dash['排放量(kgCO2e)'] = df_clean_dash.apply(lambda r: r['冷媒填充量'] * gwp_map.get(str(r['冷媒種類']).strip(), 0), axis=1)
            df_clean_dash['排放量(公噸)'] = df_clean_dash['排放量(kgCO2e)'] / 1000.0

        all_years = sorted(df_clean_dash['年份'].unique(), reverse=True) if not df_clean_dash.empty else [datetime.now().year]
        c_year, _ = st.columns([1, 3])
        sel_year = c_year.selectbox("📅 選擇統計年份", all_years)
        
        df_year = df_clean_dash[df_clean_dash['年份'] == sel_year] if not df_clean_dash.empty else pd.DataFrame()
        
        if not df_year.empty:
            st.markdown(f"<div class='dashboard-main-title'>{sel_year}年度 冷媒填充與碳排統計</div>", unsafe_allow_html=True)
            
            # KPI
            total_kg = df_year['冷媒填充量'].sum()
            total_co2_t = df_year['排放量(公噸)'].sum()
            count = len(df_year)
            
            k1, k2, k3 = st.columns(3)
            k1.markdown(f"""<div class="admin-kpi-card"><div class="admin-kpi-header" style="background-color: #A9CCE3;">📄 年度申報筆數</div><div class="admin-kpi-body"><div class="admin-kpi-value">{count}</div></div></div>""", unsafe_allow_html=True)
            k2.markdown(f"""<div class="admin-kpi-card"><div class="admin-kpi-header" style="background-color: #F5CBA7;">⚖️ 年度總填充量</div><div class="admin-kpi-body"><div class="admin-kpi-value">{total_kg:,.2f}<span class="admin-kpi-unit">公斤</span></div></div></div>""", unsafe_allow_html=True)
            k3.markdown(f"""<div class="admin-kpi-card"><div class="admin-kpi-header" style="background-color: #E6B0AA;">☁️ 年度總碳排放量</div><div class="admin-kpi-body"><div class="admin-kpi-value">{total_co2_t:,.4f}<span class="admin-kpi-unit">公噸CO<sub>2</sub>e</span></div></div></div>""", unsafe_allow_html=True)
            
            st.markdown("---")
            
            # Chart 1: 橫向標籤 V275.1 fix
            st.subheader("📈 年度冷媒填充概況")
            campus_opts = ["全校", "蘭潭校區", "民雄校區", "新民校區", "林森校區"]
            f_campus_1 = st.radio("填充概況校區選擇", campus_opts, horizontal=True, key="radio_c1", label_visibility="collapsed")
            
            df_c1 = df_year.copy()
            if f_campus_1 != "全校":
                df_c1 = df_c1[df_c1['校區'] == f_campus_1]
            
            if not df_c1.empty:
                c1_group = df_c1.groupby(['冷媒顯示名稱', '設備類型'])['冷媒填充量'].sum().reset_index()
                fig1 = px.bar(c1_group, x='冷媒顯示名稱', y='冷媒填充量', color='設備類型', 
                              text_auto='.1f', color_discrete_sequence=MORANDI_PALETTE)
                # V275.1: Force horizontal text
                fig1.update_layout(yaxis_title="冷媒填充量(公斤)", xaxis_title="冷媒種類", font=dict(size=18), showlegend=True)
                fig1.update_xaxes(tickfont=dict(size=16, color='#000000'))
                fig1.update_yaxes(tickfont=dict(size=16, color='#000000'))
                fig1.update_traces(width=0.5, textfont_size=14, textposition='inside', insidetextorientation='horizontal')
                st.plotly_chart(fig1, use_container_width=True)
            else:
                st.info("無資料")

            st.markdown("---")
            
            # Chart 2
            st.subheader("🏆 年度前十大填充單位")
            top_units = df_year.groupby('填報單位名稱')['冷媒填充量'].sum().nlargest(10).index.tolist()
            df_top10 = df_year[df_year['填報單位名稱'].isin(top_units)].copy()
            
            if not df_top10.empty:
                c2_group = df_top10.groupby(['填報單位名稱', '冷媒顯示名稱'])['冷媒填充量'].sum().reset_index()
                fig2 = px.bar(c2_group, x='填報單位名稱', y='冷媒填充量', color='冷媒顯示名稱',
                              text_auto='.1f', color_discrete_sequence=MORANDI_PALETTE)
                fig2.update_layout(xaxis={'categoryorder':'total descending'}, yaxis_title="冷媒填充量(公斤)", font=dict(size=18))
                fig2.update_xaxes(tickfont=dict(size=16, color='#000000'))
                fig2.update_yaxes(tickfont=dict(size=16, color='#000000'))
                fig2.update_traces(width=0.5, textfont_size=14, textposition='inside', insidetextorientation='horizontal')
                st.plotly_chart(fig2, use_container_width=True)
            else:
                st.info("無資料")
            
            st.markdown("---")
            
            # Chart 3
            st.subheader("🍩 冷媒填充資訊分析")
            f_campus_3 = st.radio("資訊分析校區選擇", campus_opts, horizontal=True, key="radio_c3", label_visibility="collapsed")
            df_c3 = df_year.copy()
            if f_campus_3 != "全校":
                df_c3 = df_c3[df_c3['校區'] == f_campus_3]
            
            if not df_c3.empty:
                st.markdown("##### 1. 冷媒種類填充量佔比")
                type_kg = df_c3.groupby('冷媒顯示名稱')['冷媒填充量'].sum().reset_index()
                fig3a = px.pie(type_kg, values='冷媒填充量', names='冷媒顯示名稱', hole=0.4, 
                               color_discrete_sequence=MORANDI_PALETTE)
                fig3a.update_layout(font=dict(size=18), legend=dict(font=dict(size=16)), uniformtext_minsize=16, uniformtext_mode='show')
                # V275: textposition='auto' + pull
                fig3a.update_traces(textinfo='label+percent', textfont_size=18, textposition='auto', pull=[0.1 if v < type_kg['冷媒填充量'].sum()*0.05 else 0 for v in type_kg['冷媒填充量']],
                                    hovertemplate='<b>%{label}</b><br>填充量: %{value:.1f} kg<br>佔比: %{percent:.1%}<extra></extra>')
                st.plotly_chart(fig3a, use_container_width=True)
                
                st.markdown("<br>", unsafe_allow_html=True)
                
                st.markdown("##### 2. 冷媒填充設備類型統計")
                c3_l, c3_r = st.columns(2)
                
                eq_count = df_c3.groupby('設備類型')['冷媒填充量'].count().reset_index(name='count')
                fig3b_l = px.pie(eq_count, values='count', names='設備類型', title='依填充次數統計', hole=0.4,
                                 color_discrete_sequence=MORANDI_PALETTE)
                fig3b_l.update_layout(font=dict(size=18), legend=dict(font=dict(size=16)), uniformtext_minsize=16, uniformtext_mode='show')
                fig3b_l.update_traces(textinfo='label+percent', textfont_size=18, textposition='outside',
                                      hovertemplate='<b>%{label}</b><br>填充次數: %{value} 次<br>佔比: %{percent:.1%}<extra></extra>')
                
                eq_weight = df_c3.groupby('設備類型')['冷媒填充量'].sum().reset_index()
                fig3b_r = px.pie(eq_weight, values='冷媒填充量', names='設備類型', title='依填充重量統計', hole=0.4,
                                 color_discrete_sequence=MORANDI_PALETTE)
                fig3b_r.update_layout(font=dict(size=18), legend=dict(font=dict(size=16)), uniformtext_minsize=16, uniformtext_mode='show')
                fig3b_r.update_traces(textinfo='label+percent', textfont_size=18, textposition='outside',
                                      hovertemplate='<b>%{label}</b><br>填充重量: %{value:.1f} kg<br>佔比: %{percent:.1%}<extra></extra>')
                
                with c3_l: st.plotly_chart(fig3b_l, use_container_width=True)
                with c3_r: st.plotly_chart(fig3b_r, use_container_width=True)
            else:
                st.info("無資料")
            
            st.markdown("---")
            
            # Chart 4
            st.subheader("🌍 全校冷媒填充碳排放量(公噸二氧化碳當量)結構")
            # V275: 4 decimal places
            fig_tree = px.treemap(df_year, path=['校區', '填報單位名稱'], values='排放量(公噸)', 
                                  color='校區', color_discrete_sequence=MORANDI_PALETTE)
            fig_tree.update_traces(texttemplate='%{label}<br>%{value:.4f}<br>%{percentRoot:.1%}', textfont=dict(size=24))
            st.plotly_chart(fig_tree, use_container_width=True)
            
        else:
            st.info("該年度無資料")

# ==========================================
# 7. 主程式入口
# ==========================================
if __name__ == "__main__":
    if username == 'admin':
        main_tabs = st.tabs(["👑 管理員後台", "❄️ 外部填報系統"])
        with main_tabs[0]:
            render_admin_dashboard()
        with main_tabs[1]:
            render_user_interface()
    else:
        render_user_interface()