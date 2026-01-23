import streamlit as st
import pandas as pd
import altair as alt
from datetime import datetime, date, timedelta
import gspread
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import re
import time

# ==========================================
# 0. 系統設定
# ==========================================
st.set_page_config(page_title="冷媒填報 - 嘉義大學", page_icon="❄️", layout="wide")

def get_taiwan_time():
    return datetime.utcnow() + timedelta(hours=8)

# 初始化 Session State
if 'uploader_key' not in st.session_state:
    st.session_state['uploader_key'] = 0

# 莫蘭迪色系
MORANDI_COLORS = ['#889EAF', '#A3B18A', '#D4A373', '#E07A5F', '#B5838D', '#FFB4A2', '#E5989B', '#6D6875', '#F2CC8F', '#81B29A']

# ==========================================
# 1. 內建備援資料 (Fallback Data)
# ==========================================
FALLBACK_UNIT_DATA = {'教務處': ['教務長室/副教務長室/專門委員室', '註冊與課務組', '教學發展組', '招生與出版組', '綜合行政組', '通識教育中心', '民雄教務'], '學生事務處': ['學務長室/副學務長室', '住宿服務組', '生活輔導組', '課外活動組', '學生輔導中心', '學生職涯發展中心', '衛生保健組', '原住民族學生資源中心', '特殊教育學生資源中心', '民雄學務'], '總務處': ['總務長室/副總務長室/簡任秘書室', '事務組', '出納組', '文書組', '資產經營管理組', '營繕組', '民雄總務', '新民聯辦', '駐衛警察隊'], '研究發展處': ['研發長室/副研發長室', '綜合企劃組', '學術發展組', '校務研究組'], '產學營運及推廣處': ['產學營運及推廣處長室', '行政管理組', '產學創育推廣中心'], '國際事務處': ['國際事務長室', '境外生事務組', '國際合作組'], '圖書資訊處': ['圖資長室', '圖資管理組', '資訊網路組', '諮詢服務組', '系統資訊組', '民雄圖書資訊', '新民分館', '民雄分館'], '校長室': ['校長室'], '行政副校長室': ['行政副校長室'], '學術副校長室': ['學術副校長室'], '國際副校長室': ['國際副校長室'], '秘書室': ['綜合業務組', '公共關係組', '校友服務組'], '體育室': ['蘭潭場館', '民雄場館', '林森場館', '新民場館'], '主計室': ['主計室'], '人事室': ['人事室'], '環境保護及安全管理中心': ['環境保護及安全管理中心'], '師資培育中心': ['師資培育中心主任室', '教育課程組', '實習輔導組', '綜合行政組'], '語言中心': ['主任室', '蘭潭語言中心', '民雄語言中心', '新民語言中心'], '理工學院': ['理工學院辦公室', '應用數學系', '電子物理學系', '應用化學系', '資訊工程學系', '生物機電工程學系', '土木與水資源工程學系', '水工與材料試驗場', '電機工程學系', '機械與能源工程學系'], '農學院': ['農學院辦公室', '農藝學系', '園藝學系', '森林暨自然資源學系', '木質材料與設計學系', '動物科學系', '農業經濟學系', '生物農業科技學系', '景觀學系', '植物醫學系', '農場管理進修學士學位學程'], '生命科學院': ['生命科學院辦公室', '食品科學系', '水生生物科學系', '生物資源學系', '生化科技學系', '微生物免疫與生物藥學系'], '管理學院': ['管理學院辦公室', '企業管理學系', '應用經濟學系', '生物事業管理學系', '資訊管理學系', '財務金融學系', '行銷與觀光管理學系', '全英文授課觀光暨管理學士學位學程'], '獸醫學院': ['獸醫學院辦公室', '獸醫學系', '雲嘉南動物疾病診斷中心', '動物醫院'], '師範學院': ['師範學院辦公室', '教育學系', '輔導與諮商學系', '體育與健康休閒學系', '特殊教育學系', '幼兒教育學系', '教育行政與政策發展研究所', '數理教育研究所'], '人文藝術學院': ['人文藝術學院辦公室', '中國文學系', '外國語言學系', '應用歷史學系', '視覺藝術學系', '音樂學系']}
FALLBACK_BUILDING_DATA = {'蘭潭校區': ['A01行政中心', 'A02森林館', 'A03動物科學館', 'A04農園館', 'A05工程館', 'A06食品科學館', 'A07嘉禾館', 'A08瑞穗館', 'A09游泳池', 'A10機械與能源工程學系創新育成大樓', 'A11木材利用工廠', 'A12動物試驗場', 'A13司令台', 'A14學生活動中心', 'A15電物一館', 'A16理工大樓', 'A17應化一館', 'A18A應化二館', 'A18B電物二館', 'A19農藝場管理室', 'A20國際交流學園', 'A21水工與材料試驗場', 'A22食品加工廠', 'A23機電館', 'A24生物資源館', 'A25生命科學館', 'A26農業科學館', 'A27植物醫學系館', 'A28水生生物科學館', 'A29園藝場管理室', 'A30園藝技藝中心', 'A31圖書資訊館', 'A32綜合教學大樓', 'A33生物農業科技二館', 'A34嘉大植物園', 'A35生技健康館', 'A36景觀學系大樓', 'A37森林生物多樣性館', 'A38動物產品研發推廣中心', 'A39學生活動廣場', 'A40焚化爐設備車倉庫', 'A41生物機械產業實驗室', 'A44有機蔬菜溫室', 'A45蝴蝶蘭溫室', 'A46魚類保育研究中心', 'A71員工單身宿舍', 'A72學苑餐廳', 'A73學一舍', 'A74學二舍', 'A75學三舍', 'A76學五舍', 'A77學六舍', 'A78農產品展售中心', 'A79綠建築', 'A80嘉大昆蟲館', 'A81蘭潭招待所', 'A82警衛室'], '民雄校區': ['B01創意樓', 'B02大學館', 'B03教育館', 'B04新藝樓', 'B06警衛室', 'B07鍋爐間', 'B08司令台', 'B09加氯室', 'B10游泳池', 'B12工友室', 'BA行政大樓', 'BB初等教育館', 'BC圖書館', 'BD樂育堂', 'BE學人單身宿舍', 'BF綠園二舍', 'BG餐廳', 'BH綠園一舍', 'BI科學館', 'BJ人文館', 'BK音樂館', 'BL藝術館', 'BM文薈廳', 'BN社團教室'], '林森校區': ['C01警衛室', 'C02司令台', 'CA第一棟大樓', 'CB進修部大樓', 'CD國民輔導大樓', 'CE第二棟大樓', 'CF實輔室', 'CG圖書館', 'CH視聽教室', 'CI明德齋', 'CK餐廳', 'CL青雲齋', 'CN樂育堂', 'CP空大學習指導中心'], '新民校區': ['D01管理學院大樓A棟', 'D02管理學院大樓B棟', 'D03明德樓', 'D04獸醫館(獸醫學系、動物醫院、雲嘉南動物疾病診斷中心)', 'D05游泳池', 'D06溫室', 'D07司令台', 'D08警衛室'], '社口林場': ['E01林場實習館'], '林森校區-民國路': ['F01民國路進德樓']}
FALLBACK_EQUIP_TYPES = ['冰水主機', '冰箱', '冷凍櫃', '冷氣', '冷藏櫃', '飲水機']
FALLBACK_REF_TYPES = ['HFC-1234yf 或 R-1234yf (2,3,3,3-四氟1-丙烯)，CF3CF=CH2', 'HFC-125 或 R-125 (1,1,1,2,2-五氟乙烷)，CHF2CF3', 'HFC-134a 或 R-134a (1,1,1,2-四氟乙烷)，CH2FCF3', 'HFC-143a 或 R-143a (1,1,1-三氟乙烷)，CH3CF3', 'HFC-23 或 R-23 (三氟甲烷)，CHF3', 'HFC-245fa 或 R-245fa (1,1,1,3,3-五氟丙烷)，CHF2CH2CF3', 'HFC-32 或 R-32 (二氟甲烷)，CH2F2', 'R-402A，HFC-125/HC-290/HCFC-22(60.0/2.0/38.0)', 'R-407D，HFC-32/HFC-125/HFC-134a(15.0/15.0/70.0)', 'R-411A，HC-1270/HCFC-22/HFC-152a(1.5/87.5/11.0)', 'R-507A，HFC-125/HFC-143a(50.0/50.0)', 'R-508A，HFC-23/PFC-116(39.0/61.0)', 'R-508B，HFC-23/PFC-116(46.0/54.0)', 'R404a，HFC-125/HFC-143a/HFC-134a(44.0/52.0/4.0)', 'R407c，HFC-32/HFC-125/HFC-134a(23.0/25.0/52.0)', 'R408a，HFC-125/HFC-143a/HCFC-22(7.0/46.0/47.0)', 'R410a，HFC-32/HFC-125(50.0/50.0)']
FALLBACK_GWP_MAP = {'HFC-1234yf 或 R-1234yf (2,3,3,3-四氟1-丙烯)，CF3CF=CH2': 0.0, 'HFC-125 或 R-125 (1,1,1,2,2-五氟乙烷)，CHF2CF3': 3170.0, 'HFC-134a 或 R-134a (1,1,1,2-四氟乙烷)，CH2FCF3': 1300.0, 'HFC-143a 或 R-143a (1,1,1-三氟乙烷)，CH3CF3': 4800.0, 'HFC-245fa 或 R-245fa (1,1,1,3,3-五氟丙烷)，CHF2CH2CF3': 858.0, 'R404a，HFC-125/HFC-143a/HFC-134a(44.0/52.0/4.0)': 3942.8, 'R407c，HFC-32/HFC-125/HFC-134a(23.0/25.0/52.0)': 1624.21, 'R-407D，HFC-32/HFC-125/HFC-134a(15.0/15.0/70.0)': 1487.05, 'R408a，HFC-125/HFC-143a/HCFC-22(7.0/46.0/47.0)': 2429.9, 'R410a，HFC-32/HFC-125(50.0/50.0)': 1923.5, 'R-507A，HFC-125/HFC-143a(50.0/50.0)': 3985.0, 'R-508A，HFC-23/PFC-116(39.0/61.0)': 11607.0, 'R-508B，HFC-23/PFC-116(46.0/54.0)': 11698.0, 'HFC-23 或 R-23 (三氟甲烷)，CHF3': 12400.0, 'HFC-32 或 R-32 (二氟甲烷)，CH2F2': 677.0, 'R-411A，HC-1270/HCFC-22/HFC-152a(1.5/87.5/11.0)': 0.0, 'R-402A，HFC-125/HC-290/HCFC-22(60.0/2.0/38.0)': 0.0}

# ==========================================
# 2. CSS 樣式 (UI 美化區 - 橘色分頁與莫蘭迪色)
# ==========================================
st.markdown("""
<style>
    /* 1. 分頁標籤放大且顏色醒目 (橘底白字) */
    button[data-baseweb="tab"] {
        font-size: 1.25rem !important;
        font-weight: 700 !important;
        background-color: #E67E22 !important; /* 橘色底 */
        color: white !important;
        border-radius: 8px 8px 0 0 !important;
        margin-right: 5px !important;
        padding: 8px 20px !important;
        border: none !important;
    }
    
    /* 選中狀態的分頁 */
    button[data-baseweb="tab"][aria-selected="true"] {
        background-color: #D35400 !important; /* 深橘色 */
        color: white !important;
        border-bottom: 3px solid #BA4A00 !important;
    }
    
    /* 分頁文字 */
    button[data-baseweb="tab"] div p {
        font-size: 1.25rem !important;
        font-weight: 700 !important;
        color: white !important;
    }
    
    /* 莫蘭迪色標題區塊 */
    .morandi-header {
        background-color: #EBF5FB; color: #2E4053; padding: 15px; border-radius: 8px; border-left: 8px solid #5499C7;
        font-size: 1.35rem; font-weight: 700; margin-top: 25px; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }

    /* 個資聲明區塊 */
    .privacy-box {
        background-color: #F8F9F9; border: 1px solid #BDC3C7; padding: 20px; border-radius: 8px;
        font-size: 0.95rem; color: #566573; line-height: 1.8; margin-bottom: 15px;
    }
    
    /* 誤繕提醒文字 */
    .correction-note { color: #566573; font-size: 0.9rem; margin-top: -10px; margin-bottom: 20px; }
    
    /* 個資聲明勾選文字 */
    [data-testid="stCheckbox"] label p { font-size: 1.15rem !important; font-weight: 700 !important; color: #2E4053 !important; }

    /* 橫式資訊卡 (2:8) - 前台 */
    .horizontal-card {
        display: flex; border: 1px solid #BDC3C7; border-radius: 12px; overflow: hidden; margin-bottom: 25px;
        box-shadow: 0 4px 8px rgba(0,0,0,0.08); background-color: #FFFFFF; min-height: 250px;
    }
    .card-left {
        flex: 2; background-color: #34495E; color: #FFFFFF; display: flex; flex-direction: column;
        justify-content: center; align-items: center; padding: 20px; text-align: center; border-right: 1px solid #2C3E50;
    }
    .dept-text { font-size: 1.4rem; font-weight: 700; margin-bottom: 8px; }
    .unit-text { font-size: 1.1rem; font-weight: 500; opacity: 0.9; }
    .card-right { flex: 8; padding: 20px 30px; display: flex; flex-direction: column; justify-content: center; }
    .info-row { display: flex; align-items: flex-start; padding: 10px 0; font-size: 1rem; color: #566573; border-bottom: 1px dashed #F2F3F4; }
    .info-row:last-child { border-bottom: none; }
    .info-icon { margin-right: 12px; font-size: 1.1rem; width: 25px; text-align: center; margin-top: 2px; }
    .info-label { font-weight: 600; margin-right: 10px; min-width: 150px; color: #2E4053; }
    .info-value { font-weight: 500; color: #17202A; flex: 1; }
    
    /* KPI 圖卡 - 後台 */
    .kpi-card {
        background-color: white; border-radius: 10px; padding: 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        text-align: center; border: 1px solid #E0E0E0; margin-bottom: 15px;
    }
    .kpi-title { font-size: 1.1rem; color: #555; margin-bottom: 5px; font-weight: 600; }
    .kpi-value { font-size: 2.2rem; color: #2C3E50; font-weight: 800; }
    .kpi-unit { font-size: 1rem; color: #888; margin-left: 5px; }
    .co2-card .kpi-value { color: #C0392B; }
    
    /* 上傳區樣式 */
    [data-testid="stFileUploaderDropzone"] { background-color: #D6EAF8; border: 2px dashed #2E86C1; padding: 20px; }
</style>
""", unsafe_allow_html=True)

# 3. 身份驗證
if st.session_state.get("authentication_status") is not True:
    st.warning("🔒 請先至首頁 (Hello) 登入系統")
    st.stop()

# 4. 資料庫連線
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
except Exception as e:
    st.error(f"❌ 資料庫連線失敗: {e}")
    st.stop()

# 5. 資料讀取 (全功能載入)
@st.cache_data(ttl=600)
def load_data_all_robust():
    try:
        ws_units = sh_ref.worksheet("單位資訊")
        unit_data = ws_units.get_all_values()
        unit_dict = {}
        if len(unit_data) > 1:
            for row in unit_data[1:]:
                if len(row) >= 2:
                    dept = str(row[0]).strip(); unit = str(row[1]).strip()
                    if dept and unit:
                        if dept not in unit_dict: unit_dict[dept] = []
                        if unit not in unit_dict[dept]: unit_dict[dept].append(unit)
        
        ws_buildings = sh_ref.worksheet("建築物清單")
        building_data = ws_buildings.get_all_values()
        build_dict = {}
        if len(building_data) > 1:
            for row in building_data[1:]:
                if len(row) >= 2:
                    campus = str(row[0]).strip(); b_name = str(row[1]).strip()
                    if campus and b_name:
                        if campus not in build_dict: build_dict[campus] = []
                        if b_name not in build_dict[campus]: build_dict[campus].append(b_name)

        ws_types = sh_ref.worksheet("設備類型")
        type_data = ws_types.get_all_values()
        e_types = sorted([row[0] for row in type_data[1:] if row]) if len(type_data) > 1 else []
        
        ws_coef = sh_ref.worksheet("冷媒係數表")
        coef_data = ws_coef.get_all_values()
        r_types = []; gwp_map = {}
        if len(coef_data) > 1:
            try:
                name_idx = 1; gwp_idx = 2
                for row in coef_data[1:]:
                    if len(row) > gwp_idx and row[name_idx]:
                        r_name = row[name_idx].strip()
                        gwp_val_str = row[gwp_idx].replace(',', '').strip()
                        if not gwp_val_str.replace('.', '', 1).isdigit(): gwp_val = 0.0
                        else: gwp_val = float(gwp_val_str)
                        r_types.append(r_name); gwp_map[r_name] = gwp_val
            except:
                r_types = sorted([row[1] for row in coef_data[1:] if len(row) > 1 and row[1]])

        try: ws_records = sh_ref.worksheet("冷媒填報紀錄")
        except:
            ws_records = sh_ref.add_worksheet(title="冷媒填報紀錄", rows="1000", cols="15")
            ws_records.append_row(["填報時間","填報人","填報人分機","校區","所屬單位","填報單位名稱","建築物名稱","辦公室編號","維修日期","設備類型","設備品牌型號","冷媒種類","冷媒填充量","備註","佐證資料"])
            
        records_data = ws_records.get_all_values()
        if len(records_data) > 1:
            raw_headers = records_data[0]
            col_mapping = {}
            for h in raw_headers:
                clean_h = str(h).strip()
                if "填充量" in clean_h or "重量" in clean_h: col_mapping[h] = "冷媒填充量"
                elif "種類" in clean_h or "品項" in clean_h: col_mapping[h] = "冷媒種類"
                elif "日期" in clean_h or "維修" in clean_h: col_mapping[h] = "維修日期"
                else: col_mapping[h] = clean_h
            
            df_records = pd.DataFrame(records_data[1:], columns=raw_headers)
            df_records.rename(columns=col_mapping, inplace=True)
            
            df_records['冷媒填充量'] = pd.to_numeric(df_records['冷媒填充量'], errors='coerce').fillna(0)
            df_records['維修日期'] = pd.to_datetime(df_records['維修日期'], errors='coerce')
            df_records['年份'] = df_records['維修日期'].dt.year.fillna(datetime.now().year).astype(int)
            
            def get_gwp(rtype): return gwp_map.get(rtype, 0)
            df_records['GWP'] = df_records['冷媒種類'].apply(get_gwp)
            df_records['排放量(kgCO2e)'] = df_records['冷媒填充量'] * df_records['GWP']
            df_records['碳排放量(噸)'] = df_records['排放量(kgCO2e)'] / 1000.0
            df_records['完整單位名稱'] = df_records['所屬單位'] + " - " + df_records['填報單位名稱']
            
        else:
            df_records = pd.DataFrame(columns=["填報時間","填報人","填報人分機","校區","所屬單位","填報單位名稱","建築物名稱","辦公室編號","維修日期","設備類型","設備品牌型號","冷媒種類","冷媒填充量","備註","佐證資料","排放量(kgCO2e)","碳排放量(噸)","年份"])

        return unit_dict, build_dict, e_types, sorted(r_types), gwp_map, df_records, True

    except Exception as e:
        return FALLBACK_UNIT_DATA, FALLBACK_BUILDING_DATA, FALLBACK_EQUIP_TYPES, sorted(FALLBACK_REF_TYPES), FALLBACK_GWP_MAP, pd.DataFrame(), False

# 呼叫載入
unit_dict, build_dict, e_types, r_types, gwp_map, df_records, load_success = load_data_all_robust()

# ==========================================
# 6. 介面控制邏輯
# ==========================================
if not load_success:
    st.warning("⚠️ 網路連線不穩，目前使用備援資料模式。")

is_admin = st.session_state.get("username") == "admin"
mode = "📝 一般填報/查詢"

if is_admin:
    mode = st.radio("檢視模式", ["📝 一般填報/查詢", "⚙️ 管理員後台"], horizontal=True, label_visibility="collapsed")
    st.markdown("---")

# ==========================================
# 7. 模式 A: 前台 (填報 + 查詢)
# ==========================================
if mode == "📝 一般填報/查詢":
    
    if st.button("🔄 刷新資料庫 (若更新Excel請點此)", type="secondary"):
        st.cache_data.clear()
        st.rerun()

    tabs = st.tabs(["📝 新增填報", "📋 申報動態查詢"])

    # --- Tab 1: 新增填報 ---
    with tabs[0]:
        st.markdown('<div class="morandi-header">填報單位基本資訊區</div>', unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        unit_depts = sorted(unit_dict.keys())
        sel_dept = c1.selectbox("所屬單位", unit_depts, index=None, placeholder="請選擇單位...", key="k_dept")
        unit_names = []
        if sel_dept: unit_names = sorted(unit_dict.get(sel_dept, []))
        sel_unit_name = c2.selectbox("填報單位名稱", unit_names, index=None, placeholder="請先選擇所屬單位...", key="k_unit")
        c3, c4 = st.columns(2)
        name = c3.text_input("填報人", key="k_name")
        ext = c4.text_input("填報人分機", key="k_ext")
        
        st.markdown('<div class="morandi-header">冷媒設備所在位置資訊區</div>', unsafe_allow_html=True)
        loc_campuses = sorted(build_dict.keys())
        sel_loc_campus = st.selectbox("填報單位所在校區", loc_campuses, index=None, placeholder="請選擇校區...", key="k_campus")
        c6, c7 = st.columns(2)
        buildings = []
        if sel_loc_campus: buildings = sorted(build_dict.get(sel_loc_campus, []))
        sel_build = c6.selectbox("建築物名稱", buildings, index=None, placeholder="請先選擇校區...", key="k_build")
        office = c7.text_input("辦公室編號", placeholder="例如：202辦公室、306研究室", key="k_office")
        
        st.markdown('<div class="morandi-header">冷媒設備填充資訊區</div>', unsafe_allow_html=True)
        c8, c9 = st.columns(2)
        r_date = c8.date_input("維修日期 (統一填寫發票日期)", datetime.today(), key="k_date")
        sel_etype = c9.selectbox("設備類型", e_types, index=None, placeholder="請選擇...", key="k_etype")
        c10, c11 = st.columns(2)
        e_model = c10.text_input("設備品牌型號", placeholder="例如：國際 CS-100FL+CU-100FLC", key="k_model")
        sel_rtype = c11.selectbox("冷媒種類", r_types, index=None, placeholder="請選擇...", key="k_rtype")
        amount = st.number_input("冷媒填充量 (公斤)", min_value=0.0, step=0.1, format="%.2f", key="k_amount")
        
        st.markdown("請上傳冷媒填充單據佐證資料")
        f_file = st.file_uploader("上傳佐證 (必填)", type=['pdf', 'jpg', 'png'], label_visibility="collapsed", key=f"uploader_{st.session_state.uploader_key}")
        
        st.markdown("---")
        note = st.text_input("備註內容", placeholder="備註 (選填)", key="k_note")
        st.markdown('<div class="correction-note">如有資料誤繕情形，請重新登錄1次資訊，並於備註欄填寫：「前筆資料誤繕，請刪除。」，管理單位將協助刪除誤打資訊</div>', unsafe_allow_html=True)
        
        st.markdown("""<div class="privacy-box"><strong>📜 個人資料蒐集、處理及利用告知聲明</strong><br>1. 蒐集機關：國立嘉義大學。<br>2. 蒐集目的：進行本校冷媒設備之冷媒填充紀錄管理、校園溫室氣體（碳）盤查統計、稽核佐證資料蒐集及後續能源使用分析。<br>3. 個資類別：填報人姓名。<br>4. 利用期間：姓名保留至填報年度後第二年1月1日，期滿即進行「去識別化」刪除，其餘數據永久保存。<br>5. 利用對象：本校教師、行政人員及碳盤查查驗人員。<br>6. 您有權依個資法請求查詢、更正或刪除您的個資。如不提供，將無法完成填報。</div>""", unsafe_allow_html=True)
        
        agree = st.checkbox("我已閱讀並同意個資聲明，且確認所填資料無誤。", key="k_agree")
        submitted = st.button("🚀 確認送出", type="primary", use_container_width=True)
        
        if submitted:
            if not agree: st.error("❌ 請勾選同意聲明")
            elif not sel_dept or not sel_unit_name: st.warning("⚠️ 請完整選擇【基本資訊】中的單位資訊")
            elif not name or not ext: st.warning("⚠️ 請填寫填報人與分機")
            elif not sel_loc_campus or not sel_build: st.warning("⚠️ 請完整選擇【位置資訊】中的校區與建築物")
            elif not sel_etype or not sel_rtype: st.warning("⚠️ 請選擇設備類型與冷媒種類")
            elif not f_file: st.error("⚠️ 請上傳佐證資料")
            else:
                try:
                    ws_target = sh_ref.worksheet("冷媒填報紀錄")
                    f_file.seek(0); f_ext = f_file.name.split('.')[-1]
                    clean_name = f"{sel_loc_campus}_{sel_dept}_{sel_unit_name}_{r_date}_{sel_etype}_{sel_rtype}.{f_ext}"
                    meta = {'name': clean_name, 'parents': [REF_FOLDER_ID]}
                    media = MediaIoBaseUpload(f_file, mimetype=f_file.type, resumable=True)
                    file = drive_service.files().create(body=meta, media_body=media, fields='webViewLink').execute()
                    link = file.get('webViewLink')
                    
                    current_time = get_taiwan_time().strftime("%Y-%m-%d %H:%M:%S")
                    row_data = [current_time, name, ext, sel_loc_campus, sel_dept, sel_unit_name, sel_build, office, str(r_date), sel_etype, e_model, sel_rtype, amount, note, link]
                    ws_target.append_row(row_data)
                    
                    st.success("✅ 冷媒填報成功！")
                    st.balloons()
                    st.cache_data.clear()
                    time.sleep(1.5)
                    
                    keys_to_clear = ["k_etype", "k_model", "k_rtype", "k_amount", "k_note", "k_agree"]
                    for key in keys_to_clear:
                        if key in st.session_state: del st.session_state[key]
                    st.session_state['uploader_key'] += 1
                    st.rerun()
                except Exception as e: st.error(f"上傳或寫入失敗: {e}")

    # --- Tab 2: 申報動態查詢 ---
    with tabs[1]:
        st.markdown('<div class="morandi-header">📋 申報動態查詢</div>', unsafe_allow_html=True)
        if df_records.empty: st.info("目前尚無填報紀錄。")
        else:
            st.markdown("##### 🔍 查詢條件設定")
            c_f1, c_f2 = st.columns(2)
            depts = sorted(df_records['所屬單位'].dropna().unique())
            sel_q_dept = c_f1.selectbox("所屬單位 (必選)", depts, index=None, placeholder="請選擇...")
            units = []
            if sel_q_dept: units = sorted(df_records[df_records['所屬單位'] == sel_q_dept]['填報單位名稱'].dropna().unique())
            sel_q_unit = c_f2.selectbox("填報單位名稱 (必選)", units, index=None, placeholder="請選擇...")
            
            c_f3, c_f4 = st.columns(2)
            today = datetime.now().date()
            start_of_year = date(today.year, 1, 1)
            q_start_date = c_f3.date_input("查詢起始日期", value=start_of_year, max_value=today)
            q_end_date = c_f4.date_input("查詢結束日期", value=today, max_value=today)

            if sel_q_dept and sel_q_unit and q_start_date and q_end_date:
                if q_start_date > q_end_date: st.warning("⚠️ 起始日期不能晚於結束日期")
                else:
                    start_dt = pd.Timestamp(q_start_date); end_dt = pd.Timestamp(q_end_date)
                    mask = (df_records['所屬單位'] == sel_q_dept) & (df_records['填報單位名稱'] == sel_q_unit) & (df_records['維修日期'] >= start_dt) & (df_records['維修日期'] <= end_dt)
                    df_view = df_records[mask]
                    
                    if sel_q_dept == sel_q_unit: left_html = f'<div class="dept-text">{sel_q_dept}</div>'
                    else: left_html = f'<div class="dept-text">{sel_q_dept}</div><div class="unit-text">{sel_q_unit}</div>'
                    
                    total_count = len(df_view)
                    if total_count > 0:
                        campus_str = ", ".join(sorted(df_view['校區'].unique()))
                        builds = sorted(df_view['建築物名稱'].unique())
                        build_str = ", ".join(builds[:3]) + (f" 等{len(builds)}棟" if len(builds)>3 else "")
                        equip_str = ", ".join(sorted(df_view['設備類型'].unique()))
                        
                        fill_details = []
                        for _, row in df_view.iterrows(): fill_details.append(f"<div>• {row['冷媒種類']}：{row['冷媒填充量']:.2f} 公斤</div>")
                        fill_detail_html = "".join(fill_details)
                        
                        fill_stats = []
                        fill_summary = df_view.groupby('冷媒種類')['冷媒填充量'].sum().reset_index()
                        total_fill_all = fill_summary['冷媒填充量'].sum()
                        fill_stats.append(f"<div style='font-weight:700; color:#1F618D;'>• 總計：{total_fill_all:.2f} 公斤</div>")
                        for _, row in fill_summary.iterrows(): fill_stats.append(f"<div style='color:#566573;'>　- {row['冷媒種類']}：{row['冷媒填充量']:.2f} 公斤</div>")
                        fill_stats_html = "".join(fill_stats)
                        
                        total_emission = df_view['排放量(kgCO2e)'].sum()
                    else:
                        campus_str="無資料"; build_str="無資料"; equip_str="無資料"; fill_detail_html="無資料"; fill_stats_html="無資料"; total_emission=0.0

                    st.markdown("---")
                    st.markdown(f"""
                    <div class="horizontal-card">
                        <div class="card-left">{left_html}</div>
                        <div class="card-right">
                            <div class="info-row"><span class="info-icon">📅</span><span class="info-label">查詢起訖時間區間</span><span class="info-value">{q_start_date} ~ {q_end_date}</span></div>
                            <div class="info-row"><span class="info-icon">🏫</span><span class="info-label">所在校區</span><span class="info-value">{campus_str}</span></div>
                            <div class="info-row"><span class="info-icon">🏢</span><span class="info-label">建築物名稱</span><span class="info-value">{build_str}</span></div>
                            <div class="info-row"><span class="info-icon">❄️</span><span class="info-label">設備類型</span><span class="info-value">{equip_str}</span></div>
                            <div class="info-row"><span class="info-icon">📝</span><span class="info-label">冷媒填充資訊</span><span class="info-value">{fill_detail_html}</span></div>
                            <div class="info-row"><span class="info-icon">⚖️</span><span class="info-label">冷媒填充重量統計</span><span class="info-value">{fill_stats_html}</span></div>
                            <div class="info-row"><span class="info-icon">🌍</span><span class="info-label">碳排放量</span><span class="info-value" style="color:#C0392B;"><span style="font-size: 1.8rem; font-weight:900;">{total_emission:,.2f}</span> <span style="font-size: 1rem; font-weight:normal;">公斤二氧化碳當量</span></span></div>
                            <div class="info-row"><span class="info-icon">📊</span><span class="info-label">申報次數統計</span><span class="info-value">{total_count} 次</span></div>
                        </div>
                    </div>""", unsafe_allow_html=True)

                    st.markdown(f"##### 📋 {sel_q_dept} {sel_q_unit} 填報明細")
                    show_cols = ["維修日期", "校區", "建築物名稱", "設備類型", "設備品牌型號", "冷媒種類", "冷媒填充量", "排放量(kgCO2e)", "佐證資料"]
                    valid_cols = [c for c in show_cols if c in df_view.columns]
                    st.dataframe(df_view[valid_cols].sort_values("維修日期", ascending=False), use_container_width=True, column_config={"佐證資料": st.column_config.LinkColumn("佐證連結", display_text="開啟檔案")})
            else:
                if not (sel_q_dept and sel_q_unit): st.info("👈 請先選擇「所屬單位」與「填報單位名稱」以開始查詢。")

# ==========================================
# 8. 模式 B: 後台 (管理儀表板)
# ==========================================
elif mode == "⚙️ 管理員後台":
    back_tabs = st.tabs(["📝 申報資料異動", "📊 全校冷媒填充儀表板"])
    
    # --- Tab 1: 資料異動 (含下載) ---
    with back_tabs[0]:
        if df_records.empty:
            st.warning("目前尚無填報資料")
        else:
            c_down1, c_down2 = st.columns([4, 1])
            years = sorted(df_records['年份'].unique(), reverse=True)
            sel_year_dl = c_down1.selectbox("📅 選擇下載年度", ["全部"] + list(years), index=0)
            
            if sel_year_dl == "全部": df_dl = df_records
            else: df_dl = df_records[df_records['年份'] == sel_year_dl]
            
            # 下載按鈕排版優化：插入空行讓它往下移
            st.markdown("<br>", unsafe_allow_html=True)
            csv = df_dl.to_csv(index=False).encode('utf-8-sig')
            c_down2.download_button("📥 下載資料 (CSV)", data=csv, file_name=f"冷媒填報_{sel_year_dl}.csv", mime="text/csv", type="primary")
            
            st.markdown("---")
            st.markdown("### 🛠️ 線上資料編輯")
            st.info("💡 您可以直接在下方表格修改資料。修改完畢後，請務必點擊底部的「💾 儲存變更」按鈕，資料才會寫入資料庫。")
            
            # 編輯器 (排除後台計算欄位)
            editable_cols = ["填報時間","填報人","填報人分機","校區","所屬單位","填報單位名稱","建築物名稱","辦公室編號","維修日期","設備類型","設備品牌型號","冷媒種類","冷媒填充量","備註","佐證資料"]
            valid_edit_cols = [c for c in editable_cols if c in df_records.columns]
            edited_df = st.data_editor(df_records[valid_edit_cols], num_rows="dynamic", use_container_width=True, height=500)
            
            if st.button("💾 儲存變更", type="primary"):
                try:
                    # 連線
                    ws_target = sh_ref.worksheet("冷媒填報紀錄")
                    # 清空舊資料
                    ws_target.clear()
                    
                    # 處理數據：將所有日期時間物件轉為字串，避免 JSON 序列化錯誤
                    df_to_save = edited_df.copy()
                    for col in df_to_save.columns:
                        # 強制轉為字串
                        df_to_save[col] = df_to_save[col].astype(str)
                    
                    # 準備新資料 (含標題，轉為 List)
                    updated_data = [valid_edit_cols] + df_to_save.values.tolist()
                    
                    # 寫入
                    ws_target.update(updated_data)
                    
                    st.success("✅ 資料已成功更新！")
                    st.cache_data.clear()
                    time.sleep(1.5)
                    st.rerun()
                except Exception as e:
                    st.error(f"儲存失敗: {e}")

    # --- Tab 2: 儀表板 ---
    with back_tabs[1]:
        if df_records.empty: st.warning("目前尚無填報資料"); st.stop()

        c_top1, c_top2 = st.columns([1, 4])
        years = sorted(df_records['年份'].unique(), reverse=True)
        sel_year = c_top1.selectbox("📅 選擇年度", years, index=0)
        df_yr = df_records[df_records['年份'] == sel_year]
        
        st.markdown("---")
        
        total_count = len(df_yr)
        total_weight = df_yr['冷媒填充量'].sum()
        total_emission_tons = df_yr['碳排放量(噸)'].sum()
        
        k1, k2 = st.columns(2)
        k1.markdown(f"""<div class="kpi-card"><div class="kpi-title">📊 冷媒填充次數</div><span class="kpi-value">{total_count}</span><span class="kpi-unit">次</span></div>""", unsafe_allow_html=True)
        k2.markdown(f"""<div class="kpi-card"><div class="kpi-title">⚖️ 冷媒填充重量</div><span class="kpi-value">{total_weight:,.2f}</span><span class="kpi-unit">kg</span></div>""", unsafe_allow_html=True)
        st.markdown(f"""<div class="kpi-card co2-card"><div class="kpi-title">☁️ 碳排放量</div><span class="kpi-value">{total_emission_tons:,.4f}</span><span class="kpi-unit">公噸 CO2e</span></div>""", unsafe_allow_html=True)
        
        st.markdown("---")
        
        def campus_filter(key_suffix):
            campuses = ["全校"] + sorted(df_yr['校區'].unique().tolist())
            return st.radio("統計模式切換", campuses, horizontal=True, key=f"campus_{key_suffix}")

        st.subheader("3. 全校冷媒設備填充概況統計")
        sel_campus_1 = campus_filter("1")
        chart_data_1 = df_yr if sel_campus_1 == "全校" else df_yr[df_yr['校區'] == sel_campus_1]
        if not chart_data_1.empty:
            chart1 = alt.Chart(chart_data_1).mark_bar().encode(
                x=alt.X('冷媒種類', axis=alt.Axis(labelAngle=-45, titleFontSize=14, labelFontSize=12)),
                y=alt.Y('sum(冷媒填充量)', title='填充量 (kg)', axis=alt.Axis(titleFontSize=14, labelFontSize=12)),
                color=alt.Color('設備類型', scale=alt.Scale(range=MORANDI_COLORS), legend=alt.Legend(title="設備類型")),
                tooltip=['冷媒種類', '設備類型', alt.Tooltip('sum(冷媒填充量)', format=',.2f', title='填充量(kg)')]
            ).properties(height=400).configure_axis(grid=False)
            st.altair_chart(chart1, use_container_width=True)
        
        st.markdown("---")
        
        st.subheader("4. 年度全校前十大冷媒填充單位")
        top10_units = df_yr.groupby('完整單位名稱')['冷媒填充量'].sum().nlargest(10).index.tolist()
        df_top10 = df_yr[df_yr['完整單位名稱'].isin(top10_units)]
        if not df_top10.empty:
            chart2 = alt.Chart(df_top10).mark_bar().encode(
                x=alt.X('完整單位名稱', sort=top10_units, axis=alt.Axis(labelAngle=-45, title="單位名稱", titleFontSize=14, labelFontSize=12)),
                y=alt.Y('sum(冷媒填充量)', title='填充量 (kg)', axis=alt.Axis(titleFontSize=14, labelFontSize=12)),
                color=alt.Color('冷媒種類', scale=alt.Scale(range=MORANDI_COLORS), legend=alt.Legend(title="冷媒種類")),
                tooltip=['完整單位名稱', '冷媒種類', alt.Tooltip('sum(冷媒填充量)', format=',.2f')]
            ).properties(height=450)
            st.altair_chart(chart2, use_container_width=True)
            
        st.markdown("---")
        
        st.subheader("5. 冷媒填充設備類型統計 (次數佔比)")
        sel_campus_3 = campus_filter("3")
        chart_data_3 = df_yr if sel_campus_3 == "全校" else df_yr[df_yr['校區'] == sel_campus_3]
        if not chart_data_3.empty:
            equip_counts = chart_data_3['設備類型'].value_counts().reset_index()
            equip_counts.columns = ['設備類型', '次數']
            equip_counts['百分比'] = (equip_counts['次數'] / equip_counts['次數'].sum()).round(3)
            base = alt.Chart(equip_counts).encode(theta=alt.Theta("次數", stack=True))
            pie = base.mark_arc(outerRadius=120, innerRadius=80).encode(color=alt.Color("設備類型", scale=alt.Scale(range=MORANDI_COLORS)), order=alt.Order("次數", sort="descending"), tooltip=["設備類型", "次數", alt.Tooltip("百分比", format=".1%")])
            text = base.mark_text(radius=140, size=14).encode(text=alt.Text("百分比", format=".1%"), order=alt.Order("次數", sort="descending"), color=alt.value("#555"))
            st.altair_chart((pie + text).properties(height=400), use_container_width=True)
            
        st.markdown("---")
        
        st.subheader("6. 冷媒填充碳排放量占比")
        sel_campus_4 = campus_filter("4")
        chart_data_4 = df_yr if sel_campus_4 == "全校" else df_yr[df_yr['校區'] == sel_campus_4]
        if not chart_data_4.empty:
            emission_stats = chart_data_4.groupby('冷媒種類')['碳排放量(噸)'].sum().reset_index()
            emission_stats = emission_stats[emission_stats['碳排放量(噸)'] > 0]
            emission_stats['百分比'] = (emission_stats['碳排放量(噸)'] / emission_stats['碳排放量(噸)'].sum()).round(3)
            base_e = alt.Chart(emission_stats).encode(theta=alt.Theta("碳排放量(噸)", stack=True))
            pie_e = base_e.mark_arc(outerRadius=120, innerRadius=80).encode(color=alt.Color("冷媒種類", scale=alt.Scale(range=MORANDI_COLORS)), order=alt.Order("碳排放量(噸)", sort="descending"), tooltip=["冷媒種類", alt.Tooltip("碳排放量(噸)", format=".4f"), alt.Tooltip("百分比", format=".1%")])
            text_e = base_e.mark_text(radius=140, size=14).encode(text=alt.Text("百分比", format=".1%"), order=alt.Order("碳排放量(噸)", sort="descending"), color=alt.value("#555"))
            st.altair_chart((pie_e + text_e).properties(height=400), use_container_width=True)