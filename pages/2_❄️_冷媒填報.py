import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import gspread
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# ==========================================
# 0. 系統設定
# ==========================================
st.set_page_config(page_title="冷媒填報 - 嘉義大學", page_icon="❄️", layout="wide")

def get_taiwan_time():
    return datetime.utcnow() + timedelta(hours=8)

# ==========================================
# 1. 內建靜態選單資料 (資料庫 - 單位與位置)
# ==========================================
UNIT_DATA = {
    '教務處': ['教務長室/副教務長室/專門委員室', '註冊與課務組', '教學發展組', '招生與出版組', '綜合行政組', '通識教育中心', '民雄教務'],
    '學生事務處': ['學務長室/副學務長室', '住宿服務組', '生活輔導組', '課外活動組', '學生輔導中心', '學生職涯發展中心', '衛生保健組', '原住民族學生資源中心', '特殊教育學生資源中心', '民雄學務'],
    '總務處': ['總務長室/副總務長室/簡任秘書室', '事務組', '出納組', '文書組', '資產經營管理組', '營繕組', '民雄總務', '新民聯辦', '駐衛警察隊'],
    '研究發展處': ['研發長室/副研發長室', '綜合企劃組', '學術發展組', '校務研究組'],
    '產學營運及推廣處': ['產學營運及推廣處長室', '產學營運及推廣處-行政管理組', '產學營運及推廣處-產學創育推廣中心'],
    '國際事務處': ['國際事務長室', '境外生事務組', '國際合作組'],
    '圖書資訊處': ['圖資長室', '圖資管理組', '資訊網路組', '諮詢服務組', '系統資訊組', '民雄圖書資訊', '新民分館', '民雄分館'],
    '校長室': ['校長室'],
    '行政副校長室': ['行政副校長室'],
    '學術副校長室': ['學術副校長室'],
    '國際副校長室': ['國際副校長室'],
    '秘書室': ['綜合業務組', '公共關係組', '校友服務組'],
    '體育室': ['蘭潭場館', '民雄場館', '林森場館', '新民場館'],
    '主計室': ['主計室'],
    '人事室': ['人事室'],
    '環境保護及安全管理中心': ['環境保護及安全管理中心'],
    '師資培育中心': ['師資培育中心主任室', '教育課程組', '實習輔導組', '綜合行政組'],
    '語言中心': ['主任室', '蘭潭語言中心', '民雄語言中心', '新民語言中心'],
    '理工學院': ['理工學院辦公室', '應用數學系', '電子物理學系', '應用化學系', '資訊工程學系', '生物機電工程學系', '土木與水資源工程學系', '水工與材料試驗場', '電機工程學系', '機械與能源工程學系'],
    '農學院': ['農學院辦公室', '農業推廣中心', '農藝學系', '園藝學系', '森林暨自然資源學系', '木質材料與設計學系', '木材利用工廠', '動物科學系', '動物試驗場', '農業生物科技學系', '景觀學系', '植物醫學系', '農場管理進修學士學位學程', '農林實驗場管理中心', '園藝技藝中心', '農產品驗證中心'],
    '生命科學院': ['生命科學院辦公室', '食品科學系', '生化科技學系', '水生生物科學系', '生物資源學系', '微生物免疫與生物藥學系', '檢驗分析及技術推廣服務中心', '智慧食農教研中心', '中草藥暨微生物利用研發中心'],
    '管理學院': ['管理學院辦公室', '企業管理學系', '應用經濟學系', '科技管理學系', '資訊管理學系', '財務金融學系', '行銷與觀光管理學系', '管理學院EMBA', '外籍生全英觀光暨管理碩士學程'],
    '獸醫學院': ['獸醫學院', '獸醫系', '動物醫院', '動物疾病診斷中心'],
    '師範學院': ['師範學院辦公室', '教育學系', '數位學習設計與管理學系', '特殊教育學系', '實驗教育研究中心', '輔導與諮商學系', '輔導與諮商學系-林森校區', '家庭與社區諮商中心', '幼兒教育學系', '體育與健康休閒學系', '數理教育研究所', '教育行政與政策發展研究所'],
    '人文藝術學院': ['人文藝術學院辦公室', '中國文學系', '外國語言學系', '應用歷史學系', '視覺藝術學系', '音樂學系', '台灣文化研究中心', '人文藝術中心']
}

BUILDING_DATA = {
    '蘭潭校區': ['A01行政中心', 'A02森林館', 'A03動物科學館', 'A04農園館', 'A05工程館', 'A06食品科學館', 'A07嘉禾館', 'A08瑞穗館', 'A09游泳池', 'A10機械與能源工程學系創新育成大樓', 'A11木材利用工廠', 'A12動物試驗場', 'A13司令台', 'A14學生活動中心', 'A15電物一館', 'A16理工大樓', 'A17應化一館', 'A18A應化二館', 'A18B電物二館', 'A19農藝場管理室', 'A20國際交流學園', 'A21水工與材料試驗場', 'A22食品加工廠', 'A23機電館', 'A24生物資源館', 'A25生命科學館', 'A26農業科學館', 'A27植物醫學系館', 'A28水生生物科學館', 'A29園藝場管理室', 'A30園藝技藝中心', 'A31圖書資訊館', 'A32綜合教學大樓', 'A33生物農業科技二館', 'A34嘉大植物園', 'A35生技健康館', 'A36景觀學系大樓', 'A37森林生物多樣性館', 'A38動物產品研發推廣中心', 'A39學生活動廣場', 'A40焚化爐設備車倉庫', 'A41生物機械產業實驗室', 'A44有機蔬菜溫室', 'A45蝴蝶蘭溫室', 'A46魚類保育研究中心', 'A71員工單身宿舍', 'A72學苑餐廳', 'A73學一舍', 'A74學二舍', 'A75學三舍', 'A76學五舍', 'A77學六舍', 'A78農產品展售中心', 'A79綠建築', 'A80嘉大昆蟲館', 'A81蘭潭招待所', 'A82警衛室'],
    '民雄校區': ['B01創意樓', 'B02大學館', 'B03教育館', 'B04新藝樓', 'B06警衛室', 'B07鍋爐間', 'B08司令台', 'B09加氯室', 'B10游泳池', 'B12工友室', 'BA行政大樓', 'BB初等教育館', 'BC圖書館', 'BD樂育堂', 'BE學人單身宿舍', 'BF綠園二舍', 'BG餐廳', 'BH綠園一舍', 'BI科學館', 'BJ人文館', 'BK音樂館', 'BL藝術館', 'BM文薈廳', 'BN社團教室'],
    '林森校區': ['C01警衛室', 'C02司令台', 'CA第一棟大樓', 'CB進修部大樓', 'CD國民輔導大樓', 'CE第二棟大樓', 'CF實輔室', 'CG圖書館', 'CH視聽教室', 'CI明德齋', 'CK餐廳', 'CL青雲齋', 'CN樂育堂', 'CP空大學習指導中心'],
    '新民校區': ['D01管理學院大樓A棟', 'D02管理學院大樓B棟', 'D03明德樓', 'D04嘉大動物醫院', 'D05游泳池', 'D06溫室', 'D07司令台', 'D08警衛室'],
    '蘭潭校區-社口林場': ['E01林場實習館'],
    '林森校區-民國路': ['F01民國路游泳池']
}

# ==========================================
# 2. CSS 樣式 (UI 美化區)
# ==========================================
st.markdown("""
<style>
    /* 1. 分頁標籤放大 */
    button[data-baseweb="tab"] div p {
        font-size: 1.2rem !important;
        font-weight: 700 !important;
    }
    
    /* 2. 莫蘭迪色標題區塊 */
    .morandi-header {
        background-color: #EBF5FB;
        color: #2E4053;
        padding: 15px;
        border-radius: 8px;
        border-left: 8px solid #5499C7;
        font-size: 1.35rem;
        font-weight: 700;
        margin-top: 25px;
        margin-bottom: 20px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }

    /* 3. 個資聲明區塊 */
    .privacy-box {
        background-color: #F8F9F9;
        border: 1px solid #BDC3C7;
        padding: 20px;
        border-radius: 8px;
        font-size: 0.95rem;
        color: #566573;
        line-height: 1.8;
        margin-bottom: 15px;
    }
    
    /* 誤繕提醒文字樣式 */
    .correction-note {
        color: #566573; 
        font-size: 0.9rem; 
        margin-top: -10px; 
        margin-bottom: 20px;
    }
    
    /* 個資聲明勾選文字樣式 (大字、深藍) */
    [data-testid="stCheckbox"] label p {
        font-size: 1.1rem !important;
        font-weight: 700 !important;
        color: #2E4053 !important;
    }

    /* 整合式資訊卡樣式 */
    .info-card {
        border-radius: 10px;
        overflow: hidden;
        border: 1px solid #D5DBDB;
        margin-bottom: 20px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
    }
    
    /* 卡片標題區 (莫蘭迪藍底) */
    .card-header {
        background-color: #EBF5FB; 
        padding: 15px 20px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        border-bottom: 1px solid #D6EAF8;
    }
    .card-title {
        font-size: 1.3rem; 
        font-weight: 700; 
        color: #2E4053;
    }
    .emission-value {
        font-size: 2rem; 
        font-weight: 800; 
        color: #C0392B; /* 莫蘭迪紅 */
    }
    .emission-unit {
        font-size: 1rem; 
        color: #566573;
        font-weight: normal;
        margin-left: 5px;
    }
    
    /* 卡片內容區 (白底) */
    .card-body {
        background-color: #FFFFFF;
        padding: 15px 20px;
    }
    .fill-row {
        display: flex;
        justify-content: space-between;
        padding: 8px 0;
        border-bottom: 1px dashed #EAEDED;
    }
    .fill-row:last-child {
        border-bottom: none;
    }
    .fill-type {
        font-size: 1.05rem; 
        color: #2E4053;
        font-weight: 600;
    }
    .fill-amount {
        font-size: 1.05rem; 
        color: #2874A6;
        font-weight: 600;
    }
    
    /* 卡片底部 (極淺莫蘭迪底) */
    .card-footer {
        background-color: #F8FBFD; /* 極淺色調 */
        padding: 12px 20px;
        font-size: 0.9rem;
        color: #85929E;
        border-top: 1px solid #EBEDEF;
    }
    
    /* 上傳區樣式 */
    [data-testid="stFileUploaderDropzone"] {
        background-color: #D6EAF8; border: 2px dashed #2E86C1; padding: 20px;
    }
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
    
    ws_types = sh_ref.worksheet("設備類型")
    ws_coef = sh_ref.worksheet("冷媒係數表")
    
    try: ws_records = sh_ref.worksheet("冷媒填報紀錄")
    except: 
        ws_records = sh_ref.add_worksheet(title="冷媒填報紀錄", rows="1000", cols="15")
        ws_records.append_row(["填報時間","填報人","填報人分機","校區","所屬單位","填報單位名稱","建築物名稱","辦公室編號","維修日期","設備類型","設備品牌型號","冷媒種類","冷媒填充量","備註","佐證資料"])

except Exception as e:
    st.error(f"❌ 資料庫連線失敗: {e}")
    st.stop()

# 5. 資料讀取 (選項與紀錄)
@st.cache_data(ttl=60)
def load_data_all():
    # 1. 設備類型選項
    type_data = ws_types.get_all_values()
    e_types = sorted([row[0] for row in type_data[1:] if row]) if len(type_data) > 1 else []
    
    # 2. 係數表 (建立 GWP 對照表)
    coef_data = ws_coef.get_all_values()
    r_types = []
    gwp_map = {}
    
    if len(coef_data) > 1:
        try:
            name_idx = 1
            gwp_idx = 2
            for row in coef_data[1:]:
                if len(row) > gwp_idx and row[name_idx]:
                    r_name = row[name_idx].strip()
                    gwp_val_str = row[gwp_idx].replace(',', '').strip()
                    if not gwp_val_str.replace('.', '', 1).isdigit():
                        gwp_val = 0.0
                    else:
                        gwp_val = float(gwp_val_str)
                    r_types.append(r_name)
                    gwp_map[r_name] = gwp_val
        except:
            r_types = sorted([row[1] for row in coef_data[1:] if len(row) > 1 and row[1]])

    # 3. 填報紀錄
    records_data = ws_records.get_all_values()
    if len(records_data) > 1:
        df_records = pd.DataFrame(records_data[1:], columns=records_data[0])
    else:
        df_records = pd.DataFrame(columns=["填報時間","填報人","填報人分機","校區","所屬單位","填報單位名稱","建築物名稱","辦公室編號","維修日期","設備類型","設備品牌型號","冷媒種類","冷媒填充量","備註","佐證資料"])

    return e_types, sorted(r_types), gwp_map, df_records

e_types, r_types, gwp_map, df_records = load_data_all()

# 6. 頁面介面
st.title("❄️ 冷媒填報專區")

tabs = st.tabs(["📝 新增填報", "📊 動態查詢看板"])

# ==========================================
# 分頁 1: 新增填報
# ==========================================
with tabs[0]:
    
    # === 區塊 1: 填報單位基本資訊區 ===
    st.markdown('<div class="morandi-header">填報單位基本資訊區</div>', unsafe_allow_html=True)
    
    c1, c2 = st.columns(2)
    
    unit_depts = sorted(UNIT_DATA.keys())
    sel_dept = c1.selectbox("所屬單位", unit_depts, index=None, placeholder="請選擇單位...")
    
    unit_names = []
    if sel_dept:
        unit_names = sorted(UNIT_DATA.get(sel_dept, []))
    sel_unit_name = c2.selectbox("填報單位名稱", unit_names, index=None, placeholder="請先選擇所屬單位...")
    
    c3, c4 = st.columns(2)
    name = c3.text_input("填報人")
    ext = c4.text_input("填報人分機")
    
    # === 區塊 2: 冷媒設備所在位置資訊區 ===
    st.markdown('<div class="morandi-header">冷媒設備所在位置資訊區</div>', unsafe_allow_html=True)
    
    loc_campuses = sorted(BUILDING_DATA.keys())
    sel_loc_campus = st.selectbox("填報單位所在校區", loc_campuses, index=None, placeholder="請選擇校區...")
    
    c6, c7 = st.columns(2)
    
    buildings = []
    if sel_loc_campus:
        buildings = sorted(BUILDING_DATA.get(sel_loc_campus, []))
    sel_build = c6.selectbox("建築物名稱", buildings, index=None, placeholder="請先選擇校區...")
    
    office = c7.text_input("辦公室編號", placeholder="例如：202辦公室、306研究室")
    
    # === 區塊 3: 冷媒設備填充資訊區 ===
    st.markdown('<div class="morandi-header">冷媒設備填充資訊區</div>', unsafe_allow_html=True)
    
    c8, c9 = st.columns(2)
    r_date = c8.date_input("維修日期 (統一填寫發票日期)", datetime.today())
    
    sel_etype = c9.selectbox("設備類型", e_types, index=None, placeholder="請選擇...")
    
    c10, c11 = st.columns(2)
    e_model = c10.text_input("設備品牌型號", placeholder="例如：國際 CS-100FL+CU-100FLC")
    
    sel_rtype = c11.selectbox("冷媒種類", r_types, index=None, placeholder="請選擇...")
    
    amount = st.number_input("冷媒填充量 (公斤)", min_value=0.0, step=0.1, format="%.2f")
    
    st.markdown("請上傳冷媒填充單據佐證資料")
    f_file = st.file_uploader("上傳佐證 (必填)", type=['pdf', 'jpg', 'png'], label_visibility="collapsed")
    
    st.markdown("---")
    note = st.text_input("備註內容", placeholder="備註 (選填)")
    st.markdown('<div class="correction-note">如有資料誤繕情形，請重新登錄1次資訊，並於備註欄填寫：「前筆資料誤繕，請刪除。」，管理單位將協助刪除誤打資訊</div>', unsafe_allow_html=True)
    
    # === 完整個資聲明 ===
    st.markdown("""
    <div class="privacy-box">
        <strong>📜 個人資料蒐集、處理及利用告知聲明</strong><br>
        1. 蒐集機關：國立嘉義大學。<br>
        2. 蒐集目的：進行本校冷媒設備之冷媒填充紀錄管理、校園溫室氣體（碳）盤查統計、稽核佐證資料蒐集及後續能源使用分析。<br>
        3. 個資類別：填報人姓名。<br>
        4. 利用期間：姓名保留至填報年度後第二年1月1日，期滿即進行「去識別化」刪除，其餘數據永久保存。<br>
        5. 利用對象：本校教師、行政人員及碳盤查查驗人員。<br>
        6. 您有權依個資法請求查詢、更正或刪除您的個資。如不提供，將無法完成填報。
    </div>
    """, unsafe_allow_html=True)
    
    agree = st.checkbox("我已閱讀並同意個資聲明，且確認所填資料無誤。")
    
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
                f_file.seek(0); f_ext = f_file.name.split('.')[-1]
                clean_name = f"{sel_loc_campus}_{sel_dept}_{sel_unit_name}_{r_date}_{sel_etype}_{sel_rtype}.{f_ext}"
                meta = {'name': clean_name, 'parents': [REF_FOLDER_ID]}
                media = MediaIoBaseUpload(f_file, mimetype=f_file.type, resumable=True)
                file = drive_service.files().create(body=meta, media_body=media, fields='webViewLink').execute()
                link = file.get('webViewLink')
                
                current_time = get_taiwan_time().strftime("%Y-%m-%d %H:%M:%S")
                
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

# ==========================================
# 分頁 2: 動態查詢看板 (整合式資訊卡版)
# ==========================================
with tabs[1]:
    st.markdown('<div class="morandi-header">📊 歷史填報紀錄查詢</div>', unsafe_allow_html=True)

    if df_records.empty:
        st.info("目前尚無填報紀錄。")
    else:
        # --- 1. 資料前處理 ---
        df_records['冷媒填充量'] = pd.to_numeric(df_records['冷媒填充量'], errors='coerce').fillna(0)
        df_records['維修日期'] = pd.to_datetime(df_records['維修日期'], errors='coerce')
        df_records['年份'] = df_records['維修日期'].dt.year.fillna(datetime.now().year).astype(int)
        
        # 計算排放量 (kgCO2e)
        def calc_emission(row):
            rtype = row.get('冷媒種類', '')
            amount = row.get('冷媒填充量', 0)
            gwp = gwp_map.get(rtype, 0)
            return amount * gwp

        df_records['排放量(kgCO2e)'] = df_records.apply(calc_emission, axis=1)

        # --- 2. 側邊欄篩選 ---
        st.markdown("##### 🔍 篩選條件")
        col_f1, col_f2, col_f3 = st.columns(3)
        
        years = sorted(df_records['年份'].unique(), reverse=True)
        sel_year = col_f1.selectbox("年份", ["全部"] + list(years))
        
        depts = sorted(df_records['所屬單位'].unique())
        sel_q_dept = col_f2.selectbox("所屬單位 (查詢)", ["全部"] + depts)
        
        units = []
        if sel_q_dept != "全部":
            units = sorted(df_records[df_records['所屬單位'] == sel_q_dept]['填報單位名稱'].unique())
        sel_q_unit = col_f3.selectbox("填報單位 (查詢)", ["全部"] + units)

        # 執行篩選
        df_view = df_records.copy()
        if sel_year != "全部":
            df_view = df_view[df_view['年份'] == sel_year]
        if sel_q_dept != "全部":
            df_view = df_view[df_view['所屬單位'] == sel_q_dept]
        if sel_q_unit != "全部":
            df_view = df_view[df_view['填報單位名稱'] == sel_q_unit]

        # --- 3. 整合式資訊卡 (V233 新增) ---
        st.markdown("---")
        
        if not df_view.empty:
            # 計算數據
            total_emission = df_view['排放量(kgCO2e)'].sum()
            
            # 標題名稱邏輯
            if sel_q_unit != "全部":
                card_title = sel_q_unit
            elif sel_q_dept != "全部":
                card_title = sel_q_dept
            else:
                card_title = "全校總計"
                
            # 填充資訊 (依種類加總)
            fill_summary = df_view.groupby('冷媒種類')['冷媒填充量'].sum().reset_index()
            
            # 申報履歷 (日期清單)
            dates_list = sorted(df_view['維修日期'].dt.strftime('%Y-%m-%d').unique(), reverse=True)
            dates_str = ", ".join(dates_list) if dates_list else "無"

            # 產生 HTML 卡片
            fill_rows_html = ""
            for _, row in fill_summary.iterrows():
                fill_rows_html += f"""
                <div class="fill-row">
                    <span class="fill-type">{row['冷媒種類']}</span>
                    <span class="fill-amount">{row['冷媒填充量']:.2f} kg</span>
                </div>
                """
            
            st.markdown(f"""
            <div class="info-card">
                <div class="card-header">
                    <div class="card-title">{card_title}</div>
                    <div style="text-align:right;">
                        <span style="font-size:0.9rem; color:#566573;">碳排放量</span><br>
                        <span class="emission-value">{total_emission:,.2f}</span>
                        <span class="emission-unit">kgCO2e</span>
                    </div>
                </div>
                <div class="card-body">
                    {fill_rows_html}
                </div>
                <div class="card-footer">
                    <strong>📅 歷次申報日期：</strong> {dates_str}
                </div>
            </div>
            """, unsafe_allow_html=True)
        
        else:
            st.warning("查無符合條件的資料。")

        # --- 4. 詳細資料列表 ---
        st.markdown("### 📋 詳細清單")
        show_cols = ["維修日期", "校區", "所屬單位", "填報單位名稱", "設備類型", "設備品牌型號", "冷媒種類", "冷媒填充量", "排放量(kgCO2e)", "佐證資料"]
        valid_cols = [c for c in show_cols if c in df_view.columns]
        
        st.dataframe(
            df_view[valid_cols].sort_values("維修日期", ascending=False),
            use_container_width=True,
            column_config={
                "維修日期": st.column_config.DateColumn("維修日期", format="YYYY-MM-DD"),
                "冷媒填充量": st.column_config.NumberColumn("填充量 (kg)", format="%.2f"),
                "排放量(kgCO2e)": st.column_config.NumberColumn("排放量 (kgCO2e)", format="%.2f"),
                "佐證資料": st.column_config.LinkColumn("佐證連結", display_text="開啟檔案")
            }
        )