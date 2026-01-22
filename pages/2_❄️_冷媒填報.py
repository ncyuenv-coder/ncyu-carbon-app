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
# 1. 內建靜態選單資料 (資料庫)
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
# 2. CSS 樣式
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
    
    # 只讀取選項比較固定的 Sheet，減少變數
    ws_types = sh_ref.worksheet("設備類型")
    ws_coef = sh_ref.worksheet("冷媒係數表")
    
    try: ws_records = sh_ref.worksheet("冷媒填報紀錄")
    except: 
        ws_records = sh_ref.add_worksheet(title="冷媒填報紀錄", rows="1000", cols="15")
        ws_records.append_row(["填報時間","填報人","填報人分機","校區","所屬單位","填報單位名稱","建築物名稱","辦公室編號","維修日期","設備類型","設備品牌型號","冷媒種類","冷媒填充量","備註","佐證資料"])

except Exception as e:
    st.error(f"❌ 資料庫連線失敗: {e}")
    st.stop()

# 5. 資料讀取 (讀取係數與類型)
@st.cache_data(ttl=600)
def load_static_options():
    # 設備類型
    type_data = ws_types.get_all_values()
    e_types = sorted([row[0] for row in type_data[1:] if row]) if len(type_data) > 1 else []
    
    # 冷媒係數
    coef_data = ws_coef.get_all_values()
    r_types = []
    if len(coef_data) > 1:
        # B欄是名稱 (Index 1)
        target_idx = 1 if len(coef_data[0]) > 1 else 0
        r_types = sorted([row[target_idx] for row in coef_data[1:] if len(row) > target_idx and row[target_idx]])
        
    return e_types, r_types

e_types, r_types = load_static_options()

# 6. 頁面介面 (注意：移除 st.form 以允許即時連動)
st.title("❄️ 冷媒填報專區")

tabs = st.tabs(["📝 新增填報", "📊 動態查詢看板"])

with tabs[0]:
    # ⚠️ 這裡移除了 with st.form(...)，讓選單可以即時更新
    
    # === 區塊 1: 填報人基本資訊區 ===
    st.markdown('<div class="section-header">1. 填報人基本資訊區</div>', unsafe_allow_html=True)
    
    c1, c2 = st.columns(2)
    
    # 1-1. 所屬單位
    unit_depts = sorted(UNIT_DATA.keys())
    sel_dept = c1.selectbox("所屬單位", unit_depts, index=None, placeholder="請選擇單位...")
    
    # 1-2. 填報單位名稱 (直接查字典，因無 Form 限制，此處會即時更新)
    unit_names = []
    if sel_dept:
        unit_names = sorted(UNIT_DATA.get(sel_dept, []))
    sel_unit_name = c2.selectbox("填報單位名稱", unit_names, index=None, placeholder="請先選擇所屬單位...")
    
    # 1-3. 開放欄位
    c3, c4 = st.columns(2)
    name = c3.text_input("填報人")
    ext = c4.text_input("填報人分機")
    
    st.markdown("---")
    
    # === 區塊 2: 詳細位置資訊區 ===
    st.markdown('<div class="section-header">2. 詳細位置資訊區</div>', unsafe_allow_html=True)
    c6, c7 = st.columns(2)
    
    # 2-1. 填報單位所在校區
    loc_campuses = sorted(BUILDING_DATA.keys())
    sel_loc_campus = c6.selectbox("填報單位所在校區", loc_campuses, index=None, placeholder="請選擇校區...")
    
    # 2-2. 建築物名稱
    buildings = []
    if sel_loc_campus:
        buildings = sorted(BUILDING_DATA.get(sel_loc_campus, []))
    sel_build = c6.selectbox("建築物名稱", buildings, index=None, placeholder="請先選擇校區...")
    
    # 2-3. 辦公室
    office = c7.text_input("辦公室編號", placeholder="例如：202辦公室、306研究室")
    
    st.markdown("---")
    
    # === 區塊 3: 設備修繕資訊 ===
    st.markdown('<div class="section-header">3. 設備修繕冷媒填充資訊區</div>', unsafe_allow_html=True)
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
    
    st.markdown('<div style="background-color:#F8F9F9; padding:10px; font-size:0.9rem;"><strong>📜 個資聲明</strong>：蒐集目的為設備管理與碳盤查，保存至申報後第二年。</div>', unsafe_allow_html=True)
    agree = st.checkbox("我已閱讀並同意個資聲明")
    
    # 改用一般按鈕，因為我們不在 form 裡面了
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
                # 檔名邏輯
                clean_name = f"{sel_loc_campus}_{sel_dept}_{sel_unit_name}_{r_date}_{sel_etype}_{sel_rtype}.{f_ext}"
                meta = {'name': clean_name, 'parents': [REF_FOLDER_ID]}
                media = MediaIoBaseUpload(f_file, mimetype=f_file.type, resumable=True)
                file = drive_service.files().create(body=meta, media_body=media, fields='webViewLink').execute()
                link = file.get('webViewLink')
                
                current_time = get_taiwan_time().strftime("%Y-%m-%d %H:%M:%S")
                
                # 寫入資料庫
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