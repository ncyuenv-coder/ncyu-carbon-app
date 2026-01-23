import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import gspread
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import re

# ==========================================
# 0. 系統設定
# ==========================================
st.set_page_config(page_title="冷媒填報 - 嘉義大學", page_icon="❄️", layout="wide")

def get_taiwan_time():
    return datetime.utcnow() + timedelta(hours=8)

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
    
    /* 個資聲明勾選文字樣式 (深藍、粗體、加大 - 比照燃油) */
    [data-testid="stCheckbox"] label p {
        font-size: 1.15rem !important;
        font-weight: 700 !important;
        color: #2E4053 !important;
    }

    /* 4. 橫式資訊卡 (V239 Update) */
    .horizontal-card {
        display: flex;
        border: 1px solid #BDC3C7;
        border-radius: 12px;
        overflow: hidden;
        margin-bottom: 25px;
        box-shadow: 0 4px 8px rgba(0,0,0,0.08);
        background-color: #FFFFFF;
        min-height: 250px;
    }
    
    /* 左側 30% */
    .card-left {
        flex: 3;
        background-color: #34495E;
        color: #FFFFFF;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        padding: 20px;
        text-align: center;
        border-right: 1px solid #2C3E50;
    }
    .dept-text { font-size: 1.5rem; font-weight: 700; margin-bottom: 8px; }
    .unit-text { font-size: 1.2rem; font-weight: 500; opacity: 0.9; }
    
    /* 右側 70% */
    .card-right {
        flex: 7;
        padding: 20px 30px;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }
    .info-row {
        display: flex;
        align-items: flex-start; /* 改為靠上對齊，適應多行內容 */
        padding: 10px 0;
        font-size: 1rem;
        color: #566573;
        border-bottom: 1px dashed #F2F3F4;
    }
    .info-row:last-child { border-bottom: none; }
    .info-icon { margin-right: 12px; font-size: 1.1rem; width: 25px; text-align: center; margin-top: 2px; }
    .info-label { font-weight: 600; margin-right: 10px; min-width: 150px; color: #2E4053; }
    .info-value { font-weight: 500; color: #17202A; flex: 1; }
    
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
    
    # 讀取必要分頁
    ws_units = sh_ref.worksheet("單位資訊") 
    ws_buildings = sh_ref.worksheet("建築物清單") # V239: 新增讀取建築物清單
    ws_types = sh_ref.worksheet("設備類型")
    ws_coef = sh_ref.worksheet("冷媒係數表")
    
    try: ws_records = sh_ref.worksheet("冷媒填報紀錄")
    except: 
        ws_records = sh_ref.add_worksheet(title="冷媒填報紀錄", rows="1000", cols="15")
        ws_records.append_row(["填報時間","填報人","填報人分機","校區","所屬單位","填報單位名稱","建築物名稱","辦公室編號","維修日期","設備類型","設備品牌型號","冷媒種類","冷媒填充量","備註","佐證資料"])

except Exception as e:
    st.error(f"❌ 資料庫連線失敗: {e}")
    st.stop()

# 5. 資料讀取 (選項與紀錄 - V239 完全動態版)
@st.cache_data(ttl=60)
def load_data_all():
    # 1. 單位資訊 (動態讀取)
    unit_data = ws_units.get_all_values()
    unit_dict = {}
    if len(unit_data) > 1:
        for row in unit_data[1:]:
            if len(row) >= 2:
                dept = str(row[0]).strip()
                unit = str(row[1]).strip()
                if dept and unit:
                    if dept not in unit_dict:
                        unit_dict[dept] = []
                    if unit not in unit_dict[dept]:
                        unit_dict[dept].append(unit)
    
    # 2. 建築物清單 (動態讀取 V239)
    building_data = ws_buildings.get_all_values()
    build_dict = {}
    if len(building_data) > 1:
        for row in building_data[1:]:
            if len(row) >= 2:
                campus = str(row[0]).strip()
                b_name = str(row[1]).strip()
                if campus and b_name:
                    if campus not in build_dict:
                        build_dict[campus] = []
                    if b_name not in build_dict[campus]:
                        build_dict[campus].append(b_name)

    # 3. 設備類型選項
    type_data = ws_types.get_all_values()
    e_types = sorted([row[0] for row in type_data[1:] if row]) if len(type_data) > 1 else []
    
    # 4. 係數表 (建立 GWP 對照表)
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

    # 5. 填報紀錄
    records_data = ws_records.get_all_values()
    if len(records_data) > 1:
        raw_headers = records_data[0]
        col_mapping = {}
        for h in raw_headers:
            clean_h = str(h).strip()
            if "填充量" in clean_h or "重量" in clean_h:
                col_mapping[h] = "冷媒填充量"
            elif "種類" in clean_h or "品項" in clean_h:
                col_mapping[h] = "冷媒種類"
            elif "日期" in clean_h or "維修" in clean_h:
                col_mapping[h] = "維修日期"
            else:
                col_mapping[h] = clean_h
        
        df_records = pd.DataFrame(records_data[1:], columns=raw_headers)
        df_records.rename(columns=col_mapping, inplace=True)
    else:
        df_records = pd.DataFrame(columns=["填報時間","填報人","填報人分機","校區","所屬單位","填報單位名稱","建築物名稱","辦公室編號","維修日期","設備類型","設備品牌型號","冷媒種類","冷媒填充量","備註","佐證資料"])

    return unit_dict, build_dict, e_types, sorted(r_types), gwp_map, df_records

# 呼叫載入函式
unit_dict, build_dict, e_types, r_types, gwp_map, df_records = load_data_all()

# 6. 頁面介面
st.title("❄️ 冷媒填報專區")

tabs = st.tabs(["📝 新增填報", "📋 申報動態查詢"])

# ==========================================
# 分頁 1: 新增填報
# ==========================================
with tabs[0]:
    
    st.markdown('<div class="morandi-header">填報單位基本資訊區</div>', unsafe_allow_html=True)
    
    c1, c2 = st.columns(2)
    
    # 使用動態讀取的 unit_dict
    unit_depts = sorted(unit_dict.keys())
    sel_dept = c1.selectbox("所屬單位", unit_depts, index=None, placeholder="請選擇單位...")
    
    unit_names = []
    if sel_dept:
        unit_names = sorted(unit_dict.get(sel_dept, []))
    sel_unit_name = c2.selectbox("填報單位名稱", unit_names, index=None, placeholder="請先選擇所屬單位...")
    
    c3, c4 = st.columns(2)
    name = c3.text_input("填報人")
    ext = c4.text_input("填報人分機")
    
    st.markdown('<div class="morandi-header">冷媒設備所在位置資訊區</div>', unsafe_allow_html=True)
    
    # 使用動態讀取的 build_dict
    loc_campuses = sorted(build_dict.keys())
    sel_loc_campus = st.selectbox("填報單位所在校區", loc_campuses, index=None, placeholder="請選擇校區...")
    
    c6, c7 = st.columns(2)
    
    buildings = []
    if sel_loc_campus:
        buildings = sorted(build_dict.get(sel_loc_campus, []))
    sel_build = c6.selectbox("建築物名稱", buildings, index=None, placeholder="請先選擇校區...")
    
    office = c7.text_input("辦公室編號", placeholder="例如：202辦公室、306研究室")
    
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
# 分頁 2: 申報動態查詢 (V238 完整功能版)
# ==========================================
with tabs[1]:
    st.markdown('<div class="morandi-header">📋 申報動態查詢</div>', unsafe_allow_html=True)

    if df_records.empty:
        st.info("目前尚無填報紀錄。")
    else:
        # --- 1. 資料前處理 ---
        if '冷媒填充量' not in df_records.columns or '維修日期' not in df_records.columns:
            st.error(f"❌ 關鍵欄位遺失，請檢查資料庫。目前欄位: {list(df_records.columns)}")
        else:
            df_records['冷媒填充量'] = pd.to_numeric(df_records['冷媒填充量'], errors='coerce').fillna(0)
            df_records['維修日期'] = pd.to_datetime(df_records['維修日期'], errors='coerce')
            
            def calc_emission(row):
                rtype = row.get('冷媒種類', '')
                amount = row.get('冷媒填充量', 0)
                gwp = gwp_map.get(rtype, 0)
                return amount * gwp

            df_records['排放量(kgCO2e)'] = df_records.apply(calc_emission, axis=1)

            # --- 2. 篩選區塊 ---
            st.markdown("##### 🔍 查詢條件設定")
            
            c_f1, c_f2 = st.columns(2)
            
            depts = sorted(df_records['所屬單位'].dropna().unique())
            sel_q_dept = c_f1.selectbox("所屬單位 (必選)", depts, index=None, placeholder="請選擇...")
            
            units = []
            if sel_q_dept:
                units = sorted(df_records[df_records['所屬單位'] == sel_q_dept]['填報單位名稱'].dropna().unique())
            sel_q_unit = c_f2.selectbox("填報單位名稱 (必選)", units, index=None, placeholder="請選擇...")
            
            c_f3, c_f4 = st.columns(2)
            today = datetime.now().date()
            start_of_year = date(today.year, 1, 1)
            
            q_start_date = c_f3.date_input("查詢起始日期", value=start_of_year, max_value=today)
            q_end_date = c_f4.date_input("查詢結束日期", value=today, max_value=today)

            # --- 3. 執行篩選與顯示 ---
            if sel_q_dept and sel_q_unit and q_start_date and q_end_date:
                if q_start_date > q_end_date:
                    st.warning("⚠️ 起始日期不能晚於結束日期")
                else:
                    start_dt = pd.Timestamp(q_start_date)
                    end_dt = pd.Timestamp(q_end_date)
                    
                    mask = (
                        (df_records['所屬單位'] == sel_q_dept) &
                        (df_records['填報單位名稱'] == sel_q_unit) &
                        (df_records['維修日期'] >= start_dt) &
                        (df_records['維修日期'] <= end_dt)
                    )
                    df_view = df_records[mask]
                    
                    # --- 4. 準備顯示資訊 ---
                    if sel_q_dept == sel_q_unit:
                        left_html = f'<div class="dept-text">{sel_q_dept}</div>'
                    else:
                        left_html = f'<div class="dept-text">{sel_q_dept}</div><div class="unit-text">{sel_q_unit}</div>'
                    
                    total_count = len(df_view)
                    if total_count > 0:
                        campus_str = ", ".join(sorted(df_view['校區'].unique()))
                        builds = sorted(df_view['建築物名稱'].unique())
                        build_str = ", ".join(builds[:3]) + (f" 等{len(builds)}棟" if len(builds)>3 else "")
                        
                        # 新增: 設備類型
                        equip_str = ", ".join(sorted(df_view['設備類型'].unique()))
                        
                        # 新增: 冷媒填充資訊 (逐筆明細)
                        fill_details = []
                        for _, row in df_view.iterrows():
                            fill_details.append(f"<div>• {row['冷媒種類']}：{row['冷媒填充量']:.2f} 公斤</div>")
                        fill_detail_html = "".join(fill_details)
                        
                        # 新增: 冷媒填充重量統計 (分類加總)
                        fill_summary = df_view.groupby('冷媒種類')['冷媒填充量'].sum().reset_index()
                        fill_stats = []
                        for _, row in fill_summary.iterrows():
                            fill_stats.append(f"<div>• {row['冷媒種類']}：{row['冷媒填充量']:.2f} 公斤</div>")
                        fill_stats_html = "".join(fill_stats)
                        
                        total_emission = df_view['排放量(kgCO2e)'].sum()
                    else:
                        campus_str = "無資料"
                        build_str = "無資料"
                        equip_str = "無資料"
                        fill_detail_html = "無資料"
                        fill_stats_html = "無資料"
                        total_emission = 0.0

                    # --- 5. 渲染橫式資訊卡 ---
                    st.markdown("---")
                    st.markdown(f"""
                    <div class="horizontal-card">
                        <div class="card-left">
                            {left_html}
                        </div>
                        <div class="card-right">
                            <div class="info-row">
                                <span class="info-icon">📅</span>
                                <span class="info-label">查詢起訖時間區間</span>
                                <span class="info-value">{q_start_date} ~ {q_end_date}</span>
                            </div>
                            <div class="info-row">
                                <span class="info-icon">🏫</span>
                                <span class="info-label">所在校區</span>
                                <span class="info-value">{campus_str}</span>
                            </div>
                            <div class="info-row">
                                <span class="info-icon">🏢</span>
                                <span class="info-label">建築物名稱</span>
                                <span class="info-value">{build_str}</span>
                            </div>
                            <div class="info-row">
                                <span class="info-icon">❄️</span>
                                <span class="info-label">設備類型</span>
                                <span class="info-value">{equip_str}</span>
                            </div>
                            <div class="info-row">
                                <span class="info-icon">📝</span>
                                <span class="info-label">冷媒填充資訊</span>
                                <span class="info-value">{fill_detail_html}</span>
                            </div>
                            <div class="info-row">
                                <span class="info-icon">⚖️</span>
                                <span class="info-label">冷媒填充重量統計</span>
                                <span class="info-value">{fill_stats_html}</span>
                            </div>
                            <div class="info-row">
                                <span class="info-icon">🌍</span>
                                <span class="info-label">碳排放量</span>
                                <span class="info-value" style="color:#C0392B;">
                                    <span style="font-size: 1.8rem; font-weight:900;">{total_emission:,.2f}</span> 
                                    <span style="font-size: 1rem; font-weight:normal;">公斤二氧化碳當量</span>
                                </span>
                            </div>
                            <div class="info-row">
                                <span class="info-icon">📊</span>
                                <span class="info-label">申報次數統計</span>
                                <span class="info-value">{total_count} 次</span>
                            </div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

                    # --- 6. 詳細填報明細 ---
                    st.markdown(f"##### 📋 {sel_q_dept} {sel_q_unit} 填報明細")
                    show_cols = ["維修日期", "校區", "建築物名稱", "設備類型", "設備品牌型號", "冷媒種類", "冷媒填充量", "排放量(kgCO2e)", "佐證資料"]
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
                
            else:
                if not (sel_q_dept and sel_q_unit):
                    st.info("👈 請先選擇「所屬單位」與「填報單位名稱」以開始查詢。")