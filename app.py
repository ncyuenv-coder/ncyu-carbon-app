import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import gspread
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import streamlit_authenticator as stauth
import plotly.express as px
import time

# ==========================================
# 0. 系統設定
# ==========================================
st.set_page_config(page_title="國立嘉義大學碳盤查平台", page_icon="🌍", layout="wide")

st.markdown("""
<style>
    html, body, [class*="css"] { font-family: "Microsoft JhengHei", sans-serif; }
    .login-header { font-size: 2.5rem; font-weight: 700; color: #2E4053; text-align: center; margin-bottom: 20px; padding: 20px; background-color: #F4F6F6; border-radius: 15px; }
    button[data-baseweb="tab"] { font-size: 1.5rem !important; font-weight: bold !important; padding: 1rem 2rem !important; }
    .stSelectbox label, .stTextInput label, .stNumberInput label, .stDateInput label { font-size: 1.2rem !important; color: #1B4F72 !important; font-weight: bold; }
    
    div[data-testid="stForm"] { 
        background-color: #E8F6F3; 
        padding: 30px; 
        border-radius: 20px; 
        border: 2px solid #A3E4D7;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    
    /* 強制設定輸入框背景為白，文字為黑 */
    div[data-baseweb="input"] > div, 
    div[data-baseweb="select"] > div, 
    div[data-baseweb="calendar"],
    textarea, 
    input {
        background-color: #FFFFFF !important;
        border-color: #D5DBDB !important;
        color: #000000 !important;
        -webkit-text-fill-color: #000000 !important;
    }
    
    ul[data-baseweb="menu"] li { color: #000000 !important; }
    div[data-baseweb="select"] span { color: #000000 !important; }
    
    .info-card { background-color: #FEF9E7; padding: 15px; border-left: 5px solid #F4D03F; border-radius: 5px; margin-bottom: 10px; font-size: 1.1rem; }
    .info-label { font-weight: bold; color: #7F8C8D; }
    .info-value { color: #212F3D; font-weight: 600; margin-left: 10px; }
    
    .contact-footer {
        text-align: center;
        margin-top: 50px;
        padding: 20px;
        background-color: #F8F9F9;
        border-top: 1px solid #D5DBDB;
        color: #566573;
        font-weight: bold;
    }

    /* V32.0 莫蘭迪 KPI 卡片 */
    .kpi-container {
        display: flex;
        justify-content: space-between;
        gap: 20px;
        margin-bottom: 20px;
    }
    .kpi-card {
        flex: 1;
        padding: 20px;
        border-radius: 15px;
        text-align: center;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        color: #2c3e50;
    }
    .kpi-card-total { background-color: #D6EAF8; border-left: 8px solid #5DADE2; }
    .kpi-card-gas { background-color: #D5F5E3; border-left: 8px solid #58D68D; }
    .kpi-card-diesel { background-color: #FCF3CF; border-left: 8px solid #F4D03F; }
    
    .kpi-title { font-size: 1.2rem; font-weight: bold; margin-bottom: 10px; opacity: 0.8; }
    .kpi-value { font-size: 3rem; font-weight: 800; line-height: 1.2; }
    .kpi-unit { font-size: 1rem; font-weight: normal; opacity: 0.7; }

    /* 👇 V33.0 新增：個資聲明區塊樣式 */
    .privacy-box {
        background-color: #EBF5FB; /* 淺藍底 */
        border: 1px solid #AED6F1;
        border-left: 5px solid #2874A6; /* 深藍邊 */
        padding: 20px;
        border-radius: 10px;
        margin-bottom: 20px;
        font-size: 0.95rem;
        color: #2E4053;
        line-height: 1.6;
    }
    .privacy-title {
        font-weight: bold;
        font-size: 1.2rem;
        color: #1B4F72;
        margin-bottom: 10px;
    }
</style>
""", unsafe_allow_html=True)

# ☁️ 設定區 (⚠️ 請務必填回學校的 ID)
SHEET_ID = "1gqDU21YJeBoBOd8rMYzwwZ45offXWPGEODKTF6B8k-Y" 
DRIVE_FOLDER_ID = "1DCmR0dXOdFBdTrgnvCYFPtNq_bGzSJeB" 

# ==========================================
# 1. 安全登入模組
# ==========================================
def clean_secrets(obj):
    if isinstance(obj, dict) or "AttrDict" in str(type(obj)):
        return {k: clean_secrets(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [clean_secrets(i) for i in obj]
    return obj

if 'current_page' not in st.session_state:
    st.session_state['current_page'] = 'home'

if 'reset_counter' not in st.session_state:
    st.session_state['reset_counter'] = 0

try:
    _raw_creds = st.secrets["credentials"]
    credentials_login = clean_secrets(_raw_creds)
    cookie_cfg = st.secrets["cookie"]
    
    authenticator = stauth.Authenticate(
        credentials_login,
        cookie_cfg["name"],
        cookie_cfg["key"],
        cookie_cfg["expiry_days"],
    )
    
    if st.session_state["authentication_status"] is not True:
        st.markdown('<div class="login-header">🏫 國立嘉義大學碳盤查<br>油料使用及冷媒填充回報平台</div>', unsafe_allow_html=True)
        st.markdown("---")
    
    authenticator.login('main')
    
    if st.session_state["authentication_status"] is False:
        st.error('❌ 帳號或密碼錯誤')
        st.stop()
    elif st.session_state["authentication_status"] is None:
        st.info('🔒 請輸入帳號密碼登入系統')
        st.stop()
        
    name = st.session_state["name"]
    username = st.session_state["username"]
    
    with st.sidebar:
        st.header(f"👤 {name}")
        st.success("☁️ 雲端連線正常")
        if st.button("🏠 返回主選單"):
            st.session_state['current_page'] = 'home'
            st.rerun()
        st.markdown("---")
        authenticator.logout('登出系統', 'sidebar')

except Exception as e:
    st.error(f"登入錯誤: {e}")
    st.stop()

# ==========================================
# 2. 雲端連線 (OAuth)
# ==========================================
@st.cache_resource
def init_google():
    oauth_info = st.secrets["gcp_oauth"]
    creds = Credentials(
        token=None, 
        refresh_token=oauth_info["refresh_token"],
        token_uri="https://oauth2.googleapis.com/token",
        client_id=oauth_info["client_id"],
        client_secret=oauth_info["client_secret"],
        scopes=["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/spreadsheets"]
    )
    gc = gspread.authorize(creds)
    drive_service = build('drive', 'v3', credentials=creds)
    return gc, drive_service

try:
    gc, drive_service = init_google()
    sh = gc.open_by_key(SHEET_ID)
    
    try: ws_equip = sh.worksheet("設備清單") 
    except: ws_equip = sh.sheet1 
    
    try: ws_record = sh.worksheet("填報紀錄")
    except: 
        ws_record = sh.add_worksheet(title="填報紀錄", rows="1000", cols="11")
        
    if len(ws_record.get_all_values()) == 0:
        ws_record.append_row([
            "填報時間", "填報單位", "填報人", "填報分機", 
            "設備名稱", "校內財產編號", "油料種類", 
            "加油日期", "加油量", "佐證資料", "單據備註"
        ])

except Exception as e:
    st.error(f"連線失敗: {e}")
    st.stop()

def load_data():
    df_e = pd.DataFrame(ws_equip.get_all_records())
    df_e = df_e.astype(str)
    df_r = pd.DataFrame(ws_record.get_all_records())
    return df_e, df_r

df_equip, df_records = load_data()

# ==========================================
# 3. 頁面邏輯
# ==========================================
if st.session_state['current_page'] == 'home':
    st.title("🏫 國立嘉義大學碳盤查回報平台")
    st.markdown("### 請選擇填報項目：")
    col1, col2 = st.columns(2)
    with col1:
        st.info("⛽ 車輛/機具用油")
        if st.button("前往「燃油設備填報區」", use_container_width=True, type="primary"):
            st.session_state['current_page'] = 'fuel'
            st.rerun()
    with col2:
        st.info("❄️ 冷氣/冰水主機")
        st.button("前往「冷媒類設備填報區」 (建置中)", use_container_width=True, disabled=True)
    
    st.markdown("""
        <div class="contact-footer">
        如有填報疑問，請電洽環安中心林小姐(分機 7137)，謝謝
        </div>
    """, unsafe_allow_html=True)

elif st.session_state['current_page'] == 'fuel':
    st.title("⛽ 燃油設備填報專區")
    
    # 👇 V33.0: 頁面精簡化，刪除舊看板，只保留 2 個主要分頁 (+Admin)
    tabs_list = ["📝 新增填報", "📊 動態查詢看板"]
    if username == 'admin':
        tabs_list.append("🛠️ 資料庫管理")
        
    tabs = st.tabs(tabs_list)

    # --- Tab 1: 填報 ---
    with tabs[0]:
        # 👇 V33.0: 新增「個資聲明」區塊 (使用 HTML 渲染樣式)
        st.markdown("""
        <div class="privacy-box">
            <div class="privacy-title">📜 個人資料蒐集、處理及利用告知聲明</div>
            1. <strong>蒐集機關</strong>：國立嘉義大學。<br>
            2. <strong>蒐集目的</strong>：進行本校公務車輛/機具之加油紀錄管理、校園溫室氣體（碳）盤查統計、稽核佐證資料蒐集及後續能源使用分析。<br>
            3. <strong>個資類別</strong>：填報人姓名。<br>
            4. <strong>利用期間</strong>：姓名保留至填報年度後第二年1月1日，期滿即進行「去識別化」刪除，其餘數據永久保存。<br>
            5. <strong>利用對象</strong>：本校教師、行政人員及碳盤查查驗人員。<br>
            6. <strong>您的權利</strong>：您有權依個資法請求查詢、更正或刪除您的個資。如不提供，將無法完成填報。<br>
        </div>
        """, unsafe_allow_html=True)

        if not df_equip.empty:
            st.markdown("#### 步驟 1：選擇設備")
            c1, c2 = st.columns(2)
            units = sorted([x for x in df_equip['填報單位'].unique() if x != '-' and x != '填報單位'])
            
            selected_dept = c1.selectbox(
                "填報單位", 
                units, 
                index=None, 
                placeholder="請選擇單位...", 
                key="dept_selector"
            )
            
            if selected_dept:
                filtered = df_equip[df_equip['填報單位'] == selected_dept]
                devices = sorted([x for x in filtered['設備名稱備註'].unique()])
                
                dynamic_key = f"vehicle_selector_{st.session_state['reset_counter']}"
                
                selected_device = c2.selectbox(
                    "車輛/機具名稱", 
                    devices, 
                    index=None, 
                    placeholder="請選擇車輛...", 
                    key=dynamic_key
                )
                
                if selected_device:
                    row = filtered[filtered['設備名稱備註'] == selected_device].iloc[0]
                    info_html = f"""
                    <div class="info-card">
                        <div><span class="info-label">🏢 部門：</span><span class="info-value">{row.get('設備所屬單位/部門', '-')}</span></div>
                        <div><span class="info-label">👤 保管人：</span><span class="info-value">{row.get('保管人', '-')}</span></div>
                        <div><span class="info-label">🔢 財產編號：</span><span class="info-value">{row.get('校內財產編號', '-')}</span></div>
                        <div><span class="info-label">📍 位置：</span><span class="info-value">{row.get('設備詳細位置/樓層', '-')}</span></div>
                        <div><span class="info-label">⛽ 燃料：</span><span class="info-value">{row.get('原燃物料名稱', '-')}</span></div>
                        <div><span class="info-label">📊 數量：</span><span class="info-value">{row.get('設備數量', '-')}</span></div>
                    </div>
                    """
                    st.markdown(info_html, unsafe_allow_html=True)
                    
                    st.markdown("#### 步驟 2：填寫資料")
                    with st.form("entry_form", clear_on_submit=True):
                        # 👇 V33.0: 放入個資同意勾選 (強制必選)
                        agree_privacy = st.checkbox("我已閱讀並同意上述聲明，且確認所填資料無誤。", value=False)
                        st.markdown("---")

                        col_p1, col_p2 = st.columns(2)
                        p_name = col_p1.text_input("👤 填報人姓名 (必填)")
                        p_ext = col_p2.text_input("📞 聯絡分機 (必填)")
                        
                        st.divider()
                        
                        is_unused = st.checkbox("🚫 本期無使用 (勾選後免填油量及上傳)", help="若此設備這段時間都沒有加油，請勾選此項。")
                        
                        col_a, col_b = st.columns(2)
                        
                        if is_unused:
                            d_date = col_a.date_input("📅 填報日期", datetime.today())
                            d_vol = col_b.number_input("💧 加油量 (本期無使用)", value=0.0, disabled=True)
                        else:
                            d_date = col_a.date_input("📅 加油日期 (以加油單為準)", datetime.today())
                            d_vol = col_b.number_input("💧 加油量 (公升)", min_value=0.0, step=0.1, format="%.1f")
                        
                        st.markdown("**🧾 單據備註 (選填)**")
                        note_input = st.text_input("若一張發票加多台設備，請填寫相同發票號碼以便核對")
                        st.caption("ℹ️ 如有資料誤繕情況，請重新新增1筆資料，並於備註欄註記「前一筆資料填錯，請刪除」，以利管理單位後端處理，謝謝。")

                        st.markdown("---")
                        
                        if is_unused:
                            st.markdown("**📂 上傳佐證資料 (本期無使用免附)**")
                            is_shared = False
                            f_files = None
                        else:
                            st.markdown("**📂 上傳佐證資料 (必填)**")
                            is_shared = st.checkbox("與其他設備共用加油單")
                            f_files = st.file_uploader("支援 png, jpg, pdf (最多 3 個，單檔限 10MB)", type=['png', 'jpg', 'jpeg', 'pdf'], accept_multiple_files=True)
                        
                        submitted = st.form_submit_button("🚀 確認送出資料", type="primary", use_container_width=True)
                        
                        if submitted:
                            # 1. 檢查個資同意
                            if not agree_privacy:
                                st.error("❌ 請務必勾選「我已閱讀並同意上述聲明」，才能送出資料！")
                            # 2. 必填檢查
                            elif not p_name or not p_ext:
                                st.warning("⚠️ 「填報人姓名」與「聯絡分機」為必填欄位！")
                            elif not is_unused and d_vol <= 0:
                                st.warning("⚠️ 加油量不能為 0 (若無使用請勾選上方「本期無使用」)")
                            elif not is_unused and not f_files:
                                st.error("⚠️ 請務必上傳佐證資料 (加油單據)")
                            else:
                                valid_files = True
                                if f_files:
                                    if len(f_files) > 3:
                                        st.error("❌ 超過檔案數量上限 (最多 3 個)")
                                        valid_files = False
                                    for f in f_files:
                                        if f.size > 10 * 1024 * 1024:
                                            st.error(f"❌ 檔案 {f.name} 太大 (超過 10MB)")
                                            valid_files = False
                                
                                if valid_files:
                                    progress_text = "資料處理中..."
                                    my_bar = st.progress(0, text=progress_text)
                                    
                                    file_links = []
                                    if f_files:
                                        for idx, f_file in enumerate(f_files):
                                            try:
                                                file_ext = f_file.name.split('.')[-1]
                                                fuel_name = row.get('原燃物料名稱', '未知燃料')
                                                shared_tag = "(共用)" if is_shared else ""
                                                
                                                clean_name = f"{selected_dept}_{selected_device}_{d_date}_{fuel_name}{d_vol}公升{shared_tag}_{idx+1}.{file_ext}".replace("/", "_")
                                                
                                                file_meta = {'name': clean_name, 'parents': [DRIVE_FOLDER_ID]}
                                                media = MediaIoBaseUpload(f_file, mimetype=f_file.type)
                                                file = drive_service.files().create(
                                                    body=file_meta, 
                                                    media_body=media, 
                                                    fields='webViewLink'
                                                ).execute()
                                                file_links.append(file.get('webViewLink'))
                                            except Exception as e:
                                                st.warning(f"檔案 {f_file.name} 上傳異常: {e}")
                                    
                                    final_links = "\n".join(file_links) if file_links else "無"

                                    final_note = note_input
                                    if is_unused:
                                        prefix = "【本期無使用】"
                                        final_note = f"{prefix} {note_input}" if note_input else prefix

                                    my_bar.progress(50, text="寫入資料庫...")
                                    
                                    ws_record.append_row([
                                        datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 
                                        selected_dept,                      
                                        p_name,                             
                                        p_ext,                              
                                        selected_device,                    
                                        str(row.get('校內財產編號', '-')),    
                                        str(row.get('原燃物料名稱', '-')),    
                                        str(d_date),                        
                                        d_vol,                              
                                        final_links,                        
                                        final_note                          
                                    ])
                                    
                                    my_bar.progress(100, text="完成！")
                                    time.sleep(0.5)
                                    my_bar.empty()
                                    if is_unused:
                                        st.success(f"✅ 已回報：{selected_device} 本期無使用")
                                    else:
                                        st.success(f"✅ 成功！已新增紀錄：{d_vol} L")
                                    st.balloons()
                                    
                                    st.session_state['reset_counter'] += 1
                                    st.rerun()
        
        st.markdown("""
            <div class="contact-footer">
            如有填報疑問，請電洽環安中心林小姐(分機 7137)，謝謝
            </div>
        """, unsafe_allow_html=True)

    # --- Tab 2: 動態查詢看板 (原單位查核) ---
    with tabs[1]:
        st.markdown("### 📊 動態查詢看板")
        st.info("請透過下方篩選器，檢視單位的用油統計與詳細紀錄。")
        
        if not df_records.empty and '加油量' in df_records.columns:
            # 1. 篩選器區域
            record_units = sorted([str(x) for x in df_records['填報單位'].unique() if str(x) != 'nan'])
            
            c_dept, c_date = st.columns([1, 2])
            query_dept = c_dept.selectbox("🏢 選擇查詢單位", record_units, index=None, placeholder="請選擇...")
            
            today = date.today()
            start_of_year = date(today.year, 1, 1)
            query_date_range = c_date.date_input(
                "📅 選擇查詢日期區間",
                (start_of_year, today),
                format="YYYY/MM/DD"
            )
            
            if query_dept and isinstance(query_date_range, tuple) and len(query_date_range) == 2:
                start_date, end_date = query_date_range
                
                df_dept = df_records[df_records['填報單位'] == query_dept].copy()
                df_dept['加油日期'] = pd.to_datetime(df_dept['加油日期'], errors='coerce').dt.date
                
                mask = (df_dept['加油日期'] >= start_date) & (df_dept['加油日期'] <= end_date)
                df_final = df_dept.loc[mask]
                
                if not df_final.empty:
                    # 2. KPI Cards
                    if '油料種類' in df_final.columns:
                        df_final['油料種類'] = df_final['油料種類'].astype(str)
                        gasoline_sum = df_final[df_final['油料種類'].str.contains('汽油')]['加油量'].sum()
                        diesel_sum = df_final[df_final['油料種類'].str.contains('柴油')]['加油量'].sum()
                    else:
                        gasoline_sum = 0
                        diesel_sum = 0
                    
                    total_sum = df_final['加油量'].sum()
                    
                    st.markdown("---")
                    kpi_html = f"""
                    <div class="kpi-container">
                        <div class="kpi-card kpi-card-gas">
                            <div class="kpi-title">⛽ 汽油使用量</div>
                            <div class="kpi-value">{gasoline_sum:,.1f}<span class="kpi-unit"> L</span></div>
                        </div>
                        <div class="kpi-card kpi-card-diesel">
                            <div class="kpi-title">🚛 柴油使用量</div>
                            <div class="kpi-value">{diesel_sum:,.1f}<span class="kpi-unit"> L</span></div>
                        </div>
                        <div class="kpi-card kpi-card-total">
                            <div class="kpi-title">💧 總用油量</div>
                            <div class="kpi-value">{total_sum:,.1f}<span class="kpi-unit"> L</span></div>
                        </div>
                    </div>
                    """
                    st.markdown(kpi_html, unsafe_allow_html=True)
                    
                    # 3. 趨勢圖
                    st.subheader(f"📊 每月加油趨勢分析 (依設備堆疊)")
                    df_final['月份'] = pd.to_datetime(df_final['加油日期']).dt.strftime('%Y-%m')
                    chart_data = df_final.groupby(['月份', '設備名稱'])['加油量'].sum().reset_index()
                    
                    fig = px.bar(
                        chart_data, 
                        x='月份', 
                        y='加油量', 
                        color='設備名稱', 
                        text_auto=True,
                        title=f"{query_dept} - 各設備每月用油統計",
                        labels={'加油量': '加油量 (L)', '月份': '統計月份'},
                        template="plotly_white"
                    )
                    fig.update_layout(barmode='stack')
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # 4. 詳細清單
                    st.subheader(f"📋 {query_dept} - 填報歷史明細")
                    display_cols = ["加油日期", "設備名稱", "油料種類", "加油量", "填報人", "單據備註"]
                    final_cols = [c for c in display_cols if c in df_final.columns]
                    df_display = df_final[final_cols].sort_values(by='加油日期', ascending=False)
                    st.dataframe(df_display, use_container_width=True)
                    
                else:
                    st.warning(f"⚠️ 在 {start_date} 到 {end_date} 期間，查無填報紀錄。")
            else:
                if query_dept:
                    st.info("請選擇完整的開始與結束日期。")

        st.markdown("""
            <div class="contact-footer">
            如有填報疑問，請電洽環安中心林小姐(分機 7137)，謝謝
            </div>
        """, unsafe_allow_html=True)

    # --- Tab 3: 管理 (Admin Only) ---
    if username == 'admin':
        with tabs[2]:
            st.header("🛠️ 設備資料庫管理")
            edited_df = st.data_editor(df_equip, num_rows="dynamic", use_container_width=True, key="editor")
            if st.button("💾 儲存變更", type="primary"):
                try:
                    ws_equip.clear()
                    updated_data = [edited_df.columns.tolist()] + edited_df.values.tolist()
                    ws_equip.update(updated_data)
                    st.success("✅ 更新成功！")
                    st.rerun()
                except Exception as e:
                    st.error(f"儲存失敗: {e}")