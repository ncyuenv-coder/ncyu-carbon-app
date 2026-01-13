import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
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
    
    /* 👇 優化：表單區改為「莫蘭迪綠 (Sage Green)」以區隔背景 */
    div[data-testid="stForm"] { 
        background-color: #E8F6F3; 
        padding: 30px; 
        border-radius: 20px; 
        border: 2px solid #A3E4D7;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    
    .info-card { background-color: #FEF9E7; padding: 15px; border-left: 5px solid #F4D03F; border-radius: 5px; margin-bottom: 10px; font-size: 1.1rem; }
    .info-label { font-weight: bold; color: #7F8C8D; }
    .info-value { color: #212F3D; font-weight: 600; margin-left: 10px; }
    
    /* 強制顯示必填星號的顏色 */
    .st-emotion-cache-1629p8f span { color: red; }
</style>
""", unsafe_allow_html=True)

# ☁️ 設定區
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

try:
    _raw_creds = st.secrets["credentials"]
    credentials = clean_secrets(_raw_creds)
    cookie_cfg = st.secrets["cookie"]
    
    authenticator = stauth.Authenticate(
        credentials,
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
# 2. 雲端連線
# ==========================================
@st.cache_resource
def init_google():
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
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
        ws_record = sh.add_worksheet(title="填報紀錄", rows="1000", cols="12") # 增加欄位數
        
    # 自動檢查並補齊標題列 (新增了 填報人、分機)
    if len(ws_record.get_all_values()) == 0:
        ws_record.append_row(["填報時間", "填報單位", "填報帳號", "填報人", "聯絡分機", "設備名稱", "校內財產編號", "加油日期", "加油量", "佐證檔案", "單據備註"])

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
    st.markdown("---")
    st.caption("系統維護單位：環安中心")

elif st.session_state['current_page'] == 'fuel':
    st.title("⛽ 燃油設備填報專區")
    
    if username == 'admin':
        tab1, tab2, tab3 = st.tabs(["📝 新增填報", "📊 數據看板", "🛠️ 資料庫管理"])
    else:
        tab1, tab2 = st.tabs(["📝 新增填報", "📊 數據看板"])
        tab3 = None 

    # --- Tab 1: 填報 ---
    with tab1:
        if not df_equip.empty:
            st.markdown("#### 步驟 1：選擇設備")
            c1, c2 = st.columns(2)
            units = sorted([x for x in df_equip['填報單位'].unique() if x != '-' and x != '填報單位'])
            selected_dept = c1.selectbox("填報單位", units)
            filtered = df_equip[df_equip['填報單位'] == selected_dept]
            devices = sorted([x for x in filtered['設備名稱備註'].unique()])
            selected_device = c2.selectbox("車輛/機具名稱", devices)
            
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
                with st.form("entry_form"):
                    # 第一列：填報人資訊 (新增)
                    col_p1, col_p2 = st.columns(2)
                    p_name = col_p1.text_input("👤 填報人姓名 (必填)")
                    p_ext = col_p2.text_input("📞 聯絡分機 (必填)")
                    
                    st.divider() # 分隔線
                    
                    # 第二列：油量資訊
                    col_a, col_b = st.columns(2)
                    d_date = col_a.date_input("📅 加油日期 (以加油單為準)", datetime.today())
                    d_vol = col_b.number_input("💧 加油量 (公升)", min_value=0.0, step=0.1, format="%.1f")
                    
                    # 備註欄 (修正文字)
                    st.markdown("**🧾 單據備註 (若一張發票加多台設備，請填寫相同發票號碼以便核對)**")
                    note = st.text_input("單據號碼/備註 (選填)", placeholder="例如：發票 AB-12345678")

                    st.markdown("---")
                    # 檔案上傳 (修正：多檔 + 大小限制)
                    st.markdown("**📂 上傳佐證資料 (最多 3 個檔案，單檔限制 10MB)**")
                    f_files = st.file_uploader("", type=['png', 'jpg', 'jpeg', 'pdf'], accept_multiple_files=True)
                    
                    submitted = st.form_submit_button("🚀 確認送出資料", type="primary", use_container_width=True)
                    
                    if submitted:
                        # 驗證邏輯
                        if not p_name or not p_ext:
                            st.warning("⚠️ 「填報人姓名」與「聯絡分機」為必填欄位！")
                        elif d_vol <= 0:
                            st.warning("⚠️ 加油量不能為 0")
                        else:
                            # 檔案檢查
                            valid_files = True
                            if f_files:
                                if len(f_files) > 3:
                                    st.error("❌ 超過檔案數量上限 (最多 3 個)")
                                    valid_files = False
                                for f in f_files:
                                    if f.size > 10 * 1024 * 1024: # 10MB
                                        st.error(f"❌ 檔案 {f.name} 太大 (超過 10MB)")
                                        valid_files = False
                            
                            if valid_files:
                                progress_text = "資料處理中..."
                                my_bar = st.progress(0, text=progress_text)
                                
                                # 檔案上傳迴圈
                                file_links = []
                                if f_files:
                                    for idx, f_file in enumerate(f_files):
                                        try:
                                            # 檔名加入序號避免重複
                                            clean_name = f"{selected_dept}_{selected_device}_{d_date}_{idx+1}.{f_file.name.split('.')[-1]}".replace("/", "_")
                                            file_meta = {'name': clean_name, 'parents': [DRIVE_FOLDER_ID]}
                                            media = MediaIoBaseUpload(f_file, mimetype=f_file.type)
                                            file = drive_service.files().create(body=file_meta, media_body=media, fields='webViewLink').execute()
                                            file_links.append(file.get('webViewLink'))
                                        except Exception as e:
                                            st.warning(f"檔案 {f_file.name} 上傳異常: {e}")
                                
                                # 將多個連結合併成一個字串存入 (用換行符號分隔)
                                final_links = "\n".join(file_links) if file_links else "無"

                                my_bar.progress(50, text="寫入資料庫...")
                                
                                # 寫入完整資料 (包含新增的 p_name, p_ext)
                                ws_record.append_row([
                                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                    selected_dept, name, 
                                    p_name, p_ext, # 新增欄位
                                    selected_device,
                                    str(row.get('校內財產編號', '-')), str(d_date), d_vol, 
                                    final_links, note
                                ])
                                
                                my_bar.progress(100, text="完成！")
                                time.sleep(0.5)
                                my_bar.empty()
                                st.success(f"✅ 成功！已新增紀錄：{d_vol} L")
                                st.balloons()

    # --- Tab 2: 看板 ---
    with tab2:
        col_r1, col_r2 = st.columns([4, 1])
        with col_r2:
            if st.button("🔄 刷新數據", use_container_width=True): 
                st.cache_data.clear()
                st.rerun()
        
        if not df_records.empty and '加油量' in df_records.columns:
            df_records['加油量'] = pd.to_numeric(df_records['加油量'], errors='coerce').fillna(0)
            total = df_records['加油量'].sum()
            count = len(df_records)
            last_date = df_records['加油日期'].max() if '加油日期' in df_records.columns else "-"
            
            m1, m2, m3 = st.columns(3)
            m1.metric("🛢️ 全校總油量", f"{total:,.1f} L")
            m2.metric("📝 總填報筆數", f"{count} 筆")
            m3.metric("📅 最新填報日", str(last_date))
            st.markdown("---")
            
            st.subheader("📋 詳細填報清冊")
            st.dataframe(df_records, use_container_width=True)

    # --- Tab 3: 管理 ---
    if tab3:
        with tab3:
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