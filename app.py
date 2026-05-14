import streamlit as st
import streamlit_authenticator as stauth
import time
from streamlit.runtime.scriptrunner import get_script_run_ctx

# 0. 系統設定
st.set_page_config(
    page_title="嘉義大學碳盤查", 
    page_icon="🏫", 
    layout="wide"
)

# 1. CSS 樣式
st.markdown("""
<style>
    /* 縮減 Streamlit 預設的頂部與兩側空白 */
    .block-container { 
        padding-top: 1.5rem !important; 
        padding-bottom: 1rem !important; 
    }
    
    /* 主標題樣式 */
    .main-header {
        font-size: 2rem; font-weight: 800; color: #2C3E50; text-align: center; 
        margin-bottom: 20px; padding: 12px; background-color: #FFFFFF; 
        border-bottom: 3px solid #F4D03F; border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .info-box {
        background-color: #EBF5FB; border-left: 5px solid #3498DB; 
        padding: 15px; border-radius: 5px; margin-bottom: 15px;
    }
    .stApp { background-color: #F8F9F9; }

    /* 修改 Login 與 Logout 按鈕樣式 */
    [data-testid="stSidebar"] div.stButton > button,
    div.stButton > button {
        background-color: #E67E22 !important; color: #FFFFFF !important;
        border: 2px solid #D35400 !important; border-radius: 12px !important;
        font-weight: 800 !important; box-shadow: 0 2px 4px rgba(0,0,0,0.1) !important;
    }
    [data-testid="stSidebar"] div.stButton > button:hover,
    div.stButton > button:hover {
        background-color: #D35400 !important; transform: translateY(-2px);
    }
    
    /* 微調側邊欄群組標題 */
    [data-testid="stSidebarNavGroup"] > div {
        background-color: #E8F8F5; 
        border-radius: 5px; padding-left: 5px; margin-bottom: 5px;
    }
    
    /* === 新增：右上角絕對定位燈號樣式 === */
    .online-badge {
        position: absolute;
        top: 1rem;
        right: 1.5rem;
        font-size: 0.95rem;
        font-weight: 600;
        color: #4A4A4A;
        background-color: rgba(255, 255, 255, 0.85); /* 帶有微透明背景，避免與後面文字重疊難辨識 */
        padding: 4px 10px;
        border-radius: 8px;
        z-index: 999;
    }
</style>
""", unsafe_allow_html=True)

# ===== 自動偵測線上人數邏輯 =====
@st.cache_resource
def get_session_tracker():
    return {}

def get_online_user_status():
    tracker = get_session_tracker()
    ctx = get_script_run_ctx()
    current_time = time.time()
    
    if ctx:
        tracker[ctx.session_id] = current_time
        
    active_count = 0
    for sid, last_seen in list(tracker.items()):
        if current_time - last_seen > 300:
            del tracker[sid]
        else:
            active_count += 1
            
    active_count = max(1, active_count)
    return active_count
# ===============================

# 3. 首頁引導文案 (已將主標題移出)
def home_page():
    st.markdown("""
    <div class="info-box">
        <h4>👋 您好！歡迎來到本校「溫室氣體盤查填報系統」</h4>
        <p>請查看 <strong>👈 左側側邊欄 (Sidebar)</strong> 的選單來進入對應功能：</p>
        <ul>
            <li><strong>燃油設備填報</strong>：公務車輛、發電機、除草機、農用機具等油料使用申報。</li>
            <li><strong>冷媒設備填報</strong>：冷氣機、冷凍/藏設備等冷媒填充維修申報。</li>
            <li><strong>後台管理與資料檢視</strong>：系統管理員專用之數據維護與檢視確認作業。</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)
    st.info("💡 提示：如果左側沒有看到選單，請點擊左上角的箭頭展開側邊欄 (Sidebar)。")

# 4. 身份驗證邏輯
def clean_secrets(obj):
    if isinstance(obj, dict) or "AttrDict" in str(type(obj)): return {k: clean_secrets(v) for k, v in obj.items()}
    elif isinstance(obj, list): return [clean_secrets(i) for i in obj]
    return obj

try:
    _raw_creds = st.secrets["credentials"]
    credentials_login = clean_secrets(_raw_creds)
    cookie_cfg = st.secrets["cookie"]
    authenticator = stauth.Authenticate(credentials_login, cookie_cfg["name"], cookie_cfg["key"], cookie_cfg["expiry_days"])
    
    authenticator.login('main')
except Exception as e:
    st.error(f"系統認證初始化錯誤: {e}")
    st.stop()

# 5. 判斷登入狀態與動態路由
if st.session_state.get("authentication_status"):
    current_username = st.session_state.get("username")
    
    with st.sidebar:
        authenticator.logout('登出系統', 'sidebar')
        st.markdown("---")

    # 宣告所有實體檔案路徑
    home = st.Page(home_page, title="系統首頁", icon="🏠", default=True)
    fuel_report = st.Page("pages/1_⛽_燃油設備填報.py", title="燃油設備填報", icon="⛽")
    refrig_report = st.Page("pages/2_❄️_冷媒設備填報.py", title="冷媒設備填報", icon="❄️")
    
    fuel_admin = st.Page("pages/3_⛽_燃油後台管理.py", title="燃油後台管理", icon="⚙️")
    refrig_admin = st.Page("pages/4_❄️_冷媒後台管理.py", title="冷媒後台管理", icon="⚙️")
    fuel_view = st.Page("pages/5_⛽_燃油資料檢視確認.py", title="燃油資料檢視確認", icon="📊")
    
    elec_report = st.Page("pages/6_⚡_全校電力填報.py", title="全校電力填報", icon="⚡")
    elec_admin = st.Page("pages/7_⚡_全校電力管理.py", title="全校電力管理", icon="⚡")
    gas_report = st.Page("pages/8_💨_實驗室氣體鋼瓶資料回報.py", title="氣體鋼瓶資料回報", icon="💨")
    gas_admin = st.Page("pages/9_💨_實驗室氣體鋼瓶資料管理與追蹤.py", title="氣體鋼瓶管理與追蹤", icon="💨")
    
    # 建立基礎選單
    pages_dict = {
        "": [home], 
        "📝 設備填報作業": [fuel_report, refrig_report]
    }
    
    # 🔐 權限控管
    admin_users = ["admin"] 
    if current_username in admin_users:
        pages_dict["⚙️ 燃油與冷媒管理"] = [fuel_admin, refrig_admin, fuel_view]
        pages_dict["⚡ 全校電力管理"] = [elec_report, elec_admin]
        pages_dict["💨 氣體鋼瓶管理"] = [gas_report, gas_admin]
        
    pg = st.navigation(pages_dict)

    # ==========================================================
    # 核心優化 1：絕對定位的人數燈號 (零高度佔用)
    online_count = get_online_user_status()
    if online_count >= 11: 
        status_html = f"🔴 目前線上人數: {online_count} 人 (擁擠，建議稍候)"
    elif online_count >= 6: 
        status_html = f"🟡 目前線上人數: {online_count} 人 (普通，可正常填報)"
    else: 
        status_html = f"🟢 目前線上人數: {online_count} 人 (順暢)"

    st.markdown(f'<div class="online-badge">{status_html}</div>', unsafe_allow_html=True)
    
    # 核心優化 2：動態判斷，僅在「首頁」與兩個「填報分頁」顯示系統標題
    if pg.title in ["系統首頁", "燃油設備填報", "冷媒設備填報"]:
        st.markdown('<div class="main-header">🏫 國立嘉義大學溫室氣體盤查填報系統<br><span style="font-size: 1.4rem; font-weight: 600; color: #5D6D7E;">National Chiayi University Greenhouse Gas Data Reporting System</span></div>', unsafe_allow_html=True)
    # ==========================================================
    
    # 最後再執行該分頁的內容
    pg.run()

elif st.session_state.get("authentication_status") is False:
    st.error('❌ 帳號或密碼錯誤，請重試')
elif st.session_state.get("authentication_status") is None:
    st.warning('🔒 請輸入帳號密碼以登入系統')