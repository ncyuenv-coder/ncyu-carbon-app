import streamlit as st
import streamlit_authenticator as stauth

# 0. 系統設定 (這是首頁的設定)
st.set_page_config(
    page_title="嘉義大學碳盤查", 
    page_icon="🏫", 
    layout="wide"
)

# 1. CSS 樣式
st.markdown("""
<style>
    /* 主標題樣式 */
    .main-header {
        font-size: 2.2rem; font-weight: 800; color: #2C3E50; text-align: center; 
        margin-bottom: 20px; padding: 20px; background-color: #FFFFFF; 
        border-bottom: 3px solid #F4D03F; border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    /* 資訊框樣式 */
    .info-box {
        background-color: #EBF5FB; border-left: 5px solid #3498DB; 
        padding: 20px; border-radius: 5px; margin-bottom: 20px;
    }
    .stApp { background-color: #F8F9F9; }

    /* 修改 Login 按鈕樣式 (橘色底 + 白字) */
    div.stButton > button {
        background-color: #E67E22 !important;
        color: #FFFFFF !important;
        border: 2px solid #D35400 !important;
        border-radius: 12px !important;
        font-size: 1.1rem !important;
        font-weight: 800 !important;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1) !important;
    }
    div.stButton > button:hover {
        background-color: #D35400 !important;
        transform: translateY(-2px);
    }
</style>
""", unsafe_allow_html=True)

# 2. 顯示標題
st.markdown('<div class="main-header">🏫 國立嘉義大學溫室氣體盤查填報系統<br><span style="font-size: 1.5rem; font-weight: 600; color: #5D6D7E;">National Chiayi University Greenhouse Gas Data Reporting System</span></div>', unsafe_allow_html=True)

# ===== 首頁引導文案 (虛擬分頁) =====
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

# 3. 身份驗證邏輯
def clean_secrets(obj):
    if isinstance(obj, dict) or "AttrDict" in str(type(obj)): return {k: clean_secrets(v) for k, v in obj.items()}
    elif isinstance(obj, list): return [clean_secrets(i) for i in obj]
    return obj

# 嘗試初始化登入套件
try:
    _raw_creds = st.secrets["credentials"]
    credentials_login = clean_secrets(_raw_creds)
    cookie_cfg = st.secrets["cookie"]
    authenticator = stauth.Authenticate(credentials_login, cookie_cfg["name"], cookie_cfg["key"], cookie_cfg["expiry_days"])
    
    # 執行登入
    authenticator.login('main')
except Exception as e:
    st.error(f"系統認證初始化錯誤: {e}")
    st.stop()

# 4. 判斷登入狀態與動態路由
if st.session_state.get("authentication_status"):
    # === 登入成功 ===
    
    # 取得目前登入的帳號與姓名
    current_username = st.session_state.get("username")
    current_name = st.session_state.get("name")
    
    with st.sidebar:
        st.header(f"👤 歡迎，{current_name} 師長/同仁")
        st.success("☁️ 連線成功")
        # 統一在此處登出，分頁檔案中絕對不能再寫這行！
        authenticator.logout('登出系統', 'sidebar')
        st.markdown("---")

    # 宣告所有實體檔案路徑
    home = st.Page(home_page, title="系統首頁", icon="🏠", default=True)
    fuel_report = st.Page("pages/1_⛽_燃油設備填報.py", title="燃油設備填報", icon="⛽")
    refrig_report = st.Page("pages/2_❄️_冷媒設備填報.py", title="冷媒設備填報", icon="❄️")
    fuel_admin = st.Page("pages/3_⛽_燃油後台管理.py", title="燃油後台管理", icon="⚙️")
    refrig_admin = st.Page("pages/4_❄️_冷媒後台管理.py", title="冷媒後台管理", icon="⚙️")
    fuel_view = st.Page("pages/5_⛽_燃油資料檢視確認.py", title="燃油資料檢視確認", icon="📊")
    
    # 建立基礎選單 (包含 ncyu 在內的所有人都能看見)
    pages_dict = {
        "📌 系統導覽": [home],
        "📝 設備填報作業": [fuel_report, refrig_report]
    }
    
    # 🔐 權限控管：如果是管理者 admin，才加入後台管理的選單
    # ==========================================
    admin_users = ["admin"] 
    # ==========================================
    
    if current_username in admin_users:
        pages_dict["⚙️ 管理與檢視"] = [fuel_admin, refrig_admin, fuel_view]
        
    # 建立並執行美觀的側邊欄導航
    pg = st.navigation(pages_dict)
    pg.run()

elif st.session_state.get("authentication_status") is False:
    st.error('❌ 帳號或密碼錯誤，請重試')
elif st.session_state.get("authentication_status") is None:
    st.warning('🔒 請輸入帳號密碼以登入系統')