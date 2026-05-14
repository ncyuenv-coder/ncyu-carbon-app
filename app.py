import streamlit as st
import streamlit_authenticator as stauth

# 0. 系統設定 (必須在最頂端)
st.set_page_config(
    page_title="嘉義大學碳盤查", 
    page_icon="🏫", 
    layout="wide"
)

# 1. CSS 樣式
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem; font-weight: 800; color: #2C3E50; text-align: center; 
        margin-bottom: 20px; padding: 20px; background-color: #FFFFFF; 
        border-bottom: 3px solid #F4D03F; border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .stApp { background-color: #F8F9F9; }
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

# 2. 顯示標題 (所有頁面共用的 Header)
st.markdown('<div class="main-header">🏫 國立嘉義大學溫室氣體盤查填報系統<br><span style="font-size: 1.5rem; font-weight: 600; color: #5D6D7E;">National Chiayi University Greenhouse Gas Data Reporting System</span></div>', unsafe_allow_html=True)

# 3. 身份驗證邏輯
def clean_secrets(obj):
    if isinstance(obj, dict) or "AttrDict" in str(type(obj)): return {k: clean_secrets(v) for k, v in obj.items()}
    elif isinstance(obj, list): return [clean_secrets(i) for i in obj]
    return obj

try:
    _raw_creds = st.secrets["credentials"]
    credentials_login = clean_secrets(_raw_creds)
    cookie_cfg = st.secrets["cookie"]
    authenticator = stauth.Authenticate(credentials_login, cookie_cfg["name"], cookie_cfg["key"], cookie_cfg["expiry_days"])
    
    # 執行登入
    authenticator.login('main')

    if st.session_state["authentication_status"]:
        # === 登入成功：建立導航選單 ===
        with st.sidebar:
            st.header(f"👤 歡迎，師長/同仁")
            st.success("☁️ 連線成功")
            authenticator.logout('登出系統', 'sidebar')
            st.markdown("---")
        
        # 🌟 這裡就是新版路由的核心！
        # 如果你的檔名不同，請修改對應的 "pages/檔名.py"
        pages_setup = {
            "系統概覽": [
                st.Page("pages/home.py", title="首頁與系統公告", icon="🏠", default=True),
            ],
            "碳排資料填報": [
                st.Page("pages/fuel.py", title="燃油填報作業", icon="⛽"),
                st.Page("pages/refrigerant.py", title="冷媒填報作業", icon="❄️"),
            ],
            "智慧輔助工具": [
                st.Page("pages/invoice_ocr.py", title="AI 設備油單辨識", icon="🤖"),
            ]
        }
        
        # 啟動路由
        pg = st.navigation(pages_setup)
        pg.run()

    elif st.session_state["authentication_status"] is False:
        st.error('❌ 帳號或密碼錯誤，請重試')
    elif st.session_state["authentication_status"] is None:
        st.warning('🔒 請輸入帳號密碼以登入系統')

except Exception as e:
    st.error(f"系統啟動錯誤: {e}")