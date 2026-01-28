import streamlit as st
import streamlit_authenticator as stauth

# 0. 系統設定 (這是首頁的設定)
st.set_page_config(
    page_title="嘉義大學碳盤查", 
    page_icon="🏫", 
    layout="wide"
)

# 1. CSS 樣式 (定義首頁標題與按鈕樣式)
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

# 2. 顯示標題 (不管有無登入都顯示)
st.markdown('<div class="main-header">🏫 國立嘉義大學<br>溫室氣體盤查填報系統</div>', unsafe_allow_html=True)

# 3. 身份驗證邏輯
def clean_secrets(obj):
    if isinstance(obj, dict) or "AttrDict" in str(type(obj)): return {k: clean_secrets(v) for k, v in obj.items()}
    elif isinstance(obj, list): return [clean_secrets(i) for i in obj]
    return obj

try:
    # 讀取 secrets
    _raw_creds = st.secrets["credentials"]
    credentials_login = clean_secrets(_raw_creds)
    cookie_cfg = st.secrets["cookie"]
    authenticator = stauth.Authenticate(credentials_login, cookie_cfg["name"], cookie_cfg["key"], cookie_cfg["expiry_days"])
    
    # 執行登入 (這會自動處理 UI)
    authenticator.login('main')

    # 判斷登入狀態
    if st.session_state["authentication_status"]:
        # === 登入成功 ===
        name = st.session_state["name"]
        
        # 側邊欄：顯示歡迎與登出
        with st.sidebar:
            # 修改點 3: 文字改成 "歡迎，師長/同仁"
            st.header(f"👤 歡迎，師長/同仁")
            st.success("☁️ 連線成功")
            authenticator.logout('登出系統', 'sidebar')
            st.markdown("---")
            # 修改點 2: 移除了 "👇 請點擊上方頁面切換功能"

        # 主畫面：顯示導引說明 (修改點 1: 更新歡迎詞)
        st.markdown(f"""
        <div class="info-box">
            <h4>👋 您好！歡迎來到本校「溫室氣體盤查填報系統」</h4>
            <p>請查看 <strong>👈 左側側邊欄 (Sidebar)</strong> 的選單來進入功能：</p>
            <ul>
                <li><strong>1_⛽_燃油填報</strong>：公務車輛、發電機、除草機、農用機具等油料使用申報。</li>
                <li><strong>2_❄️_冷媒填報</strong>：冷氣機、冷凍/藏設備等冷媒填充維修申報。</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
        
        st.info("💡 提示：如果左側沒有看到選單，請點擊左上角的箭頭展開側邊欄 (Sidebar)。")

    elif st.session_state["authentication_status"] is False:
        st.error('❌ 帳號或密碼錯誤，請重試')
    elif st.session_state["authentication_status"] is None:
        st.warning('🔒 請輸入帳號密碼以登入系統')

except Exception as e:
    st.error(f"系統啟動錯誤: {e}")