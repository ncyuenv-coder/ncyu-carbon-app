import streamlit as st
import streamlit_authenticator as stauth

# 0. 系統設定
st.set_page_config(
    page_title="國立嘉義大學碳盤查平台",
    page_icon="🏫",
    layout="wide"
)

# 1. CSS 優化 (全域)
st.markdown("""
<style>
    .main-header {font-size: 2.5rem; font-weight: 800; color: #2C3E50; text-align: center; margin-bottom: 20px;}
    .sub-header {font-size: 1.2rem; color: #566573; text-align: center; margin-bottom: 40px;}
    .info-box {background-color: #EBF5FB; border-left: 5px solid #3498DB; padding: 20px; border-radius: 5px; margin-bottom: 20px;}
    [data-testid="stSidebarNav"] {background-color: #F8F9F9;}
</style>
""", unsafe_allow_html=True)

# 2. 登入檢查 (共用 Session)
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

    if st.session_state["authentication_status"] is False:
        st.error('❌ 帳號或密碼錯誤')
    elif st.session_state["authentication_status"] is None:
        st.info('🔒 請輸入帳號密碼登入系統')
    else:
        # 登入成功畫面
        name = st.session_state["name"]
        with st.sidebar:
            st.write(f"👤 歡迎, {name}")
            authenticator.logout('登出', 'sidebar')

        st.markdown('<div class="main-header">🏫 國立嘉義大學<br>溫室氣體盤查填報系統</div>', unsafe_allow_html=True)
        
        st.markdown("""
        <div class="info-box">
            <h4>👋 歡迎使用本平台</h4>
            <p>本系統採用分流架構，請點擊左側選單進入所需功能：</p>
            <ul>
                <li><strong>⛽ 燃油填報</strong>：公務車輛、農用機具、發電機等油料使用申報。</li>
                <li><strong>❄️ 冷媒填報</strong>：冷氣機、冰水主機等冷媒填充維修申報。</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

        st.info("👈 請點擊左側側邊欄 (Sidebar) 切換填報項目")
        
        st.markdown("---")
        st.caption("系統維護：環安中心 | 分機 7137")

except Exception as e:
    st.error(f"系統啟動錯誤: {e}")