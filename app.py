import streamlit as st
import streamlit_authenticator as stauth

# 0. 系統設定
st.set_page_config(page_title="嘉義大學碳盤查", page_icon="🏫", layout="wide")

# 1. CSS 樣式 (定義標題與文字風格)
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem; font-weight: 800; color: #2C3E50; text-align: center; 
        margin-bottom: 20px; padding: 20px; background-color: #FFFFFF; 
        border-bottom: 3px solid #F4D03F; border-radius: 10px;
    }
    .info-box {
        background-color: #EBF5FB; border-left: 5px solid #3498DB; 
        padding: 20px; border-radius: 5px; margin-bottom: 20px;
    }
    .stApp { background-color: #F8F9F9; }
</style>
""", unsafe_allow_html=True)

# --- 🟢 修正重點：在這裡就先顯示標題，這樣登入前也看得到 ---
st.markdown('<div class="main-header">🏫 國立嘉義大學<br>溫室氣體盤查填報系統</div>', unsafe_allow_html=True)

# 2. 身份驗證邏輯
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
        # === 登入成功後顯示的內容 ===
        name = st.session_state["name"]
        
        # 側邊欄登出按鈕
        with st.sidebar:
            st.write(f"👤 歡迎, {name}")
            authenticator.logout('登出系統', 'sidebar')
            st.markdown("---")
            st.info("請點擊上方的頁面名稱進行切換")

        st.markdown(f"""
        <div class="info-box">
            <h4>👋 歡迎回來，{name}！</h4>
            <p>本系統採用分流架構，請查看 <strong>👈 左側側邊欄 (Sidebar)</strong> 的選單來進入功能：</p>
            <ul>
                <li><strong>1_⛽_燃油填報</strong>：公務車輛、農用機具、發電機等油料使用申報。</li>
                <li><strong>2_❄️_冷媒填報</strong>：冷氣機、冰水主機等冷媒填充維修申報。</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
        
        st.info("💡 如果左側沒有看到選單，請嘗試重新整理網頁，或點擊左上角的箭頭展開側邊欄。")

    elif st.session_state["authentication_status"] is False:
        st.error('❌ 帳號或密碼錯誤')
    elif st.session_state["authentication_status"] is None:
        st.warning('🔒 請輸入帳號密碼以進入系統')

except Exception as e:
    st.error(f"系統錯誤: {e}")