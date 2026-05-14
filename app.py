import streamlit as st
import streamlit_authenticator as stauth
import time
from streamlit.runtime.scriptrunner import get_script_run_ctx

# 0. 系統設定 (這是首頁的設定)
st.set_page_config(
    page_title="嘉義大學碳盤查", 
    page_icon="🏫", 
    layout="wide"
)

# 1. CSS 樣式 (已加入縮減空白與側邊欄標題優化)
st.markdown("""
<style>
    /* 縮減 Streamlit 預設的頂部與兩側空白 */
    .block-container { 
        padding-top: 1.5rem !important; 
        padding-bottom: 1rem !important; 
    }
    
    /* 主標題樣式 (已縮小 padding 與 margin) */
    .main-header {
        font-size: 2rem; font-weight: 800; color: #2C3E50; text-align: center; 
        margin-bottom: 5px; padding: 12px; background-color: #FFFFFF; 
        border-bottom: 3px solid #F4D03F; border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .info-box {
        background-color: #EBF5FB; border-left: 5px solid #3498DB; 
        padding: 15px; border-radius: 5px; margin-bottom: 15px;
    }
    .stApp { background-color: #F8F9F9; }

    /* 修改 Login 按鈕樣式 */
    div.stButton > button {
        background-color: #E67E22 !important; color: #FFFFFF !important;
        border: 2px solid #D35400 !important; border-radius: 12px !important;
        font-weight: 800 !important; box-shadow: 0 2px 4px rgba(0,0,0,0.1) !important;
    }
    div.stButton > button:hover {
        background-color: #D35400 !important; transform: translateY(-2px);
    }
    
    /* 微調側邊欄群組標題，增加視覺區隔 */
    [data-testid="stSidebarNavGroup"] > div {
        background-color: #E8F8F5; /* 淡淡的背景色 */
        border-radius: 5px;
        padding-left: 5px;
        margin-bottom: 5px;
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

# 2. 顯示標題
st.markdown('<div class="main-header">🏫 國立嘉義大學溫室氣體盤查填報系統<br><span style="font-size: 1.4rem; font-weight: 600; color: #5D6D7E;">National Chiayi University Greenhouse Gas Data Reporting System</span></div>', unsafe_allow_html=True)

# 3. 首頁引導文案 (虛擬分頁)
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
    # === 登入成功 ===
    current_username = st.session_state.get("username")
    
    # 取得當下即時人數，並套用您規劃的判斷邏輯
    online_count = get_online_user_status()
    if online_count >= 11:
        dot_color, status_text = "#FF4444", "擁擠，建議稍候"
    elif online_count >= 6:
        dot_color, status_text = "#FFBB33", "普通"
    else:
        dot_color, status_text = "#00C851", "順暢"
    
    with st.sidebar:
        # 動態顯示精簡版人數燈號
        st.markdown(f"""
            <div style="display: flex; align-items: center; margin-bottom: 15px; padding: 10px; background-color: #FFFFFF; border-radius: 8px; border: 1px solid #eee;">
                <span style="height: 12px; width: 12px; background-color: {dot_color}; border-radius: 50%; display: inline-block; margin-right: 10px; box-shadow: 0 0 5px {dot_color};"></span>
                <span style="font-weight: 700; color: #333333; font-size: 1.05rem;">線上人數: {online_count} 人 ({status_text})</span>
            </div>
        """, unsafe_allow_html=True)
        
        # 登出按鈕
        authenticator.logout('登出系統', 'sidebar')
        st.markdown("---")

    # ================= 宣告所有實體檔案路徑 =================
    # 基礎頁面
    home = st.Page(home_page, title="系統首頁", icon="🏠", default=True)
    fuel_report = st.Page("pages/1_⛽_燃油設備填報.py", title="燃油設備填報", icon="⛽")
    refrig_report = st.Page("pages/2_❄️_冷媒設備填報.py", title="冷媒設備填報", icon="❄️")
    
    # 原有的管理者頁面
    fuel_admin = st.Page("pages/3_⛽_燃油後台管理.py", title="燃油後台管理", icon="⚙️")
    refrig_admin = st.Page("pages/4_❄️_冷媒後台管理.py", title="冷媒後台管理", icon="⚙️")
    fuel_view = st.Page("pages/5_⛽_燃油資料檢視確認.py", title="燃油資料檢視確認", icon="📊")
    
    # 📌 新增的管理者頁面
    elec_report = st.Page("pages/6_⚡_全校電力填報.py", title="全校電力填報", icon="⚡")
    elec_admin = st.Page("pages/7_⚡_全校電力管理.py", title="全校電力管理", icon="⚡")
    gas_report = st.Page("pages/8_💨_實驗室氣體鋼瓶資料回報.py", title="氣體鋼瓶資料回報", icon="💨")
    gas_admin = st.Page("pages/9_💨_實驗室氣體鋼瓶資料管理與追蹤.py", title="氣體鋼瓶管理與追蹤", icon="💨")
    # ========================================================
    
    # 建立基礎選單 (所有人可見)
    pages_dict = {
        "": [home], # 隱藏大標題，精簡顯示
        "📝 設備填報作業": [fuel_report, refrig_report]
    }
    
    # 🔐 權限控管：當登入者為管理員時，加入所有後台與新增功能的選單
    admin_users = ["admin"] 
    if current_username in admin_users:
        # 將管理員功能進行分組，側邊欄才不會太亂
        pages_dict["⚙️ 燃油與冷媒管理"] = [fuel_admin, refrig_admin, fuel_view]
        pages_dict["⚡ 全校電力管理"] = [elec_report, elec_admin]
        pages_dict["💨 氣體鋼瓶管理"] = [gas_report, gas_admin]
        
    # 執行導航
    pg = st.navigation(pages_dict)
    pg.run()

elif st.session_state.get("authentication_status") is False:
    st.error('❌ 帳號或密碼錯誤，請重試')
elif st.session_state.get("authentication_status") is None:
    st.warning('🔒 請輸入帳號密碼以登入系統')