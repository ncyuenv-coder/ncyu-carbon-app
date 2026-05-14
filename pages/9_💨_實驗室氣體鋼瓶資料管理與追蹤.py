import streamlit as st
import pandas as pd
import gspread
import uuid
import datetime
import smtplib
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from google.oauth2.credentials import Credentials
import plotly.express as px
import plotly.graph_objects as go
from gspread.utils import rowcol_to_a1
import streamlit_authenticator as stauth

# ================= 系統與全域參數設定 =================
st.set_page_config(page_title="溫室氣體盤查範疇之氣體鋼瓶購買盤查與追蹤管理", page_icon="📢", layout="wide")

# ================= 身份驗證與初始化 =================
def clean_secrets(obj):
    if isinstance(obj, dict) or "AttrDict" in str(type(obj)): return {k: clean_secrets(v) for k, v in obj.items()}
    elif isinstance(obj, list): return [clean_secrets(i) for i in obj]
    return obj

try:
    _raw_creds = st.secrets["credentials"]
    credentials_login = clean_secrets(_raw_creds)
    cookie_cfg = st.secrets["cookie"]
    authenticator = stauth.Authenticate(credentials_login, cookie_cfg["name"], cookie_cfg["key"], cookie_cfg["expiry_days"])
except: pass

if st.session_state.get("authentication_status") is not True:
    st.warning("🔒 請先至首頁登入系統")
    st.stop()

username = st.session_state.get("username")
if username != 'admin':
    st.error("🚫 權限不足：此頁面僅供系統管理員使用。")
    st.stop()

# ================= 參數與色票設定 =================
SHEET_ID = '1Hw4rXo4ww7O9YXTwoUJeWioO5ZzM_bivRcLLpOl26DY'

# 🚀 精準修正：真實填報系統網址
BASE_FORM_URL = "https://ncyu-carbon-app-mduue5hffp7uknsskmjet9.streamlit.app/實驗室氣體鋼瓶資料回報"

GAS_TYPES = ["二氧化碳", "甲烷", "乙炔", "一氧化二氮(笑氣)"]
CAMPUS_OPTS = ["蘭潭校區", "民雄校區", "新民校區", "林森校區"]

MORANDI_PALETTE = ['#88B04B', '#92A8D1', '#F7CAC9', '#5D6D7E', '#7FB3D5', '#E59866', '#F7DC6F', '#CCD1D1', '#76D7C4']
MORANDI_LIGHT_BGS = ["#F0F4F8", "#FDF2E9", "#EBF5FB", "#F5EEF8", "#FDFBF7"]

# ================= 莫蘭迪風格 CSS =================
def apply_morandi_theme():
    st.markdown("""
    <style>
        [data-testid="stAppViewContainer"] { background-color: #FDFBF7; color: #2C3E50; }
        [data-testid="stHeader"] { background-color: rgba(0,0,0,0); }
        .metric-card { background-color: #FFFFFF; padding: 0; border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); text-align: center; overflow: hidden; border: 1px solid #9DB4AB; display: flex; flex-direction: column; }
        .metric-title { background-color: #4A5D6E; color: #FFFFFF; padding: 12px 15px; font-size: 18px; font-weight: bold; margin: 0; letter-spacing: 1px; }
        .metric-value { font-size: 48px; font-weight: bold; color: #000000 !important; padding: 15px; margin: 0; flex: 1; display: flex; align-items: baseline; justify-content: center; gap: 8px; }
        .metric-value span { font-size: 20px !important; color: #5C6B73 !important; font-weight: normal; }
        div[data-baseweb="tab-list"] button { background-color: #5C6B73 !important; border-radius: 8px 8px 0 0 !important; padding: 10px 25px !important; border: none !important; opacity: 0.8; }
        div[data-baseweb="tab-list"] button p { font-size: 20px !important; color: #FFFFFF !important; font-weight: 600 !important; }
        div[data-baseweb="tab-list"] button[aria-selected="true"] { background-color: #3F4E4F !important; opacity: 1; border-bottom: 3px solid #9DB4AB !important; }
        div[data-testid="stElementContainer"]:has(.basic-bg-marker) + div[data-testid="stVerticalBlock"] { background-color: #E4E9E5 !important; padding: 30px !important; border-radius: 12px !important; margin-bottom: 25px !important; border: 1px solid #C4CDC5; }
        div[data-testid="stElementContainer"]:has(.purple-bg-marker) + div[data-testid="stVerticalBlock"] { background-color: #ECE8F2 !important; padding: 30px !important; border-radius: 12px !important; margin-bottom: 25px !important; border: 1px solid #D6CFE1; }
        .edit-chem-title { color: #6C7A89; font-size: 20px; font-weight: bold; border-bottom: 2px solid #E4E9E5; padding-bottom: 8px; margin-bottom: 15px; }
        .edit-chem-block { padding: 25px; border-radius: 10px; border: 1px solid #DCE4E3; margin-bottom: 20px; box-shadow: 0 2px 5px rgba(0,0,0,0.02); }
        button[kind="primary"] { background-color: #27AE60 !important; color: #FFFFFF !important; border: none !important; font-weight: bold !important; font-size: 18px !important; padding: 10px 20px !important;}
        button[kind="primary"]:hover { background-color: #1E8449 !important; }
        div.stRadio > div[role="radiogroup"] > label { background-color: #E3F2FD !important; padding: 10px 20px !important; border-radius: 8px !important; margin-right: 15px !important; border: 2px solid #BBDEFB !important; cursor: pointer; }
        div.stRadio > div[role="radiogroup"] > label p { color: #0D47A1 !important; font-size: 18px !important; font-weight: bold !important; }
        div.stRadio > div[role="radiogroup"] > label[data-checked="true"] { background-color: #BBDEFB !important; border-color: #1976D2 !important; }
    </style>
    """, unsafe_allow_html=True)

# ================= Google API 與函式 =================
def with_retry(max_retries=3, base_delay=1.5):
    def decorator(func):
        def wrapper(*args, **kwargs):
            retries = 0
            while True:
                try: return func(*args, **kwargs)
                except Exception as e:
                    if retries >= max_retries: raise e
                    time.sleep(base_delay * (2 ** retries)); retries += 1
        return wrapper
    return decorator

@st.cache_resource
def get_credentials():
    oauth_secrets = st.secrets["gcp_oauth"]
    return Credentials(token=None, refresh_token=oauth_secrets["refresh_token"], client_id=oauth_secrets["client_id"], client_secret=oauth_secrets["client_secret"], token_uri="https://oauth2.googleapis.com/token")

def get_gc(): return gspread.authorize(get_credentials())

@st.cache_data(ttl=60, show_spinner=False)
def load_data():
    gc = get_gc(); sh = gc.open_by_key(SHEET_ID)
    try: df_inv = pd.DataFrame(sh.worksheet('氣體鋼瓶資料庫').get_all_records())
    except: df_inv = pd.DataFrame()
    try: df_pur = pd.DataFrame(sh.worksheet('氣體鋼瓶年度使用紀錄').get_all_records())
    except: df_pur = pd.DataFrame()
    try: df_trk = pd.DataFrame(sh.worksheet('發送紀錄與金鑰').get_all_records())
    except: df_trk = pd.DataFrame()
    return df_inv, df_pur, df_trk

@with_retry()
def safe_append_rows(worksheet, rows): worksheet.append_rows(rows)

@with_retry()
def safe_delete_row(worksheet, idx): worksheet.delete_rows(idx)

@with_retry()
def batch_update_safe(worksheet, updates): worksheet.batch_update(updates)

def send_email_action(to_email, subject, html_body):
    try:
        sender_email = st.secrets["email"]["sender_email"]
        app_password = st.secrets["email"]["app_password"]
        msg = MIMEMultipart('alternative')
        msg['Subject'], msg['From'], msg['To'] = subject, sender_email, to_email
        msg.attach(MIMEText(html_body, 'html'))
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(sender_email, app_password)
        server.send_message(msg); server.quit()
        return True
    except Exception as e:
        # 關鍵：將隱藏的錯誤丟到畫面上
        st.error(f"❌ 發送信件給 {to_email} 失敗！錯誤代碼：{str(e)}")
        return False

def generate_styled_email_html(email_body_text, title="溫室氣體盤查 氣體鋼瓶通知"):
    html_content = email_body_text.replace("\n", "<br>")
    return f"<html><body style='font-family: Arial, sans-serif; line-height: 1.6; color: #333; font-size: 15px;'><div style='max-width: 600px; margin: 0 auto; border: 1px solid #ddd; border-radius: 10px; overflow: hidden;'><div style='background-color: #5C6B73; color: white; padding: 15px 20px; text-align: center;'><h2 style='margin: 0;'>{title}</h2></div><div style='padding: 20px;'>{html_content}</div></div></body></html>"

# ================= 視覺化圖表渲染 =================
@st.fragment
def render_dashboard(df_pur):
    st.markdown("<br>", unsafe_allow_html=True)
    df_plot = df_pur.copy()
    if not df_plot.empty:
        df_plot['購買日期'] = pd.to_datetime(df_plot['購買日期'], errors='coerce')
        df_plot['年份'] = df_plot['購買日期'].dt.year.fillna(datetime.datetime.now().year).astype(int)
        df_plot['購買量'] = pd.to_numeric(df_plot['年度氣體鋼瓶購買量(公斤)'], errors='coerce').fillna(0)
    
    all_years = sorted(df_plot['年份'].unique().tolist(), reverse=True) if not df_plot.empty else [datetime.datetime.now().year]
    
    c_year, _ = st.columns([1, 3])
    sel_year = c_year.selectbox("📅 選擇統計年份", all_years, key="dash_year")
    st.markdown("---")
    
    df_year = df_plot[df_plot['年份'] == sel_year] if not df_plot.empty else pd.DataFrame()
    
    if not df_year.empty:
        st.markdown(f"<h3 style='text-align:center; color: #2C3E50; font-weight: bold; background-color: #FFFFFF; padding: 10px; border-radius: 10px; border: 1px solid #BDC3C7;'>{sel_year}年度 氣體鋼瓶購買統計與盤查概況</h3>", unsafe_allow_html=True)
        tot_kg = df_year['購買量'].sum()
        tot_count = len(df_year)
        tot_labs = df_year['氣體鋼瓶所在位置實驗室門牌'].nunique()
        
        c1, c2, c3 = st.columns(3)
        c1.markdown(f"""<div class="metric-card"><div class="metric-title" style="background-color: #A9CCE3;">📄 總購買申報筆數</div><div class="metric-value">{tot_count} <span>筆</span></div></div>""", unsafe_allow_html=True)
        c2.markdown(f"""<div class="metric-card"><div class="metric-title" style="background-color: #F5CBA7;">⚖️ 總氣體購買量</div><div class="metric-value">{tot_kg:,.2f} <span>公斤 (kg)</span></div></div>""", unsafe_allow_html=True)
        c3.markdown(f"""<div class="metric-card"><div class="metric-title" style="background-color: #A3E4D7;">🏢 涵蓋實驗室數量</div><div class="metric-value">{tot_labs} <span>間</span></div></div>""", unsafe_allow_html=True)
        
        st.markdown("---")
        st.markdown("<h3 style='color: #2C3E50;'>📈 年度氣體購買概況 (依氣體種類)</h3>", unsafe_allow_html=True)
        df_gas = df_year.groupby(['鋼瓶氣體種類', '校區'])['購買量'].sum().reset_index()
        fig1 = px.bar(df_gas, x='鋼瓶氣體種類', y='購買量', color='校區', text_auto='.2f', color_discrete_sequence=MORANDI_PALETTE)
        fig1.update_layout(height=500, yaxis_title="購買量 (公斤)", xaxis_title="", font=dict(size=16), plot_bgcolor='rgba(0,0,0,0)')
        fig1.update_yaxes(showgrid=True, gridcolor='#EAEDED')
        st.plotly_chart(fig1, use_container_width=True)
        
        st.markdown("---")
        st.markdown("<h3 style='color: #2C3E50;'>🏆 前十大氣體購買系所</h3>", unsafe_allow_html=True)
        top_depts = df_year.groupby('系所')['購買量'].sum().nlargest(10).index.tolist()
        df_top = df_year[df_year['系所'].isin(top_depts)].groupby(['系所', '鋼瓶氣體種類'])['購買量'].sum().reset_index()
        fig2 = px.bar(df_top, x='系所', y='購買量', color='鋼瓶氣體種類', text_auto='.2f', color_discrete_sequence=MORANDI_PALETTE)
        fig2.update_layout(height=500, xaxis={'categoryorder':'total descending'}, yaxis_title="購買量 (公斤)", xaxis_title="", font=dict(size=16), plot_bgcolor='rgba(0,0,0,0)')
        st.plotly_chart(fig2, use_container_width=True)

        st.markdown("---")
        st.markdown("<h3 style='color: #2C3E50;'>🍩 氣體購買資訊佔比分析</h3>", unsafe_allow_html=True)
        c_l, c_r = st.columns(2)
        gas_total = df_year.groupby('鋼瓶氣體種類')['購買量'].sum().reset_index()
        fig3_l = px.pie(gas_total, values='購買量', names='鋼瓶氣體種類', title='依氣體購買重量統計', hole=0.4, color_discrete_sequence=MORANDI_PALETTE)
        fig3_l.update_traces(textinfo='label+percent', textfont=dict(size=16), hovertemplate='<b>%{label}</b><br>重量: %{value:.2f} kg<br>佔比: %{percent:.1%}')
        fig3_r = px.sunburst(df_year, path=['校區', '系所', '鋼瓶氣體種類'], values='購買量', title='校區與系所結構分析 (旭日圖)', color='校區', color_discrete_sequence=MORANDI_PALETTE)
        fig3_r.update_traces(textinfo='label+percent parent')
        
        with c_l: st.plotly_chart(fig3_l, use_container_width=True)
        with c_r: st.plotly_chart(fig3_r, use_container_width=True)
    else:
        st.info("該年度目前尚無購買紀錄資料。")

# ================= 主程式 =================
def main():
    apply_morandi_theme()
    st.markdown('<div style="font-size: 2.4rem; font-weight: 900; color: #2C3E50; margin-bottom: 16px;">👑 溫室氣體盤查範疇之氣體鋼瓶購買盤查與追蹤管理</div>', unsafe_allow_html=True)
    
    if 'reset_key' not in st.session_state: st.session_state.reset_key = 0
    rk = st.session_state.reset_key
    df_inv, df_pur, df_trk = load_data()
    now_year_roc = datetime.datetime.now().year - 1911
    
    col_refresh = st.columns([8, 2])
    with col_refresh[1]:
        if st.button("🔄 強制刷新最新資料", use_container_width=True):
            load_data.clear(); st.rerun()

    tab1, tab2, tab3, tab4 = st.tabs(["📧 批次發送作業", "📊 回報進度追蹤與稽催", "📈 年度氣體鋼瓶購買量統計", "🛠️ 庫存資料管理"])

    # ================= Tab 1: 批次發送作業 =================
    with tab1:
        st.markdown("### 🔄 建立新一期盤查通知")
        if df_inv.empty: st.warning("⚠️ 氣體鋼瓶資料庫尚無任何紀錄，無法執行發送作業。")
        else:
            labs = df_inv[['系所', '實驗室老師', '電子郵件', '氣體鋼瓶所在位置實驗室門牌', '校區']].drop_duplicates()
            with st.form("send_form"):
                batch_name = st.text_input("📝 批次名稱", value=f"{now_year_roc}年度溫室氣體盤查範疇之氣體鋼瓶購買調查")
                data_year_roc = st.text_input("📅 盤查年度", value=f"{now_year_roc-1}年度")
                is_test_mode = st.radio("寄送模式", ["🧪 測試寄信模式", "🚀 正式寄信模式"], horizontal=True) == "🧪 測試寄信模式"
                test_email = st.text_input("📩 測試接收信箱") if is_test_mode else ""
                batch_subject = st.text_input("信件主旨", value="國立嘉義大學 {批次名稱}")
                batch_body = st.text_area("信件內容", value="{系所} {老師名稱} 您好：\n\n為配合溫室氣體盤查作業，請協助檢視與更新貴實驗室之基本資料、氣體鋼瓶庫存及使用量，並上傳{年度}之購買佐證單據。\n\n📍 您所管理的實驗室專屬確認連結如下：\n{links}\n\n⚠️ 提醒：若您管理多間實驗室，請於「氣體鋼瓶所在位置實驗室門牌」欄位填寫多個位置。\n本信件由系統自動發送，請勿直接回覆。", height=220)
                submit_btn = st.form_submit_button("🚀 產生金鑰並發送通知信", type="primary")
                
            if submit_btn:
                if is_test_mode and not test_email: st.error("請輸入測試信箱！")
                else:
                    my_bar = st.progress(0, text="處理中...")
                    grouped = labs.groupby(['系所', '實驗室老師', '電子郵件'])
                    all_appends, success_count = [], 0
                    total_groups = len(grouped)
                    now_time = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
                    
                    for idx, ((dept, mgr, email), group) in enumerate(grouped):
                        target_email = test_email if is_test_mode else email
                        links_html = ""
                        for _, row in group.iterrows():
                            token = uuid.uuid4().hex
                            # 精準修正：處理 URL 串接 ?token=
                            link = f"{BASE_FORM_URL}?token={token}"
                            links_html += f'<div style="margin-bottom: 8px; background-color: #FDFBF7; padding: 10px; border-left: 5px solid #D4A373;"><b>{row["校區"]} - 門牌 {row["氣體鋼瓶所在位置實驗室門牌"]}</b>：<a href="{link}" style="color: #3498DB; font-weight: bold;">點此進入系統</a></div>'
                            all_appends.append([batch_name, dept, mgr, row["校區"], row["氣體鋼瓶所在位置實驗室門牌"], email, token, link, "", now_time, "待回覆", ""])
                        
                        mail_content = batch_body.replace("{系所}", dept).replace("{老師名稱}", mgr).replace("{年度}", data_year_roc).replace("{links}", links_html)
                        html_wrap = generate_styled_email_html(mail_content, title=batch_name)
                        subject_content = batch_subject.replace("{批次名稱}", batch_name)
                        
                        # 🚀 精準修正：確認發信成功才標記成功
                        if send_email_action(target_email, subject_content, html_wrap):
                            success_count += 1
                            for record in all_appends[-len(group):]: record[8] = "發送成功"
                        else:
                            for record in all_appends[-len(group):]: record[8] = "發送失敗"
                        
                        my_bar.progress((idx+1)/total_groups, text=f"發送中 ({idx+1}/{total_groups})")
                    
                    safe_append_rows(get_gc().open_by_key(SHEET_ID).worksheet('發送紀錄與金鑰'), all_appends)
                    st.success(f"✅ 完成！成功寄出 {success_count} 位老師。請檢查 Gmail 已寄出資料夾。")
                    time.sleep(2); load_data.clear(); st.rerun()

    # ================= Tab 2: 進度追蹤與稽催 =================
    with tab2:
        if df_trk.empty: st.info("尚無發送追蹤紀錄。")
        else:
            batches = df_trk['發送批次'].unique().tolist()
            sel_batch = st.selectbox("📌 篩選發送批次", batches, index=len(batches)-1)
            df_cur = df_trk[df_trk['發送批次'] == sel_batch]
            
            pending = df_cur[df_cur['回覆狀態'] == '待回覆']
            replied = df_cur[df_cur['回覆狀態'] == '已回報']
            
            st.markdown(f"""
            <div style="display: flex; gap: 20px; margin-bottom: 25px;">
                <div class="metric-card" style="flex: 1;"><div class="metric-title">應回報總數</div><div class="metric-value">{len(df_cur)} <span>間</span></div></div>
                <div class="metric-card" style="flex: 1; border-color: #F39C12;"><div class="metric-title" style="background-color: #F39C12;">尚未回報</div><div class="metric-value">{len(pending)} <span>間</span></div></div>
                <div class="metric-card" style="flex: 1; border-color: #27AE60;"><div class="metric-title" style="background-color: #27AE60;">已回報完成</div><div class="metric-value">{len(replied)} <span>間</span></div></div>
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown("### ✉️ 一鍵稽催通知")
            follow_mode = st.radio("稽催發送模式", ["🧪 測試寄信模式", "🚀 正式寄信模式"], horizontal=True) == "🧪 測試寄信模式"
            f_test_email = st.text_input("📩 稽催測試接收信箱") if follow_mode else ""
            
            if not pending.empty:
                for dept in pending['系所'].unique():
                    dept_pending = pending[pending['系所'] == dept]
                    with st.expander(f"🏢 {dept} (未回報 {len(dept_pending)} 間)"):
                        col_list, col_btn = st.columns([8, 2])
                        grouped_pending = dept_pending.groupby(['實驗室老師', '電子郵件'])
                        with col_list:
                            for (mgr, email), group in grouped_pending:
                                rooms = ", ".join(group['氣體鋼瓶所在位置實驗室門牌'].tolist())
                                st.write(f"- **{mgr}** 老師 (門牌: {rooms})")
                        with col_btn:
                            if st.button(f"✉️ 稽催 {dept}", key=f"f_{dept}"):
                                if follow_mode and not f_test_email: st.error("請輸入測試信箱！")
                                else:
                                    f_count = 0
                                    for (mgr, email), group in grouped_pending:
                                        target = f_test_email if follow_mode else email
                                        display_mgr = f"{mgr} 老師" if "老師" not in mgr else mgr
                                        links_html = "".join([f'<div style="margin-bottom: 8px; background-color: #FDFBF7; padding: 10px; border-left: 5px solid #F39C12;">門牌 <b>{r["氣體鋼瓶所在位置實驗室門牌"]}</b>：<a href="{r["專屬填報網址"]}" style="color: #3498DB; font-weight: bold;">點此補登</a></div>' for _, r in group.iterrows()])
                                        msg_body = f"{dept} {display_mgr} 您好：<br><br>提醒您，系統尚未收到您的氣體鋼瓶購買與盤查資料。敬請撥冗點擊下方連結補登：<br><br>{links_html}"
                                        html_wrap = generate_styled_email_html(msg_body, title="氣體鋼瓶盤查 稽催提醒")
                                        if send_email_action(target, f"國立嘉義大學 {sel_batch} 未回報提醒", html_wrap): f_count += 1
                                    st.success(f"✅ {dept} 稽催發送完成 ({f_count} 封信)。")

    # ================= Tab 3: 年度氣體鋼瓶購買量統計 (圖表) =================
    with tab3:
        render_dashboard(df_pur)
        
        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown("### 📁 原始資料下載區")
        c_raw1, c_raw2 = st.tabs(["氣體鋼瓶資料庫 (現有總庫存)", "年度氣體鋼瓶使用紀錄 (購買憑證)"])
        with c_raw1: st.dataframe(df_inv)
        with c_raw2: 
            if not df_pur.empty:
                st.data_editor(df_pur, column_config={"購買單據連結": st.column_config.LinkColumn("單據檢視")}, hide_index=True)

    # ================= Tab 4: 🛠️ 庫存資料管理 =================
    with tab4:
        st.markdown("### 🛠️ 氣體鋼瓶資料庫管理")
        m_tab1, m_tab2 = st.tabs(["➕ 新增實驗室與庫存", "🔄 現有庫存查詢與異動"])
        gc = get_gc()
        ws_inv = gc.open_by_key(SHEET_ID).worksheet('氣體鋼瓶資料庫')
        
        with m_tab1:
            st.markdown('<div class="basic-bg-marker"></div>', unsafe_allow_html=True)
            with st.container():
                st.markdown('<div style="background-color: #8A9A8A; padding: 12px 20px; border-radius: 8px; margin-bottom: 20px;"><h3 style="color: #FFFFFF; margin: 0; font-weight: 600;">🏢 場所基本資料建立</h3></div>', unsafe_allow_html=True)
                col1, col2, col3 = st.columns(3)
                with col1: a_dept = st.text_input("系所", key=f"a_dept_{rk}")
                with col2: a_campus = st.selectbox("校區", CAMPUS_OPTS, key=f"a_camp_{rk}")
                with col3: a_room = st.text_input("氣體鋼瓶所在位置實驗室門牌", key=f"a_room_{rk}")
                col4, col5, col6 = st.columns(3)
                with col4: a_mgr = st.text_input("實驗室老師", key=f"a_mgr_{rk}")
                with col5: a_mail = st.text_input("電子郵件", key=f"a_mail_{rk}")
                with col6: a_ext = st.text_input("分機", key=f"a_ext_{rk}")
            
            st.markdown('<div class="purple-bg-marker"></div>', unsafe_allow_html=True)
            with st.container():
                st.markdown('<div style="background-color: #A393B3; padding: 12px 20px; border-radius: 8px; margin-bottom: 20px;"><h3 style="color: #FFFFFF; margin: 0; font-weight: 600;">💨 氣體鋼瓶初始庫存</h3></div>', unsafe_allow_html=True)
                num_gases = st.number_input("🧪 此場所氣體種類數量", min_value=1, max_value=5, value=1, step=1, key=f"a_num_{rk}")
                
                gas_entries = []
                for i in range(int(num_gases)):
                    st.markdown(f"#### 🔹 第 {i+1} 項氣體")
                    gc1, gc2, gc3 = st.columns(3)
                    with gc1: a_gas = st.selectbox("鋼瓶氣體種類", ["請選擇"] + GAS_TYPES, key=f"a_gas_{i}_{rk}")
                    with gc2: a_qty = st.number_input("鋼瓶數量", min_value=1, step=1, key=f"a_qty_{i}_{rk}")
                    with gc3: a_year = st.text_input("建檔年度", value=f"{now_year_roc}年", key=f"a_yr_{i}_{rk}")
                    gas_entries.append({"gas": a_gas, "qty": a_qty, "year": a_year})
                    st.markdown("<hr style='border-top: 1px dashed #B8A9C9; margin: 15px 0;'>", unsafe_allow_html=True)
            
            if st.button("🚀 確認新增建檔", type="primary", use_container_width=True):
                if not a_dept or not a_mgr or not a_room:
                    st.error("請確實填寫系所、老師與門牌！")
                else:
                    rows_to_append = []
                    for entry in gas_entries:
                        if entry["gas"] != "請選擇":
                            rows_to_append.append([a_dept, a_mgr, a_campus, a_room, a_mail, a_ext, entry["gas"], entry["qty"], entry["year"]])
                    if rows_to_append:
                        safe_append_rows(ws_inv, rows_to_append)
                        st.success(f"✅ {a_mgr} 老師的庫存已建檔！")
                        st.session_state.reset_key += 1
                        time.sleep(1); load_data.clear(); st.rerun()
                    else:
                        st.error("請至少選擇一種氣體！")

        with m_tab2:
            st.markdown("### 🔍 查詢與異動管理")
            if df_inv.empty: st.info("目前資料庫為空。")
            else:
                existing_depts = df_inv['系所'].astype(str).unique().tolist()
                f_col1, f_col2 = st.columns(2)
                with f_col1: sel_dept = st.selectbox("1. 篩選系所", existing_depts)
                
                dept_labs = df_inv[df_inv['系所'] == sel_dept]
                lab_opts = dept_labs.apply(lambda x: f"{x['實驗室老師']}({x['氣體鋼瓶所在位置實驗室門牌']})", axis=1).unique().tolist()
                
                with f_col2: sel_lab_opt = st.selectbox("2. 篩選老師與門牌", lab_opts)
                
                mgr_sel = sel_lab_opt.split("(")[0]
                room_sel = sel_lab_opt.split("(")[1].replace(")", "")
                
                target_df = df_inv[(df_inv['系所'] == sel_dept) & (df_inv['實驗室老師'] == mgr_sel) & (df_inv['氣體鋼瓶所在位置實驗室門牌'] == room_sel)]
                
                if not target_df.empty:
                    base_row = target_df.iloc[0]
                    st.markdown("<hr style='border-top: 2px dashed #DCE4E3; margin: 25px 0;'>", unsafe_allow_html=True)
                    st.markdown("#### 📝 基本資料修改")
                    ce1, ce2, ce3 = st.columns(3)
                    with ce1: e_mail = st.text_input("修改電子郵件", value=str(base_row['電子郵件']), key=f"e_m_{mgr_sel}_{room_sel}")
                    with ce2: e_ext = st.text_input("修改分機", value=str(base_row['分機']), key=f"e_e_{mgr_sel}_{room_sel}")
                    
                    st.markdown("#### 🧪 現有氣體庫存編輯")
                    collected_updates = []
                    all_rows = ws_inv.get_all_records()
                    
                    for idx, row in target_df.iterrows():
                        gas_name = row['鋼瓶氣體種類']
                        bg_color = MORANDI_LIGHT_BGS[idx % len(MORANDI_LIGHT_BGS)]
                        
                        orig_idx = next((i + 2 for i, r in enumerate(all_rows) if str(r['系所'])==sel_dept and str(r['實驗室老師'])==mgr_sel and str(r['氣體鋼瓶所在位置實驗室門牌'])==room_sel and str(r['鋼瓶氣體種類'])==gas_name), -1)
                        
                        st.markdown(f"<div class='edit-chem-block' style='background-color: {bg_color};'>", unsafe_allow_html=True)
                        st.markdown(f"<div class='edit-chem-title'>🔹 {gas_name}</div>", unsafe_allow_html=True)
                        uc1, uc2, uc3 = st.columns([1, 1, 2])
                        with uc1: u_qty = st.number_input("鋼瓶數量", value=int(row['鋼瓶數量']), min_value=0, key=f"uq_{idx}_{mgr_sel}_{room_sel}")
                        with uc3:
                            st.markdown("<br>", unsafe_allow_html=True)
                            del_flag = st.checkbox("🗑️ 刪除此項氣體", key=f"del_{idx}_{mgr_sel}_{room_sel}")
                        st.markdown("</div>", unsafe_allow_html=True)
                        
                        collected_updates.append({"orig_idx": orig_idx, "gas": gas_name, "qty": u_qty, "delete": del_flag})
                        
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button("💾 儲存所有異動", type="primary", use_container_width=True):
                        with st.spinner("同步至雲端資料庫中..."):
                            batch_data = []
                            rows_to_delete = []
                            for u in collected_updates:
                                if u["orig_idx"] != -1:
                                    if u["delete"]:
                                        rows_to_delete.append(u["orig_idx"])
                                    else:
                                        batch_data.append({'range': f"E{u['orig_idx']}", 'values': [[e_mail]]})
                                        batch_data.append({'range': f"F{u['orig_idx']}", 'values': [[e_ext]]})
                                        batch_data.append({'range': f"H{u['orig_idx']}", 'values': [[u["qty"]]]})
                                        
                            if batch_data: batch_update_safe(ws_inv, batch_data)
                            for r_idx in sorted(rows_to_delete, reverse=True):
                                safe_delete_row(ws_inv, r_idx)
                                
                            st.success("✅ 實驗室資料與庫存已成功更新！")
                            st.session_state.reset_key += 1
                            time.sleep(1); load_data.clear(); st.rerun()

if __name__ == "__main__":
    main()