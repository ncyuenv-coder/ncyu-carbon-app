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

# ================= 系統與全域參數設定 =================
st.set_page_config(page_title="氣體鋼瓶管理與追蹤", page_icon="📢", layout="wide")

SHEET_ID = '1Hw4rXo4ww7O9YXTwoUJeWioO5ZzM_bivRcLLpOl26DY'

# ⚠️ 部署 8號程式 後，請將下方網址替換為實際的 Streamlit App 網址
BASE_FORM_URL = "https://ncyu-gas-cylinder-report.streamlit.app/"

GAS_TYPES = ["二氧化碳", "甲烷", "乙炔", "一氧化二氮(笑氣)"]

# ================= 莫蘭迪風格 CSS (精準對齊 11_📢 毒化物後台) =================
def apply_morandi_theme():
    st.markdown("""
    <style>
        /* 數據儀表板 Metric Cards */
        .metric-card { background-color: #FDFBF7; padding: 0; border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); text-align: center; overflow: hidden; border: 1px solid #9DB4AB; display: flex; flex-direction: column; }
        .metric-title { background-color: #4A5D6E; color: #FFFFFF; padding: 12px 15px; font-size: 18px; font-weight: bold; margin: 0; letter-spacing: 1px; }
        .metric-value { font-size: 54px; font-weight: bold; color: #000000 !important; padding: 15px; margin: 0; flex: 1; display: flex; align-items: baseline; justify-content: center; gap: 8px; }
        .metric-value span { font-size: 22px !important; color: #5C6B73 !important; font-weight: normal; }
        
        /* 沉浸式頁籤 Tabs */
        div[data-baseweb="tab-list"] button { background-color: #5C6B73 !important; border-radius: 8px 8px 0 0 !important; padding: 10px 25px !important; border: none !important; opacity: 0.8; }
        div[data-baseweb="tab-list"] button p { font-size: 20px !important; color: #FFFFFF !important; font-weight: 600 !important; }
        div[data-baseweb="tab-list"] button[aria-selected="true"] { background-color: #3F4E4F !important; opacity: 1; border-bottom: 3px solid #9DB4AB !important; }
        
        /* 按鈕視覺 */
        button[kind="primary"] { background-color: #27AE60 !important; color: #FFFFFF !important; border: none !important; font-weight: bold !important; font-size: 18px !important; padding: 10px 20px !important;}
        button[kind="primary"]:hover { background-color: #1E8449 !important; }
        button:has(div:contains("🗑️")) { background-color: #E74C3C !important; color: #FFFFFF !important; border: none !important; font-weight: bold !important;}
        button:has(div:contains("🗑️")):hover { background-color: #C0392B !important; }
        
        /* 客製化 Radio 單選按鈕 */
        div.stRadio > div[role="radiogroup"] > label {
            background-color: #E3F2FD !important; padding: 10px 20px !important; border-radius: 8px !important;
            margin-right: 15px !important; border: 2px solid #BBDEFB !important; box-shadow: 0 2px 4px rgba(0,0,0,0.05); cursor: pointer;
        }
        div.stRadio > div[role="radiogroup"] > label p { color: #0D47A1 !important; font-size: 18px !important; font-weight: bold !important; }
        div.stRadio > div[role="radiogroup"] > label[data-checked="true"] { background-color: #BBDEFB !important; border-color: #1976D2 !important; }

        /* 黃色高光折疊面板 Expander */
        div.element-container:has(.yellow-marker) + div.element-container div[data-testid="stExpander"] details summary {
            background-color: #FFF3CD !important; border-left: 5px solid #F1C40F !important; border-radius: 6px !important;
        }
        div.element-container:has(.yellow-marker) + div.element-container div[data-testid="stExpander"] details summary p {
            color: #856404 !important; font-weight: bold !important; font-size: 18px !important;
        }
        
        /* 編輯區塊設計 */
        .edit-box { background-color: #F8F9FA; padding: 25px; border-radius: 12px; border: 1px dashed #ABB2B9; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.02); }
    </style>
    """, unsafe_allow_html=True)

# ================= Google API 與快取防呆 =================
def with_retry(max_retries=3, base_delay=1.5):
    def decorator(func):
        def wrapper(*args, **kwargs):
            retries = 0
            while True:
                try: return func(*args, **kwargs)
                except Exception as e:
                    if retries >= max_retries: raise e
                    time.sleep(base_delay * (2 ** retries))
                    retries += 1
        return wrapper
    return decorator

@st.cache_resource
def get_credentials():
    oauth_secrets = st.secrets["gcp_oauth"]
    return Credentials(token=None, refresh_token=oauth_secrets["refresh_token"], client_id=oauth_secrets["client_id"], client_secret=oauth_secrets["client_secret"], token_uri="https://oauth2.googleapis.com/token")

def get_gc(): return gspread.authorize(get_credentials())

@st.cache_data(ttl=60, show_spinner=False)
def load_data():
    gc = get_gc()
    sh = gc.open_by_key(SHEET_ID)
    try: df_inv = pd.DataFrame(sh.worksheet('氣體鋼瓶資料庫').get_all_records())
    except: df_inv = pd.DataFrame()
    try: df_pur = pd.DataFrame(sh.worksheet('氣體鋼瓶年度使用紀錄').get_all_records())
    except: df_pur = pd.DataFrame()
    try: df_trk = pd.DataFrame(sh.worksheet('發送紀錄與金鑰').get_all_records())
    except: df_trk = pd.DataFrame()
    return df_inv, df_pur, df_trk

@with_retry(max_retries=3)
def safe_append_rows(worksheet, rows): worksheet.append_rows(rows)

@with_retry(max_retries=3)
def safe_delete_row(worksheet, idx): worksheet.delete_rows(idx)

# ================= 發信引擎 =================
def send_email_action(to_email, subject, html_body):
    try:
        sender_email = st.secrets["email"]["sender_email"]
        app_password = st.secrets["email"]["app_password"]
        msg = MIMEMultipart('alternative')
        msg['Subject'] = subject
        msg['From'] = sender_email
        msg['To'] = to_email
        msg.attach(MIMEText(html_body, 'html'))
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(sender_email, app_password)
        server.send_message(msg)
        server.quit()
        return True
    except: return False

def generate_styled_email_html(email_body_text, title="國立嘉義大學 氣體鋼瓶盤查確認"):
    html_content = email_body_text.replace("\n", "<br>")
    return f"""
    <html><body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333; font-size: 15px;">
        <div style="max-width: 600px; margin: 0 auto; border: 1px solid #ddd; border-radius: 10px; overflow: hidden;">
            <div style="background-color: #5C6B73; color: white; padding: 15px 20px; text-align: center;"><h2 style="margin: 0;">{title}</h2></div>
            <div style="padding: 20px;">{html_content}</div>
        </div>
    </body></html>
    """

# ================= 主程式 =================
def main():
    apply_morandi_theme()
    st.title("📢 氣體鋼瓶年度盤查與追蹤管理")
    
    df_inv, df_pur, df_trk = load_data()
    
    col_refresh = st.columns([8, 2])
    with col_refresh[1]:
        if st.button("🔄 強制刷新最新資料", use_container_width=True):
            load_data.clear()
            st.rerun()

    tab1, tab2, tab3, tab4 = st.tabs(["📧 批次發送作業", "📊 回報進度追蹤與稽催", "📁 最新盤查數據檢視", "🛠️ 庫存資料管理"])

    # ================= Tab 1: 批次發送作業 =================
    with tab1:
        st.markdown("### 🔄 建立新一期盤查通知")
        if df_inv.empty: 
            st.warning("⚠️ 氣體鋼瓶資料庫尚無任何紀錄，無法執行發送作業。")
        else:
            labs = df_inv[['系所', '實驗室老師', '電子郵件', '氣體鋼瓶所在位置實驗室門牌', '校區']].drop_duplicates()
            total_teachers = labs['實驗室老師'].nunique()
            total_labs = len(labs)
            
            st.markdown(f"""
            <div style="display: flex; gap: 20px; margin-bottom: 25px;">
                <div class="metric-card" style="flex: 1;"><div class="metric-title">資料庫涵蓋老師總數</div><div class="metric-value">{total_teachers} <span>位</span></div></div>
                <div class="metric-card" style="flex: 1;"><div class="metric-title">資料庫涵蓋實驗室總數</div><div class="metric-value">{total_labs} <span>間</span></div></div>
            </div>
            """, unsafe_allow_html=True)

            with st.form("send_form"):
                now_year = datetime.datetime.now().year - 1911
                batch_name = st.text_input("📝 請輸入本次發送批次名稱", value=f"{now_year}年度氣體鋼瓶盤查確認")
                
                st.markdown("#### 🛡️ 發送模式設定")
                is_test_mode = st.radio("寄送模式選擇", ["🧪 測試寄信模式", "🚀 正式寄信模式"], horizontal=True) == "🧪 測試寄信模式"
                test_email = st.text_input("📩 測試接收信箱 (測試模式下所有信件將寄至此信箱)") if is_test_mode else ""
                
                st.markdown("<div class='yellow-marker'></div>", unsafe_allow_html=True)
                with st.expander("📝 點此編輯通知信件內容範本", expanded=False):
                    st.info("💡 下方欄位中的 `{批次名稱}`, `{系所}`, `{老師名稱}`, `{links}` 為系統保留字，請保留這些設定。")
                    batch_subject = st.text_input("信件主旨", value="【通知】國立嘉義大學 {批次名稱}")
                    b_default_body = """{系所} {老師名稱} 您好：

為配合溫室氣體盤查作業，請協助檢視與更新貴實驗室之氣體鋼瓶庫存，並上傳本年度之購買佐證單據。

📍 您所管理的實驗室專屬確認連結如下：
{links}

⚠️ 注意：若您管理多間實驗室，請逐一點擊上方各門牌專屬連結完成回報。
本信件由系統自動發送，請勿直接回覆。"""
                    batch_body = st.text_area("信件內容", value=b_default_body, height=250)
                
                st.markdown("<br>", unsafe_allow_html=True)
                submit_btn = st.form_submit_button("🚀 產生專屬金鑰並發送通知信", type="primary")
                
            if submit_btn:
                if is_test_mode and not test_email:
                    st.error("測試模式下，請務必輸入測試接收信箱！")
                else:
                    my_bar = st.progress(0, text="系統處理中...")
                    grouped = labs.groupby(['系所', '實驗室老師', '電子郵件'])
                    dict_records_to_append = []
                    success_count, current_idx = 0, 0
                    total_groups = len(grouped)
                    now_time = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
                    
                    for (dept, mgr, email), group in grouped:
                        display_mgr = f"{mgr} 老師" if "老師" not in mgr else mgr
                        target_email = test_email if is_test_mode else email
                        links_html = ""
                        
                        for _, row in group.iterrows():
                            token = uuid.uuid4().hex
                            room, campus = row["氣體鋼瓶所在位置實驗室門牌"], row["校區"]
                            link = f"{BASE_FORM_URL}?token={token}"
                            links_html += f'<div style="margin-bottom: 8px; background-color: #FDFBF7; padding: 10px; border-left: 5px solid #D4A373;"><b>{campus} - 門牌 {room}</b>：<a href="{link}" style="color: #3498DB; font-weight: bold;">點此進入盤查系統</a></div>'
                            dict_records_to_append.append([batch_name, dept, mgr, campus, room, email, token, link, "發送成功", now_time, "待回覆", ""])
                            
                        mail_content = batch_body.replace("{系所}", dept).replace("{老師名稱}", display_mgr).replace("{批次名稱}", batch_name).replace("{links}", links_html)
                        html_wrap = generate_styled_email_html(mail_content, title=f"【通知】{batch_name}")
                        subject_content = batch_subject.replace("{批次名稱}", batch_name).replace("{系所}", dept).replace("{老師名稱}", display_mgr)
                        
                        if send_email_action(target_email, subject_content, html_wrap): success_count += 1
                        current_idx += 1
                        my_bar.progress(current_idx / total_groups, text=f"發送信件中... ({current_idx}/{total_groups})")
                    
                    if dict_records_to_append:
                        safe_append_rows(get_gc().open_by_key(SHEET_ID).worksheet('發送紀錄與金鑰'), dict_records_to_append)
                    st.success(f"✅ 作業完成！共成功寄出 {success_count} 封通知信。")
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
            
            st.markdown("### ✉️ 稽催通知設定與發送")
            follow_mode = st.radio("稽催發送模式", ["🧪 測試寄信模式", "🚀 正式寄信模式"], horizontal=True) == "🧪 測試寄信模式"
            f_test_email = st.text_input("📩 稽催測試接收信箱") if follow_mode else ""
            
            if not pending.empty:
                st.markdown("### 📋 未回報名單 (依系所)")
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
                            if st.button(f"✉️ 一鍵稽催 {dept}", key=f"f_{dept}"):
                                if follow_mode and not f_test_email: st.error("請輸入測試信箱！")
                                else:
                                    f_count = 0
                                    for (mgr, email), group in grouped_pending:
                                        target = f_test_email if follow_mode else email
                                        display_mgr = f"{mgr} 老師" if "老師" not in mgr else mgr
                                        links_html = "".join([f'<div style="margin-bottom: 8px; background-color: #FDFBF7; padding: 10px; border-left: 5px solid #F39C12;">門牌 <b>{r["氣體鋼瓶所在位置實驗室門牌"]}</b>：<a href="{r["專屬填報網址"]}" style="color: #3498DB; font-weight: bold;">點此補登</a></div>' for _, r in group.iterrows()])
                                        msg_body = f"{dept} {display_mgr} 您好：<br><br>提醒您，系統尚未收到您的氣體鋼瓶盤查資料。敬請撥冗點擊下方連結補登：<br><br>{links_html}"
                                        html_wrap = generate_styled_email_html(msg_body, title="盤查未完成 稽催提醒")
                                        if send_email_action(target, f"【稽催提醒】{sel_batch} 尚未完成", html_wrap): f_count += 1
                                    st.success(f"✅ {dept} 稽催發送完成 ({f_count} 封信)。")
            else:
                st.success("🎉 太棒了！此批次所有實驗室皆已完成回報！")

    # ================= Tab 3: 最新盤查數據檢視 =================
    with tab3:
        st.markdown("### 📁 氣體鋼瓶資料庫 (現有庫存總表)")
        st.info("💡 此處顯示前台老師送出後，系統自動覆寫的最新庫存資料。")
        if not df_inv.empty: st.dataframe(df_inv, use_container_width=True, height=250)
        else: st.write("目前尚無資料。")
            
        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown("### 🧾 年度使用紀錄 (碳盤查活動數據)")
        st.info("💡 顯示本年度申報的購買紀錄，可直接點擊右側連結查看原始憑證。")
        if not df_pur.empty:
            st.data_editor(df_pur, column_config={"購買單據連結": st.column_config.LinkColumn("購買單據連結 (點擊開啟)")}, hide_index=True, use_container_width=True)
        else: st.write("目前尚無資料。")

    # ================= Tab 4: 🛠️ 庫存資料管理 =================
    with tab4:
        st.markdown("### 🛠️ 氣體鋼瓶資料庫 (母表) 直接編輯")
        mode = st.radio("選擇操作模式", ["新增實驗室鋼瓶", "修改/刪除現有紀錄"], horizontal=True)
        gc = get_gc()
        ws_inv = gc.open_by_key(SHEET_ID).worksheet('氣體鋼瓶資料庫')

        if mode == "新增實驗室鋼瓶":
            st.markdown("<div class='edit-box'>", unsafe_allow_html=True)
            st.markdown("#### ➕ 手動新增氣體鋼瓶紀錄")
            with st.form("add_inv_form", clear_on_submit=True):
                col1, col2, col3 = st.columns(3)
                a_dept = col1.text_input("系所")
                a_mgr = col1.text_input("實驗室老師")
                a_campus = col2.selectbox("校區", ["蘭潭校區", "民雄校區", "新民校區"])
                a_room = col2.text_input("實驗室門牌")
                a_mail = col3.text_input("電子郵件")
                a_ext = col3.text_input("分機")
                
                col4, col5, col6 = st.columns(3)
                a_gas = col4.selectbox("氣體種類", GAS_TYPES)
                a_qty = col5.number_input("鋼瓶數量", min_value=1, step=1)
                a_year = col6.text_input("建檔年度", value=f"{datetime.datetime.now().year-1911}年")
                
                if st.form_submit_button("🚀 確認新增", type="primary"):
                    if a_dept and a_mgr and a_room:
                        safe_append_rows(ws_inv, [[a_dept, a_mgr, a_campus, a_room, a_mail, a_ext, a_gas, a_qty, a_year]])
                        st.success(f"✅ 已成功新增：{a_mgr} 老師 - {a_gas}")
                        time.sleep(1); load_data.clear(); st.rerun()
                    else: st.error("請至少填寫系所、老師與門牌！")
            st.markdown("</div>", unsafe_allow_html=True)

        elif mode == "修改/刪除現有紀錄":
            if df_inv.empty: st.info("資料庫目前為空。")
            else:
                st.markdown("<div class='edit-box'>", unsafe_allow_html=True)
                st.markdown("#### 📝 定位並編輯特定項目")
                all_labs = df_inv.apply(lambda x: f"{x['系所']}-{x['實驗室老師']} ({x['氣體鋼瓶所在位置實驗室門牌']})", axis=1).unique()
                sel_lab = st.selectbox("📌 1. 請先選擇實驗室", ["請選擇"] + list(all_labs))
                
                if sel_lab != "請選擇":
                    dept_sel, mgr_sel = sel_lab.split("-")[0], sel_lab.split("-")[1].split(" (")[0]
                    room_sel = sel_lab.split("(")[1].replace(")", "")
                    
                    lab_records = df_inv[(df_inv['系所'] == dept_sel) & (df_inv['實驗室老師'] == mgr_sel) & (df_inv['氣體鋼瓶所在位置實驗室門牌'] == room_sel)]
                    sel_gas = st.selectbox("📌 2. 選擇欲修改的氣體", lab_records['鋼瓶氣體種類'].unique())
                    
                    target_row = lab_records[lab_records['鋼瓶氣體種類'] == sel_gas].iloc[0]
                    all_rows = ws_inv.get_all_records()
                    orig_idx = next((i + 2 for i, r in enumerate(all_rows) if str(r['系所'])==dept_sel and str(r['實驗室老師'])==mgr_sel and str(r['氣體鋼瓶所在位置實驗室門牌'])==room_sel and str(r['鋼瓶氣體種類'])==sel_gas), -1)
                    
                    if orig_idx != -1:
                        st.markdown("<hr style='border:1px dashed #ABB2B9;'>", unsafe_allow_html=True)
                        with st.form("edit_form"):
                            st.write(f"正在編輯：**{mgr_sel}** 老師的 **{sel_gas}**")
                            ce1, ce2, ce3 = st.columns(3)
                            u_qty = ce1.number_input("鋼瓶數量", value=int(target_row['鋼瓶數量']))
                            u_ext = ce2.text_input("分機", value=str(target_row['分機']))
                            u_mail = ce3.text_input("電子郵件", value=str(target_row['電子郵件']))
                            
                            if st.form_submit_button("💾 儲存修改"):
                                ws_inv.update_cell(orig_idx, 6, u_ext)
                                ws_inv.update_cell(orig_idx, 5, u_mail)
                                ws_inv.update_cell(orig_idx, 8, u_qty)
                                st.success("✅ 資料已更新！")
                                time.sleep(1); load_data.clear(); st.rerun()

                        st.markdown("<br>", unsafe_allow_html=True)
                        if st.button("🗑️ 刪除此筆氣體紀錄"):
                            safe_delete_row(ws_inv, orig_idx)
                            st.error("🚫 資料已刪除。")
                            time.sleep(1); load_data.clear(); st.rerun()
                st.markdown("</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()