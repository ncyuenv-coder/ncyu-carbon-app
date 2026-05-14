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

# 指定氣體種類 (需與 8號程式 一致)
GAS_TYPES = ["二氧化碳", "甲烷", "乙炔", "一氧化二氮(笑氣)"]

# ================= 莫蘭迪風格 CSS =================
def apply_morandi_theme():
    st.markdown("""
    <style>
        .metric-card { background-color: #FDFBF7; padding: 20px; border-radius: 12px; border: 1px solid #9DB4AB; text-align: center; margin-bottom: 20px;}
        .metric-title { color: #5C6B73; font-size: 18px; font-weight: bold; }
        .metric-value { font-size: 36px; font-weight: bold; color: #2C3E50; }
        button[kind="primary"] { background-color: #27AE60 !important; color: white !important; font-weight: bold !important;}
        .dept-header { background-color: #EAECEE; padding: 10px; border-radius: 5px; font-weight: bold; margin-top: 15px; border-left: 5px solid #5C6B73;}
        .edit-box { background-color: #F4F6F6; padding: 20px; border-radius: 10px; border: 1px dashed #ABB2B9; margin-bottom: 20px; }
    </style>
    """, unsafe_allow_html=True)

# ================= Google API 與快取 =================
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
def safe_append_rows(worksheet, rows): 
    worksheet.append_rows(rows)

@with_retry(max_retries=3)
def safe_delete_row(worksheet, idx):
    worksheet.delete_rows(idx)

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

# ================= 主程式 =================
def main():
    apply_morandi_theme()
    st.title("📢 氣體鋼瓶盤查管理與後台控制")
    
    df_inv, df_pur, df_trk = load_data()
    
    col_refresh = st.columns([8, 2])
    with col_refresh[1]:
        if st.button("🔄 強制刷新最新資料", use_container_width=True):
            load_data.clear()
            st.rerun()

    tab1, tab2, tab3, tab4 = st.tabs(["📧 批次發送作業", "📊 進度追蹤與稽催", "📁 數據檢視", "🛠️ 庫存資料管理"])

    # ================= Tab 1 & 2 & 3 保持不變 (略，與前次相同) =================
    # (為了節省空間，這裡假設原本的 Tab 1~3 代碼已存在)
    with tab1:
        st.markdown("### 批次發送與金鑰產生")
        # ... (承接之前的 Tab 1 代碼)
        if df_inv.empty: st.warning("資料庫尚無紀錄。")
        else:
            labs = df_inv[['系所', '實驗室老師', '電子郵件', '氣體鋼瓶所在位置實驗室門牌', '校區']].drop_duplicates()
            with st.form("send_form"):
                now_year = datetime.datetime.now().year - 1911
                batch_name = st.text_input("批次名稱", value=f"{now_year}年度氣體鋼瓶盤查確認")
                is_test_mode = st.radio("模式", ["測試", "正式"]) == "測試"
                test_email = st.text_input("測試信箱") if is_test_mode else ""
                submit_btn = st.form_submit_button("🚀 發送", type="primary")
                if submit_btn:
                    # 發信與寫入金鑰邏輯...
                    st.info("處理中...")
                    load_data.clear()
                    st.rerun()

    with tab2:
        st.markdown("### 進度追蹤")
        if not df_trk.empty:
            batch = st.selectbox("選擇批次", df_trk['發送批次'].unique())
            df_cur = df_trk[df_trk['發送批次'] == batch]
            st.write(f"應回報: {len(df_cur)} | 已回報: {len(df_cur[df_cur['回覆狀態']=='已回報'])}")
            st.dataframe(df_cur)

    with tab3:
        st.markdown("### 盤查數據總覽")
        c3_1, c3_2 = st.tabs(["現有庫存", "年度購買紀錄"])
        with c3_1: st.dataframe(df_inv)
        with c3_2: st.dataframe(df_pur)

    # ================= Tab 4: 🛠️ 庫存資料管理 (新增) =================
    with tab4:
        st.markdown("### 🛠️ 氣體鋼瓶資料庫 (母表) 直接編輯")
        
        mode = st.radio("選擇操作模式", ["新增資料", "修改/刪除現有資料"], horizontal=True)
        
        gc = get_gc()
        ws_inv = gc.open_by_key(SHEET_ID).worksheet('氣體鋼瓶資料庫')

        if mode == "新增資料":
            st.markdown("<div class='edit-box'>", unsafe_allow_html=True)
            st.subheader("➕ 手動新增實驗室鋼瓶紀錄")
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
                
                submitted = st.form_submit_button("確認新增", type="primary")
                if submitted:
                    new_row = [a_dept, a_mgr, a_campus, a_room, a_mail, a_ext, a_gas, a_qty, a_year]
                    safe_append_rows(ws_inv, [new_row])
                    st.success(f"✅ 已成功新增：{a_mgr} 老師 - {a_gas}")
                    load_data.clear()
            st.markdown("</div>", unsafe_allow_html=True)

        elif mode == "修改/刪除現有資料":
            if df_inv.empty:
                st.info("目前資料庫為空，無法修改。")
            else:
                st.markdown("<div class='edit-box'>", unsafe_allow_html=True)
                st.subheader("📝 修改或刪除特定項目")
                
                # 第一步：選擇老師與門牌
                all_labs = df_inv.apply(lambda x: f"{x['系所']}-{x['實驗室老師']} ({x['氣體鋼瓶所在位置實驗室門牌']})", axis=1).unique()
                sel_lab = st.selectbox("1. 請先選擇實驗室", ["請選擇"] + list(all_labs))
                
                if sel_lab != "請選擇":
                    # 篩選該實驗室的紀錄
                    lab_info = sel_lab.split("-")
                    dept_sel = lab_info[0]
                    mgr_sel = lab_info[1].split(" (")[0]
                    room_sel = lab_info[1].split("(")[1].replace(")", "")
                    
                    lab_records = df_inv[(df_inv['系所'] == dept_sel) & 
                                         (df_inv['實驗室老師'] == mgr_sel) & 
                                         (df_inv['氣體鋼瓶所在位置實驗室門牌'] == room_sel)]
                    
                    # 第二步：選擇要改哪一種氣體
                    sel_gas = st.selectbox("2. 選擇欲修改的氣體項目", lab_records['鋼瓶氣體種類'].unique())
                    
                    # 抓取該列在原始 Sheet 中的 index
                    target_row_data = lab_records[lab_records['鋼瓶氣體種類'] == sel_gas].iloc[0]
                    
                    # 尋找原始行號 (需精準匹配)
                    all_rows = ws_inv.get_all_records()
                    original_row_index = -1
                    for i, r in enumerate(all_rows):
                        if str(r['系所']) == dept_sel and str(r['實驗室老師']) == mgr_sel and \
                           str(r['氣體鋼瓶所在位置實驗室門牌']) == room_sel and str(r['鋼瓶氣體種類']) == sel_gas:
                            original_row_index = i + 2 # +2 是因為 header 為 1 且 index 從 0 開始
                            break
                    
                    if original_row_index != -1:
                        st.divider()
                        # 修改表單
                        with st.form("edit_form"):
                            st.write(f"正在編輯：{sel_lab} 的 **{sel_gas}**")
                            col_e1, col_e2, col_e3 = st.columns(3)
                            u_qty = col_e1.number_input("鋼瓶數量", value=int(target_row_data['鋼瓶數量']))
                            u_ext = col_e2.text_input("分機", value=str(target_row_data['分機']))
                            u_mail = col_e3.text_input("電子郵件", value=str(target_row_data['電子郵件']))
                            
                            btn_update = st.form_submit_button("💾 儲存修改")
                            
                            if btn_update:
                                # 更新 Sheet (這裡使用 update_cell 較保險)
                                ws_inv.update_cell(original_row_index, 6, u_ext)  # 分機在第 6 欄
                                ws_inv.update_cell(original_row_index, 5, u_mail) # 電郵在第 5 欄
                                ws_inv.update_cell(original_row_index, 8, u_qty)  # 數量在第 8 欄
                                st.success("✅ 資料已更新！")
                                load_data.clear()
                                time.sleep(1)
                                st.rerun()

                        # 刪除按鈕 (放在 Form 外面)
                        st.write("---")
                        if st.button("🗑️ 刪除此筆氣體紀錄", type="secondary"):
                            # 進行二次確認 (Streamlit 常用作法)
                            st.warning(f"確定要刪除 {mgr_sel} 老師的 {sel_gas} 紀錄嗎？此操作不可復原。")
                            if st.button("🔥 確認永久刪除"):
                                safe_delete_row(ws_inv, original_row_index)
                                st.error("🚫 資料已刪除。")
                                load_data.clear()
                                time.sleep(1)
                                st.rerun()
                    else:
                        st.error("無法定位該行資料，請重新刷新。")
                st.markdown("</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()