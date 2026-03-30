import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import gspread
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import streamlit_authenticator as stauth
import time
import uuid
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ==========================================
# 0. 系統設定 (微觀查核後台)
# ==========================================
st.set_page_config(page_title="燃油資料檢視確認", page_icon="🔍", layout="wide")

# ==========================================
# 1. CSS 樣式表
# ==========================================
st.markdown("""
<style>
    :root {
        color-scheme: light;
        --orange-bg: #E67E22;     
        --orange-dark: #D35400;
        --text-main: #2C3E50;
        --text-sub: #566573;
        --morandi-blue: #34495E;
    }
    [data-testid="stAppViewContainer"] { background-color: #F8F9F9; color: var(--text-main); }
    [data-testid="stHeader"] { background-color: rgba(0,0,0,0); }
    [data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 1px solid #BDC3C7; }
    
    div[data-baseweb="input"] > div, div[data-baseweb="base-input"] > input, textarea, input { background-color: #FFFFFF !important; border-color: #BDC3C7 !important; color: #000000 !important; font-size: 1.15rem !important; }
    div[data-baseweb="select"] > div { border-color: #BDC3C7 !important; background-color: #FFFFFF !important; }
    
    button[kind="primary"] { background-color: var(--orange-bg) !important; color: #FFFFFF !important; border: 2px solid var(--orange-dark) !important; border-radius: 12px !important; font-size: 1.3rem !important; font-weight: 800 !important; padding: 0.7rem 1.5rem !important; box-shadow: 0 4px 6px rgba(230, 126, 34, 0.3) !important; width: 100%; }
    button[kind="primary"] p { color: #FFFFFF !important; } 
    button[kind="primary"]:hover { background-color: var(--orange-dark) !important; transform: translateY(-2px) !important; }
    
    button[data-baseweb="tab"] div p { font-size: 1.3rem !important; font-weight: 900 !important; color: var(--text-sub); }
    button[data-baseweb="tab"][aria-selected="true"] div p { color: #E67E22 !important; border-bottom: 3px solid #E67E22; }
    
    [data-testid="stDataFrame"] { font-size: 1.15rem !important; }
    
    .unreported-block { padding: 15px 20px; border-radius: 12px; margin-bottom: 20px; color: #2C3E50; box-shadow: 0 2px 6px rgba(0,0,0,0.08); border: 1px solid rgba(0,0,0,0.05); }
    .unreported-title { font-size: 1.6rem; font-weight: 900; margin-bottom: 12px; border-bottom: 2px solid rgba(0,0,0,0.1); padding-bottom: 8px; }
    
    .dashboard-main-title { font-size: 2.4rem; font-weight: 900; text-align: center; color: #1A5276; margin-bottom: 25px; margin-top: 10px; letter-spacing: 2px; }
    
    .stRadio div[role="radiogroup"] label { background-color: #D6EAF8 !important; border: 1px solid #AED6F1 !important; border-radius: 8px !important; padding: 8px 15px !important; margin-right: 10px !important; }
    .stRadio div[role="radiogroup"] label p { font-size: 1.25rem !important; font-weight: 800 !important; color: #1A5276 !important; }
    .stRadio div[role="radiogroup"] label[data-checked="true"] { background-color: #1A5276 !important; border-color: #1A5276 !important; }
    .stRadio div[role="radiogroup"] label[data-checked="true"] p { color: #FFFFFF !important; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. 身份驗證與初始化
# ==========================================
def clean_secrets(obj):
    if isinstance(obj, dict) or "AttrDict" in str(type(obj)): return {k: clean_secrets(v) for k, v in obj.items()}
    elif isinstance(obj, list): return [clean_secrets(i) for i in obj]
    return obj

if 'unreported_df' not in st.session_state: st.session_state['unreported_df'] = None

if st.session_state.get("authentication_status") is not True:
    st.warning("🔒 請先至首頁 (Hello) 登入系統")
    st.stop()

username = st.session_state.get("username")
name = st.session_state.get("name")

if username != 'admin':
    st.error("🚫 權限不足：此頁面僅供系統管理員使用。")
    st.stop()

try:
    _raw_creds = st.secrets["credentials"]
    credentials_login = clean_secrets(_raw_creds)
    cookie_cfg = st.secrets["cookie"]
    authenticator = stauth.Authenticate(credentials_login, cookie_cfg["name"], cookie_cfg["key"], cookie_cfg["expiry_days"])
except: pass

with st.sidebar:
    st.header(f"👤 {name} (查核員)")
    st.success("☁️ 查核後台連線正常")
    authenticator.logout('登出系統', 'sidebar')

# ==========================================
# 3. 資料庫連線與常數
# ==========================================
SHEET_ID = "1gqDU21YJeBoBOd8rMYzwwZ45offXWPGEODKTF6B8k-Y" 
UNREPORTED_COLORS = ["#D5DBDB", "#FAD7A0", "#D2B4DE", "#AED6F1", "#A3E4D7", "#F5B7B1"]

@st.cache_resource
def init_google_fuel():
    oauth = st.secrets["gcp_oauth"]
    creds = Credentials(token=None, refresh_token=oauth["refresh_token"], token_uri="https://oauth2.googleapis.com/token", client_id=oauth["client_id"], client_secret=oauth["client_secret"], scopes=["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/spreadsheets"])
    gc = gspread.authorize(creds); drive = build('drive', 'v3', credentials=creds)
    return gc, drive

try:
    gc_conn, drive_service = init_google_fuel()
except Exception as e: 
    st.error(f"連線失敗: {e}")
    st.stop()

@st.cache_data(ttl=300) # 更短的快取時間，適合頻繁異動的查核頁面
def load_fuel_data():
    sh = gc_conn.open_by_key(SHEET_ID)
    try: ws_equip = sh.worksheet("設備清單") 
    except: ws_equip = sh.sheet1 
    try: ws_record = sh.worksheet("油料填報紀錄")
    except: ws_record = sh.worksheet("填報紀錄")
    
    df_e = pd.DataFrame(ws_equip.get_all_records()).astype(str)
    df_e['設備數量_num'] = pd.to_numeric(df_e['設備數量'], errors='coerce').fillna(1)
    
    data = ws_record.get_all_values()
    df_r = pd.DataFrame(data[1:], columns=data[0]) if len(data) > 1 else pd.DataFrame(columns=data[0])
    return df_e, df_r

# ==========================================
# 4. 輔助函數 (Email 發信模組)
# ==========================================
def send_system_email(to_email, subject, body):
    try:
        smtp_cfg = st.secrets.get("smtp", {})
        if not smtp_cfg: 
            return False, "系統尚未設定 SMTP 參數 (st.secrets['smtp'])"
        if not str(to_email).strip() or str(to_email).strip().lower() in ['nan', 'none', '']: 
            return False, "無效的電子郵件地址"
        
        msg = MIMEMultipart()
        msg['From'] = smtp_cfg.get("email", "noreply@system.com")
        msg['To'] = str(to_email).strip()
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'html'))
        
        server = smtplib.SMTP(smtp_cfg["server"], smtp_cfg["port"])
        server.starttls()
        server.login(smtp_cfg["email"], smtp_cfg["password"])
        server.send_message(msg)
        server.quit()
        return True, "發送成功"
    except Exception as e:
        return False, f"連線或驗證失敗: {str(e)}"

# ==========================================
# 5. 微觀查核分頁 Fragment 模組
# ==========================================

@st.fragment
def render_tab1_missing(df_clean, df_equip_full, all_years):
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<h3 style='color: #2C3E50;'>⚙️ 請選擇作業模式：</h3>", unsafe_allow_html=True)
    mode = st.radio("請選擇作業模式：", ["📅 每月申報提醒通知", "🔍 篩選未申報名單催報"], horizontal=True, label_visibility="collapsed")
    st.markdown("---")
    
    if "每月申報提醒" in mode:
        st.markdown("#### 📅 批次寄送每月申報提醒通知")
        st.info("此功能將會向資料庫內所有設備單位寄送當月例行性申報提醒通知。")
        
        c_y, c_m = st.columns(2)
        cur_year = datetime.now().year
        cur_month = datetime.now().month
        
        sel_year = c_y.number_input("通知年度", value=cur_year, min_value=2020, max_value=2100, step=1)
        sel_month = c_m.selectbox("通知月份", range(1, 13), index=cur_month-1)
        
        if '設備檢視年度' in df_equip_full.columns:
            df_equip = df_equip_full[df_equip_full['設備檢視年度'].astype(str) == str(sel_year)].copy()
        else:
            df_equip = df_equip_full.copy()

        if st.button(f"🔔 寄送 {sel_year}年{sel_month}月申報提醒通知", use_container_width=True):
            if '電子郵件' not in df_equip.columns:
                st.error("❌ 找不到「電子郵件」欄位，請確認 Sheet1 的 K 欄已正確建檔。")
            else:
                with st.spinner("正在批次發送每月提醒信件..."):
                    groups = df_equip.groupby(['填報單位', '電子郵件'])
                    success_count, fail_count = 0, 0
                    error_logs = [] 
                    
                    subject = f"【提醒】{sel_year}年{sel_month}月份燃油設備油料使用申報通知，請於次月5日前完成申報"
                    
                    for (unit, email), group in groups:
                        eq_li = "".join([f"<li>{r['設備名稱備註']} ：{r.get('設備數量','1')}台，保管人：{r.get('保管人','-')}</li>" for _, r in group.iterrows()])
                        body = f"""
                        <p>您好：</p>
                        <p>依據環境部「事業應盤查登錄溫室氣體排放量之排放源」及「溫室氣體排放量盤查登錄及查驗管理辦法」，本校須依辦理溫室氣體排放量盤查登錄作業。</p>
                        <p>提醒您，請於每月5日前(遇假日順延至下一工作日)至校內溫室氣體盤查填報系統(連結：<a href="https://ncyu-carbon-app-mduue5hffp7uknsskmjet9.streamlit.app/">https://ncyu-carbon-app-mduue5hffp7uknsskmjet9.streamlit.app/</a>)，申報貴單位前一月份之燃油設備油料使用情形。</p>
                        <p>貴單位 ({unit}) 負責之設備清單如下，請撥冗確認並至系統填報：</p>
                        <ul>{eq_li}</ul>
                        <p>若有疑問，請洽環境保護及安全管理中心林小姐 (分機7137)</p>
                        <br>
                        <p>感謝您的配合！<br>
                        環境保護及安全管理中心 敬上<br>
                        <span style="color:#7F8C8D;font-size:12px;">(此為系統自動發送，請勿直接回覆)</span></p>
                        """
                        ok, msg = send_system_email(email, subject, body)
                        if ok: 
                            success_count += 1
                        else: 
                            fail_count += 1
                            if f"[{email}] {msg}" not in error_logs:
                                error_logs.append(f"[{email}] {msg}")
                    
                    if success_count > 0: st.success(f"✅ 成功寄送 {success_count} 封提醒信件。")
                    if fail_count > 0: 
                        st.warning(f"⚠️ 有 {fail_count} 封信件發送失敗。請查看下方錯誤詳情：")
                        with st.expander("展開查看發送失敗原因"):
                            for log in error_logs:
                                st.code(log)
    else:
        st.markdown("#### 🔍 篩選未申報清單並催報")
        c_f1, c_f2 = st.columns(2)
        default_year = all_years[0] if all_years else datetime.now().year
        d_start = c_f1.date_input("查詢起始日", date(default_year, 1, 1), key="t3_d1")
        d_end = c_f2.date_input("查詢結束日", date.today(), key="t3_d2")
        
        target_year = d_end.year
        if '設備檢視年度' in df_equip_full.columns:
            df_equip = df_equip_full[df_equip_full['設備檢視年度'].astype(str) == str(target_year)].copy()
        else:
            df_equip = df_equip_full.copy()

        if st.button("🔍 開始篩選未申報單位", use_container_width=True):
            if not df_clean.empty:
                mask = (df_clean['日期格式'].dt.date >= d_start) & (df_clean['日期格式'].dt.date <= d_end)
                reported_keys = set(df_clean[mask]['填報單位'].astype(str) + "|||" + df_clean[mask]['設備名稱備註'].astype(str))
                df_eq_copy = df_equip.copy()
                df_eq_copy['已申報'] = df_eq_copy.apply(lambda r: (str(r.get('填報單位', '')) + "|||" + str(r.get('設備名稱備註', ''))) in reported_keys, axis=1)
                
                unreported = df_eq_copy[~df_eq_copy['已申報']]
                st.session_state['unreported_df'] = unreported
            else:
                st.warning("無資料可供篩選。")

        st.markdown("---")
        
        if st.session_state.get('unreported_df') is not None:
            unreported = st.session_state['unreported_df']
            if not unreported.empty:
                st.error(f"🚩 期間 [{d_start} ~ {d_end}] 共有 {len(unreported)} 台設備未申報！")
                
                if st.button("📨 寄送催報通知", use_container_width=True):
                    if '電子郵件' not in unreported.columns:
                        st.error("❌ 找不到「電子郵件」欄位，請確認 Sheet1 已正確建檔。")
                    else:
                        with st.spinner("正在發送催報信件..."):
                            groups = unreported.groupby(['填報單位', '電子郵件'])
                            success_count, fail_count = 0, 0
                            error_logs = [] 
                            subject = f"【催報提醒】查貴單位尚未申報({d_start} ~ {d_end})之燃油設備油料使用情形，請於通知日起3日內申報，謝謝"
                            
                            for (unit, email), group in groups:
                                eq_li = "".join([f"<li>{r['設備名稱備註']} ：{r.get('設備數量','1')}台，保管人：{r.get('保管人','-')}</li>" for _, r in group.iterrows()])
                                body = f"""
                                <p>您好：</p>
                                <p>依據環境部「事業應盤查登錄溫室氣體排放量之排放源」及「溫室氣體排放量盤查登錄及查驗管理辦法」，本校須依辦理溫室氣體排放量盤查登錄作業，先予說明。</p>
                                <p>經系統比對，貴單位 ({unit}) 於 {d_start} 至 {d_end} 期間有以下設備尚未完成油料使用申報：</p>
                                <ul>{eq_li}</ul>
                                <p>請於通知日起3日內至校內溫室氣體盤查填報系統(連結：<a href="https://ncyu-carbon-app-mduue5hffp7uknsskmjet9.streamlit.app/">https://ncyu-carbon-app-mduue5hffp7uknsskmjet9.streamlit.app/</a>)，申報貴單位之燃油設備油料使用情形，以免影響全校溫室氣體排放量盤查統計正確性。</p>
                                <p>若有疑問，請洽環境保護及安全管理中心林小姐 (分機7137)</p>
                                <br>
                                <p>感謝您的配合！<br>
                                環境保護及安全管理中心 敬上<br>
                                <span style="color:#7F8C8D;font-size:12px;">(此為系統自動發送，請勿直接回覆)</span></p>
                                """
                                ok, msg = send_system_email(email, subject, body)
                                if ok: 
                                    success_count += 1
                                else: 
                                    fail_count += 1
                                    if f"[{email}] {msg}" not in error_logs:
                                        error_logs.append(f"[{email}] {msg}")
                            
                            if success_count > 0: st.success(f"✅ 成功寄送 {success_count} 封催報信件。")
                            if fail_count > 0: 
                                st.warning(f"⚠️ 有 {fail_count} 封信件發送失敗。請查看下方錯誤詳情：")
                                with st.expander("展開查看發送失敗原因"):
                                    for log in error_logs:
                                        st.code(log)

                for idx, (unit, group) in enumerate(unreported.groupby('填報單位')):
                    bg_color = UNREPORTED_COLORS[idx % len(UNREPORTED_COLORS)]
                    st.markdown(f"""<div class="unreported-block" style="background-color: {bg_color};"><div class="unreported-title">🏢 {unit} (未申報數: {len(group)})</div></div>""", unsafe_allow_html=True)
                    st.dataframe(group[['設備名稱備註', '保管人', '校內財產編號', '電子郵件']], use_container_width=True)
            else: 
                st.success("🎉 太棒了！該期間全數設備皆已完成申報。")


@st.fragment
def render_tab2_edit(df_clean, df_records, all_years):
    st.markdown("<br>", unsafe_allow_html=True)
    selected_admin_year = st.selectbox("📅 請選擇檢視年度", all_years, index=0, key="t4_year")
    st.markdown("---")
    
    st.subheader("🔍 申報資料異動 (目前使用原生編輯器，待後續視覺化升級)")
    df_year = df_clean[df_clean['年份'] == selected_admin_year]
    
    if not df_year.empty:
        df_year['加油日期'] = pd.to_datetime(df_year['加油日期']).dt.date
        edited = st.data_editor(df_year, column_config={"佐證資料": st.column_config.LinkColumn("佐證", display_text="🔗"), "加油日期": st.column_config.DateColumn("日期", format="YYYY-MM-DD"), "加油量": st.column_config.NumberColumn("油量", format="%.2f"), "填報時間": st.column_config.TextColumn("填報時間", disabled=True)}, num_rows="dynamic", use_container_width=True, key="editor_v122")
        
        if st.button("💾 儲存變更", type="primary"):
            try:
                df_all_data = df_records.copy()
                df_all_data['temp_date'] = pd.to_datetime(df_all_data['加油日期'], errors='coerce')
                df_all_data['temp_year'] = df_all_data['temp_date'].dt.year.fillna(0).astype(int)
                df_keep = df_all_data[df_all_data['temp_year'] != selected_admin_year].copy()
                df_new = edited.copy()
                df_final = pd.concat([df_keep, df_new], ignore_index=True)
                if 'temp_date' in df_final.columns: del df_final['temp_date']
                if 'temp_year' in df_final.columns: del df_final['temp_year']
                if '加油日期' in df_final.columns: df_final['加油日期'] = df_final['加油日期'].astype(str)
                df_final = df_final[df_records.columns.tolist()].sort_values(by='加油日期', ascending=False)

                sh_obj = gc_conn.open_by_key(SHEET_ID)
                ws_r = sh_obj.worksheet("油料填報紀錄") if "油料填報紀錄" in [w.title for w in sh_obj.worksheets()] else sh_obj.worksheet("填報紀錄")
                ws_r.clear()
                ws_r.update([df_final.columns.tolist()] + df_final.astype(str).values.tolist())
                
                st.success("✅ 更新成功！資料已安全合併存檔。")
                st.cache_data.clear()
                
                time.sleep(0.5)
                st.rerun()
                
            except Exception as e: st.error(f"更新失敗: {e}")
    else: st.info(f"{selected_admin_year} 年度尚無資料。")

# ==========================================
# 6. 主程式入口
# ==========================================
def main():
    st.markdown('<div class="dashboard-main-title">🔍 燃油資料檢視與確認專區<br><span style="font-size: 1.3rem; color: #5D6D7E; font-weight: 600;">Data Verification & Audit Area</span></div>', unsafe_allow_html=True)
    
    try:
        df_equip, df_records = load_fuel_data()
    except Exception as e:
        st.error(f"資料庫讀取失敗，請重新整理頁面。錯誤: {e}")
        return

    df_clean = df_records.copy()
    if not df_clean.empty:
        df_clean['加油量'] = pd.to_numeric(df_clean['加油量'], errors='coerce').fillna(0)
        df_clean['日期格式'] = pd.to_datetime(df_clean['加油日期'], errors='coerce')
        df_clean['年份'] = df_clean['日期格式'].dt.year.fillna(0).astype(int)

    all_years = sorted(df_clean['年份'][df_clean['年份']>0].unique(), reverse=True) if not df_clean.empty else [datetime.now().year]

    admin_tabs = st.tabs([
        "⚠️ 寄送通知與未申報篩選", 
        "🔍 申報資料異動"
    ])
    
    with admin_tabs[0]: render_tab1_missing(df_clean, df_equip, all_years)
    with admin_tabs[1]: render_tab2_edit(df_clean, df_records, all_years) 

if __name__ == "__main__":
    main()