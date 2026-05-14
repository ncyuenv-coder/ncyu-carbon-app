import streamlit as st
import pandas as pd
import gspread
from google.oauth2.credentials import Credentials
from datetime import datetime, date, timedelta
from dateutil.relativedelta import relativedelta
import time

# =========================================================
# 1. 頁面基本設定 & CSS
# =========================================================
st.set_page_config(page_title="全校電力填報", page_icon="⚡", layout="wide")

# 狀態管理初始化：確保表單重置 Key 跨域穩定
if 'form_reset_key' not in st.session_state:
    st.session_state['form_reset_key'] = 0

# =========================================================
# CSS 美化：包含電力卡片樣式與莫蘭迪深色調頁籤
# =========================================================
st.markdown("""
    <style>
    /* 修復頁首色差，強制透明以顯示主系統的米灰底色 */
    [data-testid="stHeader"] { background-color: transparent !important; }
    .stApp { color: #2C3E50; background-color: #F8F9F9; }
    
    /* 1. 電號資訊卡片 (橫幅滿版) */
    .meter-card {
        background-color: #F8F9F9; 
        border-left: 5px solid #BDC3C7;
        margin-top: 10px;
        margin-bottom: 10px;
        border-radius: 5px;
        box-shadow: 0 1px 2px rgba(0,0,0,0.05);
        overflow: hidden; 
    }
    
    /* 電號資訊卡片標題區 (橫幅滿版淡藍色) */
    .meter-header {
        background-color: #D0EFFF; 
        padding: 10px 15px; 
        margin-bottom: 5px; 
        border-bottom: 1px solid #AED6F1; 
    }

    .meter-name {
        font-size: 1.2rem;
        font-weight: 600; 
        color: #000000 !important;
        display: inline-block;
        margin-right: 15px;
    }
    .meter-id {
        font-size: 1rem;
        color: #566573 !important;
        font-family: monospace;
        font-weight: 500;
        display: inline-block;
    }

    /* 2. 輸入框優化 */
    .stNumberInput, .stDateInput, .stSelectbox {
        padding: 5px 15px;
    }

    .stNumberInput label, .stDateInput label, .stSelectbox label {
        color: #2C3E50 !important;
        font-weight: 600 !important;
        font-size: 0.95rem !important;
    }
    div[data-baseweb="input"] > div, div[data-baseweb="select"] > div {
        background-color: #FFFFFF !important;
        border: 1px solid #D5DBDB !important;
    }

    /* 3. 校區標題 */
    .campus-header {
        padding: 10px 20px;
        border-radius: 8px;
        color: #2C3E50;
        font-weight: 900;
        font-size: 1.3rem;
        margin-top: 25px;
        margin-bottom: 10px;
        display: flex;
        align-items: center;
    }

    /* 4. 確認申報按鈕 (純色莫蘭迪風格) */
    div.stButton > button {
        background-color: #5D6D7E !important; 
        color: white !important;
        font-size: 1.5rem !important;
        font-weight: 800 !important;
        border-radius: 8px !important;
        padding: 12px 0 !important;
        border: none !important;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1) !important;
        transition: all 0.3s ease !important;
        width: 100% !important;
        display: block;
        letter-spacing: 2px;
        margin-top: 20px;
    }
    div.stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 12px rgba(0,0,0,0.15) !important;
        background-color: #34495E !important; 
    }

    /* 5. 申報模式按鈕 (同一列顯示) */
    div[role="radiogroup"] {
        display: flex;
        flex-direction: row;
        gap: 10px;
    }
    div[role="radiogroup"] label {
        background-color: #D6EAF8 !important;
        color: #2E4053 !important;
        border: 1px solid #AED6F1 !important;
        border-radius: 8px !important;
        padding: 8px 15px !important;
        font-size: 1.1rem !important;
        font-weight: 600 !important;
        margin: 0 !important;
        flex: 1;
        text-align: center;
        justify-content: center;
    }
    div[role="radiogroup"] label[data-checked="true"] {
        background-color: #3498DB !important;
        color: white !important;
        border-color: #2980B9 !important;
        box-shadow: 0 2px 4px rgba(52, 152, 219, 0.2);
    }
    
    hr {
        margin-top: 10px;
        margin-bottom: 10px;
    }

    /* ========================================================= */
    /* 頁籤 Tab 客製化樣式：莫蘭迪深色調與放大字體    */
    /* ========================================================= */
    div[data-testid="stTabs"] button[data-baseweb="tab"] {
        background-color: #34495E !important; 
        color: #BDC3C7 !important;            
        border-radius: 8px 8px 0 0;
        padding: 12px 24px;
        border: none;
        margin-right: 5px;
    }
    div[data-testid="stTabs"] button[data-baseweb="tab"] > div {
        font-size: 1.15rem !important; 
        font-weight: 700 !important;
    }
    div[data-testid="stTabs"] button[data-baseweb="tab"][aria-selected="true"] {
        background-color: #4A6572 !important; 
        color: #FFFFFF !important;            
        border-bottom: 4px solid #E59866 !important; 
    }
    </style>
""", unsafe_allow_html=True)

# =========================================================
# 2. Google Sheets 連線設定
# =========================================================
SHEET_ID = "17BTvriIxK1ibFaPlmDUaTYwGrbmq556XNbzythA3rVo" 

MORANDI_BG = {
    "蘭潭校區": "#D4E6F1", 
    "民雄校區": "#FAD7A0", 
    "新民校區": "#D7BDE2", 
    "林森校區": "#A3E4D7", 
    "其他": "#EAEDED"
}

@st.cache_resource
def init_google_sheet():
    try:
        creds_dict = dict(st.secrets["gcp_oauth"])
        creds = Credentials.from_authorized_user_info(creds_dict, scopes=["https://www.googleapis.com/auth/spreadsheets"])
        gc = gspread.authorize(creds)
        return gc
    except Exception as e:
        st.error(f"連線失敗: {e}")
        return None

@st.cache_data(ttl=600)
def load_data(_gc):
    try:
        sh = _gc.open_by_key(SHEET_ID)
        
        ws_meters = sh.worksheet("電號對照表")
        meters_data = ws_meters.get_all_records()
        meters_df = pd.DataFrame(meters_data)
        if not meters_df.empty:
            meters_df.columns = meters_df.columns.str.strip()
            col_map = {c: c for c in meters_df.columns}
            for c in meters_df.columns:
                # 🔥 修正為 計費模式
                if "方式" in c or "模式" in c or "週期" in c or "計費" in c:
                    col_map[c] = "計費模式"
            meters_df.rename(columns=col_map, inplace=True)

        try:
            ws_records = sh.worksheet("用電填報紀錄")
            records_data = ws_records.get_all_records()
            records_df = pd.DataFrame(records_data)
            if not records_df.empty:
                records_df.columns = records_df.columns.str.strip()
        except:
            records_df = pd.DataFrame()

        return meters_df, records_df
    except Exception as e:
        st.error(f"資料讀取失敗: {e}")
        return pd.DataFrame(), pd.DataFrame()

def save_new_data(new_records):
    try:
        gc = init_google_sheet()
        sh = gc.open_by_key(SHEET_ID)
        try:
            ws = sh.worksheet("用電填報紀錄")
        except:
            ws = sh.add_worksheet(title="用電填報紀錄", rows=1000, cols=12)
            ws.append_row(["填報時間", "使用期間(起)", "使用期間(訖)", "計費模式", "校區", "電號", "用電地址", "用電量(度數)", "電費金額(元)", "帳單年度", "帳單月份", "備註"])
        
        rows = []
        t_now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for r in new_records:
            rows.append([
                t_now,
                str(r['使用期間(起)']),
                str(r['使用期間(訖)']),
                r['計費模式'], # 🔥 修正
                r['校區'],
                str(r['電號']),
                r['用電地址'],
                r['用電量(度數)'],
                r['電費金額(元)'],
                r['帳單年度'],
                r['帳單月份'],
                r.get('備註', '')
            ])
        ws.append_rows(rows)
        load_data.clear()
        return True
    except Exception as e:
        st.error(f"寫入失敗: {e}")
        return False

# =========================================================
# 局部重跑 Fragment：獨立填報區塊
# =========================================================
@st.fragment
def render_form_fragment(unreported, mode, bill_year, bill_month, default_start, default_end, monthly_start, monthly_end):
    if unreported.empty:
        st.success(f"🎉 {bill_year} 年 {bill_month} 月的「{mode}」電號皆已完成填報！")
        return

    st.markdown(f"##### 2️⃣ 填寫電號數據")
    
    campus_groups = unreported.groupby("校區")
    campus_order = ["蘭潭校區", "民雄校區", "新民校區", "林森校區"]
    sorted_groups = sorted(campus_groups, key=lambda x: campus_order.index(x[0]) if x[0] in campus_order else 99)

    input_data = {}
    current_reset_key = st.session_state['form_reset_key']

    for campus, group_data in sorted_groups:
        bg_color = MORANDI_BG.get(campus, "#EAEDED")
        border_color = bg_color.replace('D', '9').replace('E', 'A')
        
        st.markdown(f"""
        <div class="campus-header" style="background-color: {bg_color}; border-left: 8px solid {border_color};">
            🏫 {campus}
        </div>
        """, unsafe_allow_html=True)
        
        for idx, row in group_data.iterrows():
            m_id = str(row['電號'])
            m_name = row['用電地址']
            
            st.markdown(f"""
            <div class="meter-card" style="border-left-color: {border_color};">
                <div class="meter-header">
                    <span class="meter-name">{m_name}</span>
                    <span class="meter-id">電號: {m_id}</span>
                </div>
            """, unsafe_allow_html=True)
            
            k_u = f"u_{m_id}_{idx}_{current_reset_key}"
            k_c = f"c_{m_id}_{idx}_{current_reset_key}"
            k_d1 = f"d1_{m_id}_{idx}_{current_reset_key}"
            k_d2 = f"d2_{m_id}_{idx}_{current_reset_key}"
            
            if mode == "抄表":
                c_u, c_c, c_d1, c_d2 = st.columns([1, 1, 1, 1])
                val_u = c_u.number_input("用電度數", min_value=0, step=1, key=k_u)
                val_c = c_c.number_input("電費金額(元)", min_value=0, step=1, key=k_c)
                current_start = c_d1.date_input("計費起始日", default_start, key=k_d1)
                current_end = c_d2.date_input("計費結束日", default_end, key=k_d2)
            else:
                c_u, c_c = st.columns([1, 1])
                val_u = c_u.number_input("用電度數", min_value=0, step=1, key=k_u)
                val_c = c_c.number_input("電費金額(元)", min_value=0, step=1, key=k_c)
                current_start = monthly_start
                current_end = monthly_end

            st.markdown("</div>", unsafe_allow_html=True)

            if val_u > 0 or val_c > 0:
                input_data[m_id] = {
                    "使用期間(起)": current_start,
                    "使用期間(訖)": current_end,
                    "計費模式": mode, # 🔥 修正
                    "校區": campus,
                    "電號": m_id,
                    "用電地址": m_name,
                    "用電量(度數)": val_u,
                    "電費金額(元)": val_c,
                    "帳單年度": bill_year,
                    "帳單月份": bill_month
                }

    st.markdown("<br>", unsafe_allow_html=True)
    
    if st.button("📤 確認資料無誤，送出申報"):
        if input_data:
            with st.spinner("資料上傳中..."):
                final_records = list(input_data.values())
                if save_new_data(final_records):
                    st.success(f"🎉 成功申報 {len(final_records)} 筆資料！")
                    st.balloons()
                    st.session_state['form_reset_key'] += 1
                    time.sleep(2)
                    st.rerun()
                else:
                    st.error("❌ 寫入失敗，請檢查網路。")
        else:
            st.warning("⚠️ 您尚未填寫任何數值。")

# =========================================================
# Tab 1 封裝：資料填報
# =========================================================
def render_tab1_input():
    st.markdown("##### 1️⃣ 申報模式與週期")
    c_mode, c_year, c_month = st.columns([2, 1, 1])
    
    with c_mode:
        mode = st.radio("計費模式", ["整月", "抄表"], horizontal=True, label_visibility="collapsed", key="input_mode") # 🔥 修正

    current_year = datetime.now().year
    with c_year:
        bill_year = st.number_input("帳單年度 (西元)", min_value=2020, max_value=2030, value=current_year, key="input_year")
    with c_month:
        bill_month = st.selectbox("帳單月份", range(1, 13), index=datetime.now().month - 1, key="input_month")
    
    bill_date = date(bill_year, bill_month, 1)
    default_start = bill_date - relativedelta(months=2)
    default_end = bill_date - timedelta(days=1)
    monthly_start = (bill_date - relativedelta(months=1)).replace(day=1)
    monthly_end = (bill_date - timedelta(days=1))

    st.info(f"您目前正在填報 **{bill_year} 年 {bill_month} 月** 的帳單資訊")
    st.markdown("---")

    gc = init_google_sheet()
    if not gc: return
    meters_df, records_df = load_data(gc)
    
    if meters_df.empty:
        st.warning("⚠️ 找不到電號資料。")
        return

    # 🔥 修正為 計費模式
    if "計費模式" not in meters_df.columns: meters_df["計費模式"] = "整月"
    if "校區" not in meters_df.columns: meters_df["校區"] = "其他"

    is_manual_reading = meters_df['計費模式'].astype(str).str.contains("抄表", na=False) # 🔥 修正

    if mode == "抄表":
        target_meters = meters_df[is_manual_reading].copy()
    else:
        target_meters = meters_df[~is_manual_reading].copy()

    meters_df['電號'] = meters_df['電號'].astype(str)
    if not records_df.empty and '帳單年度' in records_df.columns:
        records_df['電號'] = records_df['電號'].astype(str)
        filled_mask = (records_df['帳單年度'] == bill_year) & (records_df['帳單月份'] == bill_month)
        filled_ids = set(records_df[filled_mask]['電號'].unique())
        unreported = target_meters[~target_meters['電號'].isin(filled_ids)].copy()
    else:
        unreported = target_meters.copy()

    render_form_fragment(unreported, mode, bill_year, bill_month, default_start, default_end, monthly_start, monthly_end)

# =========================================================
# Tab 2 封裝：資料異動 (視覺化編輯)
# =========================================================
@st.fragment
def render_tab2_edit():
    st.info("💡 **操作說明**：系統會自動篩選出「有申報紀錄的期間與計費模式」。請選擇欲管理的條件，即可在下方卡片進行精準修改與覆蓋。")
    
    gc = init_google_sheet()
    if not gc: return
    _, records_df = load_data(gc)

    if records_df.empty:
        st.warning("目前資料庫為空，無資料可供異動。")
        return

    # =========================================================
    # 防呆機制：確保舊版資料表也有必要的欄位
    # =========================================================
    if "計費模式" not in records_df.columns: # 🔥 修正
        records_df["計費模式"] = "整月"  
    if "校區" not in records_df.columns:
        records_df["校區"] = "其他"
    if "帳單年度" not in records_df.columns:
        records_df["帳單年度"] = datetime.now().year
    if "帳單月份" not in records_df.columns:
        records_df["帳單月份"] = datetime.now().month

    valid_years = records_df['帳單年度'].dropna().unique()
    years = sorted(list(set(valid_years.astype(int))), reverse=True) if len(valid_years) > 0 else [datetime.now().year]

    col_y, col_m, col_mode = st.columns([1, 1, 1])
    
    with col_y:
        selected_year = st.selectbox("📅 選擇帳單年度", years, key="edit_year")

    mask_year = records_df['帳單年度'] == selected_year
    valid_months = records_df.loc[mask_year, '帳單月份'].dropna().unique()
    month_options = sorted(list(set(valid_months.astype(int)))) if len(valid_months) > 0 else [datetime.now().month]

    with col_m:
        selected_month = st.selectbox("🗓️ 選擇帳單月份", month_options, key="edit_month")

    # 🔥 修正為 計費模式
    mask_month = mask_year & (records_df['帳單月份'] == selected_month)
    valid_modes = records_df.loc[mask_month, '計費模式'].dropna().unique()
    mode_options = list(set(valid_modes)) if len(valid_modes) > 0 else ["整月", "抄表"]

    with col_mode:
        selected_mode = st.selectbox("📝 選擇計費模式", mode_options, key="edit_mode") # 🔥 修正

    st.markdown("---")

    # 🔥 修正為 計費模式
    mask_target = mask_month & (records_df['計費模式'] == selected_mode)
    df_target = records_df[mask_target].copy()

    if df_target.empty:
        st.info(f"📌 {selected_year} 年 {selected_month} 月的「{selected_mode}」尚無填報紀錄。")
        return
    else:
        st.success(f"✅ 已載入 {selected_year} 年 {selected_month} 月的「{selected_mode}」紀錄，請於下方進行修改。")

    campus_groups = df_target.groupby("校區")
    campus_order = ["蘭潭校區", "民雄校區", "新民校區", "林森校區"]
    sorted_groups = sorted(campus_groups, key=lambda x: campus_order.index(x[0]) if x[0] in campus_order else 99)

    with st.form("visual_edit_power_form", clear_on_submit=False):
        updated_records = []

        for campus, group_data in sorted_groups:
            bg_color = MORANDI_BG.get(campus, "#EAEDED")
            border_color = bg_color.replace('D', '9').replace('E', 'A')
            
            st.markdown(f"""
            <div class="campus-header" style="background-color: {bg_color}; border-left: 8px solid {border_color};">
                🏫 {campus}
            </div>
            """, unsafe_allow_html=True)
            
            for idx, row in group_data.iterrows():
                m_id = str(row['電號'])
                m_name = row['用電地址']
                
                st.markdown(f"""
                <div class="meter-card" style="border-left-color: {border_color};">
                    <div class="meter-header">
                        <span class="meter-name">{m_name}</span>
                        <span class="meter-id">電號: {m_id}</span>
                    </div>
                """, unsafe_allow_html=True)
                
                def_u = int(row.get('用電量(度數)', 0))
                def_c = int(row.get('電費金額(元)', 0))
                
                k_u = f"edit_u_{m_id}_{idx}"
                k_c = f"edit_c_{m_id}_{idx}"
                
                if selected_mode == "抄表":
                    try: def_start = pd.to_datetime(str(row.get('使用期間(起)', datetime.now().date()))).date()
                    except: def_start = datetime.now().date()
                    
                    try: def_end = pd.to_datetime(str(row.get('使用期間(訖)', datetime.now().date()))).date()
                    except: def_end = datetime.now().date()

                    c_u, c_c, c_d1, c_d2 = st.columns([1, 1, 1, 1])
                    val_u = c_u.number_input("用電度數", min_value=0, step=1, value=def_u, key=k_u)
                    val_c = c_c.number_input("電費金額(元)", min_value=0, step=1, value=def_c, key=k_c)
                    val_start = c_d1.date_input("計費起始日", value=def_start, key=f"edit_d1_{m_id}_{idx}")
                    val_end = c_d2.date_input("計費結束日", value=def_end, key=f"edit_d2_{m_id}_{idx}")
                else:
                    c_u, c_c = st.columns([1, 1])
                    val_u = c_u.number_input("用電度數", min_value=0, step=1, value=def_u, key=k_u)
                    val_c = c_c.number_input("電費金額(元)", min_value=0, step=1, value=def_c, key=k_c)
                    val_start = row.get('使用期間(起)', '')
                    val_end = row.get('使用期間(訖)', '')

                st.markdown("</div>", unsafe_allow_html=True)

                if val_u > 0 or val_c > 0:
                    updated_records.append({
                        "填報時間": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "使用期間(起)": str(val_start),
                        "使用期間(訖)": str(val_end),
                        "計費模式": selected_mode, # 🔥 修正
                        "校區": campus,
                        "電號": m_id,
                        "用電地址": m_name,
                        "用電量(度數)": val_u,
                        "電費金額(元)": val_c,
                        "帳單年度": selected_year,
                        "帳單月份": selected_month,
                        "備註": row.get('備註', '')
                    })

        st.markdown("<br>", unsafe_allow_html=True)
        submitted = st.form_submit_button("💾 精準更新並儲存", type="primary", use_container_width=True)

        if submitted:
            try:
                sh = gc.open_by_key(SHEET_ID)
                ws = sh.worksheet("用電填報紀錄")
                
                df_keep = records_df[~mask_target].copy()
                
                if updated_records:
                    df_new = pd.DataFrame(updated_records)
                    df_final = pd.concat([df_keep, df_new], ignore_index=True)
                else:
                    df_final = df_keep 
                    
                # 🔥 確保標題列也是 計費模式
                cols = ["填報時間", "使用期間(起)", "使用期間(訖)", "計費模式", "校區", "電號", "用電地址", "用電量(度數)", "電費金額(元)", "帳單年度", "帳單月份", "備註"]
                for c in cols:
                    if c not in df_final.columns:
                        df_final[c] = ""
                df_final = df_final[cols]
                
                ws.clear()
                ws.update([cols] + df_final.values.tolist())
                
                st.toast(f"✅ {selected_year} 年 {selected_month} 月 ({selected_mode}) 資料已精準更新完成！", icon="🎉")
                load_data.clear()
                time.sleep(1.5)
                st.rerun()
                
            except Exception as e:
                st.error(f"❌ 更新失敗: {str(e)}")

# =========================================================
# 3. 主程式邏輯佈署
# =========================================================
def main():
    st.subheader("⚡ 全校電力填報系統")

    tab1, tab2 = st.tabs(["📤 資料填報", "✏️ 資料異動"])

    with tab1:
        render_tab1_input()

    with tab2:
        render_tab2_edit()

if __name__ == "__main__":
    main()