import streamlit as st
import pandas as pd
import gspread
import datetime
import time
import io
import os
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# ================= 系統與全域參數設定 =================
st.set_page_config(page_title="實驗室氣體鋼瓶資料回報", page_icon="💨", layout="wide")

SHEET_ID = '1Hw4rXo4ww7O9YXTwoUJeWioO5ZzM_bivRcLLpOl26DY'
FOLDER_ID = '1cpj6l-zTgqLwphMFzLvvr53-Alf60idN'

# 指定氣體種類
GAS_TYPES = ["請選擇", "二氧化碳", "甲烷", "乙炔", "一氧化二氮(笑氣)"]

# ================= 莫蘭迪風格 CSS & UI =================
def apply_morandi_theme():
    st.markdown("""
    <style>
        .title-box { background-color: #5C6B73; color: #FFFFFF; padding: 20px; border-radius: 10px; text-align: center; margin-bottom: 20px; }
        .section-header { background-color: #9DB4AB; color: #FFFFFF; padding: 10px 15px; border-radius: 8px; font-size: 20px; font-weight: bold; margin-top: 25px; margin-bottom: 15px; }
        .card { background-color: #FDFBF7; border: 1px solid #DCE4E3; padding: 15px 20px; border-radius: 8px; margin-bottom: 15px; }
        .badge-inventory { background-color: #E29578; color: white; padding: 5px 10px; border-radius: 4px; font-weight: bold; }
        .badge-purchase { background-color: #83C5BE; color: white; padding: 5px 10px; border-radius: 4px; font-weight: bold; }
        button[kind="primary"] { background-color: #27AE60 !important; color: #FFFFFF !important; font-weight: bold !important; font-size: 16px !important; padding: 10px 20px !important;}
        [data-testid="stSidebar"], [data-testid="collapsedControl"], [data-testid="stHeader"] { display: none !important; }
        .block-container { padding-top: 2rem !important; }
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

@st.cache_data(ttl=120, show_spinner=False)
def fetch_tracker_records(): 
    try: return get_gc().open_by_key(SHEET_ID).worksheet('發送紀錄與金鑰').get_all_records()
    except: return []

@st.cache_data(ttl=120, show_spinner=False)
def fetch_main_inventory(): 
    try: return get_gc().open_by_key(SHEET_ID).worksheet('氣體鋼瓶資料庫').get_all_records()
    except: return []

@with_retry(max_retries=3)
def append_dicts(worksheet, list_of_dicts):
    headers = worksheet.row_values(1)
    if headers and list_of_dicts: 
        worksheet.append_rows([[str(d.get(str(h).strip(), "")) for h in headers] for d in list_of_dicts])

@with_retry(max_retries=3)
def safe_delete_row(worksheet, idx): 
    worksheet.delete_rows(idx)

@with_retry(max_retries=3)
def upload_to_drive(uploaded_file, file_name):
    creds = get_credentials()
    drive_service = build('drive', 'v3', credentials=creds)
    file_metadata = {'name': file_name, 'parents': [FOLDER_ID]}
    media = MediaIoBaseUpload(io.BytesIO(uploaded_file.getvalue()), mimetype=uploaded_file.type, resumable=True)
    file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    return f"https://drive.google.com/file/d/{file.get('id')}/view?usp=sharing"

# ================= 主程式 =================
def main():
    apply_morandi_theme()
    url_token = st.query_params.get("token", None)
    
    if not url_token: 
        st.error("🚫 請透過專屬通知信內的連結進入系統。")
        st.stop()

    if st.session_state.get("token") != url_token:
        st.session_state.clear()
        st.session_state.token = url_token
        
        records = fetch_tracker_records()
        record, row_idx = next(((r, i+2) for i, r in enumerate(records) if str(r.get("專屬金鑰(Token)")).strip() == url_token), (None, None))
        
        if not record: 
            st.error("🚫 無效的專屬連結，或連結已失效。")
            st.stop()
            
        st.session_state.status = str(record.get("回覆狀態", ""))
        st.session_state.dept = str(record.get("系所", ""))
        st.session_state.manager = str(record.get("實驗室老師", ""))
        st.session_state.room = str(record.get("氣體鋼瓶所在位置實驗室門牌", ""))
        st.session_state.campus = str(record.get("校區", ""))
        st.session_state.email = str(record.get("電子郵件", ""))
        st.session_state.row_idx = row_idx
        
        # 抓取該實驗室的現有鋼瓶庫存
        df_inv = pd.DataFrame(fetch_main_inventory())
        if not df_inv.empty:
            mask = (df_inv['系所'].astype(str) == st.session_state.dept) & \
                   (df_inv['實驗室老師'].astype(str) == st.session_state.manager) & \
                   (df_inv['氣體鋼瓶所在位置實驗室門牌'].astype(str) == st.session_state.room)
            st.session_state.inventory = df_inv[mask].to_dict('records')
        else:
            st.session_state.inventory = []

    is_readonly = (st.session_state.status == "已回報")
    
    st.markdown(f"""
    <div class="title-box">
        <h2 style="margin: 0;">國立嘉義大學 氣體鋼瓶年度盤查與回報系統</h2>
        <p style="margin: 10px 0 0 0; font-size: 20px;">🏢 {st.session_state.campus} | {st.session_state.dept} - {st.session_state.manager} 老師 ({st.session_state.room})</p>
    </div>
    """, unsafe_allow_html=True)
    
    if is_readonly: 
        st.success("✅ 貴實驗室已完成本期回報，目前為唯讀檢視模式。")
    
    # ================= 任務一：現有庫存盤點 =================
    st.markdown('<div class="section-header">🏭 任務一：確認目前實驗室鋼瓶庫存 (現況總量)</div>', unsafe_allow_html=True)
    st.info("💡 請確認下方列出的鋼瓶現況。若數量有異動請直接修改；若某種氣體已清空，請將數量改為 0 (系統將自動刪除該筆紀錄)。")
    
    collected_inv = []
    
    # 呈現既有庫存
    for idx, inv in enumerate(st.session_state.inventory):
        st.markdown(f"<div class='card'><span class='badge-inventory'>現有庫存 {idx+1}</span><br><br>", unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        g_type = str(inv.get("鋼瓶氣體種類", "請選擇")).strip()
        g_idx = GAS_TYPES.index(g_type) if g_type in GAS_TYPES else 0
        
        with c1: e_gas = st.selectbox("氣體種類", GAS_TYPES, index=g_idx, disabled=is_readonly, key=f"ig_{idx}")
        with c2: e_qty = st.number_input("鋼瓶數量 (支)", min_value=0, value=int(inv.get("鋼瓶數量", 0) or 0), disabled=is_readonly, key=f"iq_{idx}")
        st.markdown("</div>", unsafe_allow_html=True)
        
        if not is_readonly and e_gas != "請選擇" and e_qty > 0:
            collected_inv.append({"鋼瓶氣體種類": e_gas, "鋼瓶數量": e_qty})

    # 提供新增庫存選項
    if not is_readonly:
        st.markdown("##### ➕ 新增庫存氣體種類")
        add_inv = st.number_input("欲新增的氣體種類數量", min_value=0, max_value=5, value=0)
        for j in range(int(add_inv)):
            st.markdown(f"<div class='card'><span class='badge-inventory'>新增庫存</span><br><br>", unsafe_allow_html=True)
            c1, c2 = st.columns(2)
            with c1: n_gas = st.selectbox("氣體種類", GAS_TYPES, key=f"nig_{j}")
            with c2: n_qty = st.number_input("鋼瓶數量 (支)", min_value=1, value=1, key=f"niq_{j}")
            st.markdown("</div>", unsafe_allow_html=True)
            
            if n_gas != "請選擇":
                collected_inv.append({"鋼瓶氣體種類": n_gas, "鋼瓶數量": n_qty})

    # ================= 任務二：年度購買申報 =================
    st.markdown('<div class="section-header">🧾 任務二：年度購買紀錄申報 (碳盤查活動數據)</div>', unsafe_allow_html=True)
    st.write("請依據本年度購買的單據，逐筆登錄氣體購買量並上傳佐證照片或 PDF。 **若本年度無購買，此區塊留空數量填 0 即可**。")
    
    collected_purchases = []
    
    if not is_readonly:
        add_pur = st.number_input("➕ 新增本年度購買單據筆數", min_value=0, max_value=10, value=0)
        for k in range(int(add_pur)):
            st.markdown(f"<div class='card'><span class='badge-purchase'>購買紀錄 {k+1}</span><br><br>", unsafe_allow_html=True)
            pc1, pc2, pc3 = st.columns(3)
            with pc1: p_gas = st.selectbox("購買氣體種類", GAS_TYPES, key=f"pg_{k}")
            with pc2: p_date = st.date_input("購買日期", key=f"pd_{k}")
            with pc3: p_amt = st.number_input("本張單據購買量 (公斤)", min_value=0.0, step=0.1, key=f"pa_{k}")
            
            p_file = st.file_uploader(f"📎 上傳此筆購買佐證單據 (限 JPG, PNG, PDF)", type=['png', 'jpg', 'jpeg', 'pdf'], key=f"pf_{k}")
            st.markdown("</div>", unsafe_allow_html=True)
            
            if p_gas != "請選擇" and p_amt > 0 and p_file:
                collected_purchases.append({"鋼瓶氣體種類": p_gas, "購買日期": p_date.strftime("%Y-%m-%d"), "年度氣體鋼瓶購買量(公斤)": p_amt, "file": p_file})
            elif p_gas != "請選擇" and p_amt > 0 and not p_file:
                st.warning(f"⚠️ 購買紀錄 {k+1} 尚未上傳佐證單據，請補齊後再送出！")

        st.markdown("<hr style='border: 1px solid #EAECEE;'>", unsafe_allow_html=True)
        
        # 取得老師分機 (可選填)
        st.markdown("##### 📞 聯絡資訊確認")
        e_ext = st.text_input("實驗室分機 (選填)", value=st.session_state.inventory[0].get("分機", "") if st.session_state.inventory else "")

        if st.button("🚀 確認盤點結果並送出表單", type="primary", use_container_width=True):
            # 防呆：檢查是否有未完成單據的購買紀錄
            if any(p_gas != "請選擇" and p_amt > 0 and not p_file for p_gas, p_amt, p_file in [(st.session_state.get(f"pg_{x}"), st.session_state.get(f"pa_{x}"), st.session_state.get(f"pf_{x}")) for x in range(int(add_pur))]):
                st.error("❌ 您有已填寫數量的購買紀錄尚未上傳單據，請檢查！")
                st.stop()

            now_str = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
            now_year = str(datetime.datetime.now().year - 1911) + "年" # 民國年
            
            with st.spinner("正在上傳單據並將資料寫入系統資料庫中，請勿關閉網頁..."):
                gc = get_gc()
                sh_main = gc.open_by_key(SHEET_ID)
                ws_inv = sh_main.worksheet('氣體鋼瓶資料庫')
                ws_pur = sh_main.worksheet('氣體鋼瓶年度使用紀錄')
                ws_trk = sh_main.worksheet('發送紀錄與金鑰')
                
                # --- 1. 寫入庫存：精準覆寫 ---
                # 刪除舊有該實驗室的紀錄 (從下往上刪除避免 index 跑掉)
                all_inv_records = ws_inv.get_all_records()
                rows_to_del = []
                for i, r in enumerate(all_inv_records):
                    if str(r.get('系所', '')) == st.session_state.dept and \
                       str(r.get('實驗室老師', '')) == st.session_state.manager and \
                       str(r.get('氣體鋼瓶所在位置實驗室門牌', '')) == st.session_state.room:
                        rows_to_del.append(i + 2)
                        
                for r_idx in sorted(rows_to_del, reverse=True):
                    safe_delete_row(ws_inv, r_idx)
                    time.sleep(0.3) # 避免 API 速率限制
                
                # 新增本次盤點結果
                inv_to_append = []
                for item in collected_inv:
                    inv_to_append.append({
                        "系所": st.session_state.dept, 
                        "實驗室老師": st.session_state.manager, 
                        "校區": st.session_state.campus,
                        "氣體鋼瓶所在位置實驗室門牌": st.session_state.room, 
                        "電子郵件": st.session_state.email,
                        "分機": e_ext,
                        "鋼瓶氣體種類": item["鋼瓶氣體種類"], 
                        "鋼瓶數量": item["鋼瓶數量"],
                        "建檔年度": now_year
                    })
                if inv_to_append:
                    append_dicts(ws_inv, inv_to_append)
                
                # --- 2. 寫入年度使用紀錄 (含上傳 Drive) ---
                pur_to_append = []
                for p in collected_purchases:
                    ext = os.path.splitext(p["file"].name)[1]
                    # 命名規則: 系所-實驗室老師-鋼瓶氣體種類-購買日期-年度氣體鋼瓶購買量(公斤)
                    f_name = f"{st.session_state.dept}-{st.session_state.manager}-{p['鋼瓶氣體種類']}-{p['購買日期']}-{p['年度氣體鋼瓶購買量(公斤)']}{ext}"
                    drive_link = upload_to_drive(p["file"], f_name)
                    
                    pur_to_append.append({
                        "填報時間": now_str,
                        "系所": st.session_state.dept, 
                        "實驗室老師": st.session_state.manager, 
                        "校區": st.session_state.campus,
                        "氣體鋼瓶所在位置實驗室門牌": st.session_state.room, 
                        "鋼瓶氣體種類": p["鋼瓶氣體種類"], 
                        "購買日期": p["購買日期"], 
                        "年度氣體鋼瓶購買量(公斤)": p["年度氣體鋼瓶購買量(公斤)"],
                        "購買單據連結": drive_link
                    })
                if pur_to_append:
                    append_dicts(ws_pur, pur_to_append)
                
                # --- 3. 更新追蹤表狀態 ---
                ws_trk.update_cell(st.session_state.row_idx, 11, "已回報") # 假設 '回覆狀態' 是第 11 欄
                ws_trk.update_cell(st.session_state.row_idx, 12, now_str)  # 假設 '回覆時間' 是第 12 欄
                
            st.session_state.status = "已回報"
            st.success("✅ 資料盤點與單據上傳已全部完成，感謝您的配合！")
            st.balloons()
            time.sleep(2)
            st.rerun()

if __name__ == "__main__":
    main()