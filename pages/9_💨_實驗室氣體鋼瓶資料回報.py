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
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT

# ================= 系統與全域參數設定 =================
st.set_page_config(page_title="實驗室氣體鋼瓶資料回報", page_icon="💨", layout="wide")

SHEET_ID = '1Hw4rXo4ww7O9YXTwoUJeWioO5ZzM_bivRcLLpOl26DY'
FOLDER_ID = '1cpj6l-zTgqLwphMFzLvvr53-Alf60idN'

CAMPUS_OPTS = ["蘭潭校區", "民雄校區", "新民校區", "林森校區"]
GAS_TYPES = ["請選擇", "二氧化碳", "甲烷", "乙炔", "一氧化二氮(笑氣)"]

# ================= 莫蘭迪風格 CSS & UI =================
def apply_morandi_theme():
    st.markdown("""
    <style>
        /* 字體全面放大一號 */
        label, p, .st-emotion-cache-1y4p8pa, .st-emotion-cache-1104ytp p { font-size: 18px !important; }
        
        .title-box { background-color: #5C6B73; color: #FFFFFF; padding: 20px; border-radius: 10px; text-align: center; margin-bottom: 20px; }
        .section-header { background-color: #9DB4AB; color: #FFFFFF; padding: 10px 15px; border-radius: 8px; font-size: 22px; font-weight: bold; margin-top: 25px; margin-bottom: 15px; }
        
        /* 淺藍色 Radio 按鈕 (有無購買) */
        div.stRadio > div[role="radiogroup"] > label {
            background-color: #E3F2FD !important;
            padding: 10px 20px !important;
            border-radius: 8px !important;
            border: 2px solid #BBDEFB !important;
            cursor: pointer;
        }
        div.stRadio > div[role="radiogroup"] > label p { color: #0D47A1 !important; font-weight: bold !important; margin: 0;}
        div.stRadio > div[role="radiogroup"] > label[data-checked="true"] { background-color: #90CAF9 !important; border-color: #1976D2 !important; }
        
        /* 上傳檔案區塊：全淺藍色無灰底 */
        [data-testid="stFileUploader"] { background-color: #E3F2FD !important; border: 1px dashed #64B5F6 !important; border-radius: 8px; padding: 15px; }
        [data-testid="stFileUploaderDropzone"] { background-color: transparent !important; }
        [data-testid="stFileUploaderDropzone"] > div { background-color: transparent !important; }
        
        /* 一般按鈕 */
        div.stButton > button { background-color: #5C6B73 !important; color: #FFFFFF !important; font-weight: bold !important; font-size: 18px !important; padding: 12px 20px !important;}
        div.stButton > button:hover { background-color: #4A565C !important; }
        
        /* 特製橘色下載按鈕 (套用於正式申報相關) */
        div.stButton > button.st-emotion-cache-1v0mbdj, div[data-testid="stDownloadButton"] button { 
            background-color: #E67E22 !important; border: 2px solid #D35400 !important; color: white !important; 
        }
        div.stButton > button.st-emotion-cache-1v0mbdj:hover, div[data-testid="stDownloadButton"] button:hover { 
            background-color: #D35400 !important; 
        }
        
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
                    time.sleep(base_delay * (2 ** retries)); retries += 1
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
def safe_delete_row(worksheet, idx): worksheet.delete_rows(idx)

@with_retry(max_retries=3)
def upload_to_drive(uploaded_file, file_name):
    creds = get_credentials(); drive_service = build('drive', 'v3', credentials=creds)
    file_metadata = {'name': file_name, 'parents': [FOLDER_ID]}
    media = MediaIoBaseUpload(io.BytesIO(uploaded_file.getvalue()), mimetype=uploaded_file.type, resumable=True)
    file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    return f"https://drive.google.com/file/d/{file.get('id')}/view?usp=sharing"

# ================= 產生 Word 申報表 =================
def create_report_docx(base_data, pur_data, status):
    doc = Document()
    
    # 🟢 版面使用 "中等邊界"
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    title = doc.add_heading('溫室氣體盤查 - 實驗室氣體鋼瓶申報表', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 🟢 增加盤查年度資訊
    doc.add_paragraph(f"盤查年度：2026年")
    doc.add_paragraph(f"填報日期：{datetime.datetime.now().strftime('%Y/%m/%d')}")
    
    # 實驗室基本資訊表
    doc.add_heading('【實驗室基本資訊】', level=2)
    t_base = doc.add_table(rows=3, cols=2)
    t_base.style = 'Table Grid'
    t_base.cell(0, 0).text = "系所 / 老師"; t_base.cell(0, 1).text = f"{base_data['dept']} / {base_data['mgr']}"
    t_base.cell(1, 0).text = "校區 / 門牌"; t_base.cell(1, 1).text = f"{base_data['campus']} / {base_data['room']}"
    t_base.cell(2, 0).text = "電子郵件 / 分機"; t_base.cell(2, 1).text = f"{base_data['mail']} / {base_data['ext']}"
    
    for row in t_base.rows:
        for cell in row.cells: cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        
    doc.add_heading('【年度購買紀錄申報】', level=2)
    if status == "無購買" or not pur_data:
        doc.add_paragraph("☑️ 盤查期間無購買氣體鋼瓶")
    else:
        for idx, p in enumerate(pur_data):
            p_table = doc.add_table(rows=3, cols=2)
            p_table.style = 'Table Grid'
            p_table.cell(0, 0).text = "鋼瓶氣體種類"
            p_table.cell(0, 1).text = str(p['鋼瓶氣體種類'])
            # 🟢 順序調整：先購買日期，再購買量
            p_table.cell(1, 0).text = "購買日期"
            p_table.cell(1, 1).text = str(p['購買日期'])
            p_table.cell(2, 0).text = "購買量(公斤)"
            p_table.cell(2, 1).text = f"{p['年度氣體鋼瓶購買量(公斤)']} kg"
            
            for row in p_table.rows:
                for cell in row.cells: cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            
            # 🟢 於購買量下方加入單據圖片
            doc.add_paragraph("")
            file_obj = p['file']
            if file_obj and file_obj.type in ['image/jpeg', 'image/png']:
                try:
                    file_obj.seek(0)
                    doc.add_picture(io.BytesIO(file_obj.read()), width=Inches(5))
                except Exception: doc.add_paragraph("⚠️ 無法載入單據圖片")
            elif file_obj: doc.add_paragraph(f"📎 單據格式為 {file_obj.name}，請至雲端資料夾檢視。")
            doc.add_paragraph("\n")

    doc.add_paragraph("\n申報聲明：本人確認以上資料屬實，並依規定完成申報。")
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

# ================= 主程式 =================
def main():
    apply_morandi_theme()
    url_token = st.query_params.get("token", None)
    
    if not url_token: st.error("🚫 請透過專屬通知信內的連結進入系統。"); st.stop()

    if st.session_state.get("token") != url_token:
        st.session_state.clear(); st.session_state.token = url_token
        records = fetch_tracker_records()
        record, row_idx = next(((r, i+2) for i, r in enumerate(records) if str(r.get("專屬金鑰(Token)")).strip() == url_token), (None, None))
        
        if not record: st.error("🚫 無效的專屬連結，或連結已失效。"); st.stop()
            
        st.session_state.status = str(record.get("回覆狀態", ""))
        st.session_state.dept = str(record.get("系所", ""))
        st.session_state.manager = str(record.get("實驗室老師", ""))
        st.session_state.room = str(record.get("氣體鋼瓶所在位置實驗室門牌", ""))
        st.session_state.campus = str(record.get("校區", ""))
        st.session_state.email = str(record.get("電子郵件", ""))
        st.session_state.row_idx = row_idx
        
        df_inv = pd.DataFrame(fetch_main_inventory())
        if not df_inv.empty:
            mask = (df_inv['系所'].astype(str) == st.session_state.dept) & (df_inv['實驗室老師'].astype(str) == st.session_state.manager) & (df_inv['氣體鋼瓶所在位置實驗室門牌'].astype(str) == st.session_state.room)
            st.session_state.inventory = df_inv[mask].to_dict('records')
        else: st.session_state.inventory = []
        
        st.session_state.step = "input"

    is_readonly = (st.session_state.status == "已回報")
    
    st.markdown(f"""
    <div class="title-box">
        <h2 style="margin: 0;">國立嘉義大學 氣體鋼瓶年度盤查與回報系統</h2>
    </div>
    """, unsafe_allow_html=True)
    
    if is_readonly: st.success("✅ 貴實驗室已完成本期回報，目前為唯讀檢視模式。"); st.stop()
    
    # 取出預設值
    inv_data = st.session_state.inventory[0] if st.session_state.inventory else {}
    def_room = inv_data.get("氣體鋼瓶所在位置實驗室門牌", st.session_state.room)
    def_ext = inv_data.get("分機", "")
    def_mail = inv_data.get("電子郵件", st.session_state.email)
    
    # ================= 第一階段：填寫資料 =================
    if st.session_state.step == "input":
        
        # 🟢 開放基本資訊編輯與確認
        st.markdown('<div class="section-header">👩‍🔬 實驗室基本資訊確認與修改</div>', unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        with c1: e_dept = st.text_input("系所", value=st.session_state.dept)
        with c2: e_mgr = st.text_input("實驗室老師", value=st.session_state.manager)
        with c3: e_campus = st.selectbox("校區", CAMPUS_OPTS, index=CAMPUS_OPTS.index(st.session_state.campus) if st.session_state.campus in CAMPUS_OPTS else 0)
        
        c4, c5, c6 = st.columns(3)
        with c4: e_room = st.text_input("實驗室門牌", value=def_room, help="若管理多間請一併填寫")
        with c5: e_mail = st.text_input("電子郵件", value=def_mail)
        with c6: e_ext = st.text_input("聯絡分機", value=def_ext)
        
        st.markdown('<div class="section-header">🏭 確認目前實驗室鋼瓶庫存 (現況總量)</div>', unsafe_allow_html=True)
        st.info("💡 請確認下方列出的鋼瓶現況。若數量有異動請直接修改；若某種氣體已清空，請將數量改為 0。")
        
        collected_inv = []
        for idx, inv in enumerate(st.session_state.inventory):
            st.markdown(f"**🔹 溫室氣體盤查範疇之氣體鋼瓶使用種類現況 {idx+1}**", unsafe_allow_html=True)
            cc1, cc2 = st.columns(2)
            g_type = str(inv.get("鋼瓶氣體種類", "請選擇")).strip()
            g_idx = GAS_TYPES.index(g_type) if g_type in GAS_TYPES else 0
            
            with cc1: e_gas = st.selectbox("氣體種類", GAS_TYPES, index=g_idx, key=f"ig_{idx}")
            with cc2: e_qty = st.number_input("鋼瓶數量 (支)", min_value=0, value=int(inv.get("鋼瓶數量", 0) or 0), key=f"iq_{idx}")
            st.markdown("<hr style='border: 1px dashed #BDC3C7;'>", unsafe_allow_html=True)
            if e_gas != "請選擇" and e_qty > 0: collected_inv.append({"鋼瓶氣體種類": e_gas, "鋼瓶數量": e_qty})

        st.markdown("##### ➕ 新增庫存氣體種類")
        add_inv = st.number_input("欲新增的現有氣體種類數量", min_value=0, max_value=5, value=0)
        for j in range(int(add_inv)):
            st.markdown(f"**🔹 新增現有庫存 {j+1}**", unsafe_allow_html=True)
            c1, c2 = st.columns(2)
            with c1: n_gas = st.selectbox("氣體種類", GAS_TYPES, key=f"nig_{j}")
            with c2: n_qty = st.number_input("鋼瓶數量 (支)", min_value=1, value=1, key=f"niq_{j}")
            st.markdown("<hr style='border: 1px dashed #BDC3C7;'>", unsafe_allow_html=True)
            if n_gas != "請選擇": collected_inv.append({"鋼瓶氣體種類": n_gas, "鋼瓶數量": n_qty})

        st.markdown('<div class="section-header">🧾 年度購買紀錄申報 (碳盤查活動數據)</div>', unsafe_allow_html=True)
        st.markdown("請依據本年度購買的單據登錄購買量並上傳佐證照。 **若本年度無購買請直接選擇「無購買」選項**。")
        
        # 🟢 淺藍色 Radio 介面
        e_buy_status = st.radio("調查期間有無購買", ["有購買", "無購買"], horizontal=True)
        collected_purchases = []
        
        if e_buy_status == "有購買":
            add_pur = st.number_input("➕ 新增本年度購買單據筆數", min_value=1, max_value=10, value=1)
            for k in range(int(add_pur)):
                st.markdown(f"**📦 購買單據 {k+1}**", unsafe_allow_html=True)
                pc1, pc2, pc3 = st.columns(3)
                with pc1: p_gas = st.selectbox("購買氣體種類", GAS_TYPES, key=f"pg_{k}")
                with pc2: p_date = st.date_input("購買日期", key=f"pd_{k}")
                with pc3: p_amt = st.number_input("購買量 (公斤)", min_value=0.0, step=0.1, key=f"pa_{k}")
                
                # 🟢 粗體提示字與全淺藍色的上傳區塊
                st.markdown("**📎 上傳購買單據** (限 JPG, PNG, PDF)")
                p_file = st.file_uploader("", type=['png', 'jpg', 'jpeg', 'pdf'], key=f"pf_{k}", label_visibility="collapsed")
                st.markdown("<hr style='border: 1px dashed #BDC3C7;'>", unsafe_allow_html=True)
                
                if p_gas != "請選擇" and p_amt > 0 and p_file:
                    collected_purchases.append({"鋼瓶氣體種類": p_gas, "購買日期": p_date.strftime("%Y-%m-%d"), "年度氣體鋼瓶購買量(公斤)": p_amt, "file": p_file})
                elif p_gas != "請選擇" and p_amt > 0 and not p_file:
                    st.warning(f"⚠️ 單據 {k+1} 尚未上傳佐證檔案，請補齊後再送出！")
        else:
            st.info("☑️ 您選擇了「無購買」，資料庫將記錄本年度無新增活動數據，請直接點擊下方按鈕前往確認。")

        st.markdown("<br>", unsafe_allow_html=True)

        if st.button("下一步：預覽確認資料", type="primary", use_container_width=True):
            if e_buy_status == "有購買" and not collected_purchases:
                st.error("❌ 您選擇了「有購買」，請至少完整填寫一筆單據並上傳附檔！")
            else:
                # 🟢 KeyError 修復：確保將使用者編輯後的資訊完整打包
                st.session_state.form_data = {
                    "dept": e_dept, "mgr": e_mgr, "campus": e_campus, 
                    "room": e_room, "mail": e_mail, "ext": e_ext, 
                    "status": e_buy_status, "inv": collected_inv, "pur": collected_purchases
                }
                st.session_state.step = "preview"
                st.rerun()

    # ================= 第二階段：預覽確認與送出 =================
    elif st.session_state.step == "preview":
        st.markdown('<div class="section-header">🔍 申報內容預覽確認</div>', unsafe_allow_html=True)
        d = st.session_state.form_data
        
        # 🟢 高質感預覽卡片
        st.markdown(f"""
        <div style="background: #FDFBF7; border: 2px solid #9DB4AB; border-radius: 10px; padding: 25px; margin-bottom: 25px;">
            <h3 style="color: #2C3E50; margin-top: 0;">📋 基本與庫存資訊確認</h3>
            <table style="width:100%; font-size:18px; line-height:2;">
                <tr><td style="width:120px; color:#7F8C8D;"><b>系所/老師</b></td><td>{d['dept']} / {d['mgr']} 老師</td></tr>
                <tr><td style="color:#7F8C8D;"><b>校區/門牌</b></td><td>{d['campus']} / {d['room']}</td></tr>
                <tr><td style="color:#7F8C8D;"><b>聯絡資訊</b></td><td>{d['ext']} / {d['mail']}</td></tr>
            </table>
            <br>
            <div style="background: #EBF5FB; padding: 10px 15px; border-radius: 5px; border-left: 5px solid #3498DB;">
                <b>最新庫存申報：</b> {', '.join([f"{i['鋼瓶氣體種類']} ({i['鋼瓶數量']}支)" for i in d['inv']]) if d['inv'] else '無庫存'}
            </div>
            
            <h3 style="color: #2C3E50; margin-top: 30px; border-top: 1px solid #EAECEE; padding-top: 20px;">🧾 年度購買狀況：<span style="color: {'#B03A2E' if d['status']=='有購買' else '#27AE60'};">{d['status']}</span></h3>
        </div>
        """, unsafe_allow_html=True)
        
        if d['status'] == "有購買":
            for idx, p in enumerate(d['pur']):
                st.markdown(f"**📌 購買單據 {idx+1}：** {p['鋼瓶氣體種類']} | {p['購買日期']} | **{p['年度氣體鋼瓶購買量(公斤)']} 公斤**")
                if p['file'].type in ['image/jpeg', 'image/png']: st.image(p['file'], width=300)
                st.markdown("<hr style='border: 1px solid #EAECEE;'>", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        col_btn1, col_btn2, col_btn3 = st.columns(3)
        
        with col_btn1:
            if st.button("⬅️ 回去修改資料", use_container_width=True):
                st.session_state.step = "input"; st.rerun()
                
        word_buf = create_report_docx(d, d['pur'], d['status'])
        
        def process_submission():
            with st.spinner("資料同步中，請稍候..."):
                gc = get_gc(); sh_main = gc.open_by_key(SHEET_ID)
                ws_inv = sh_main.worksheet('氣體鋼瓶資料庫'); ws_pur = sh_main.worksheet('氣體鋼瓶年度使用紀錄'); ws_trk = sh_main.worksheet('發送紀錄與金鑰')
                
                # 精準替換庫存
                all_inv_records = ws_inv.get_all_records()
                rows_to_del = [i + 2 for i, r in enumerate(all_inv_records) if str(r.get('系所', '')) == st.session_state.dept and str(r.get('實驗室老師', '')) == st.session_state.manager and str(r.get('氣體鋼瓶所在位置實驗室門牌', '')) == st.session_state.room]
                for r_idx in sorted(rows_to_del, reverse=True): safe_delete_row(ws_inv, r_idx); time.sleep(0.3)
                
                inv_to_append = [{"系所": d['dept'], "實驗室老師": d['mgr'], "校區": d['campus'], "氣體鋼瓶所在位置實驗室門牌": d['room'], "電子郵件": d['mail'], "分機": d['ext'], "鋼瓶氣體種類": item["鋼瓶氣體種類"], "鋼瓶數量": item["鋼瓶數量"], "建檔年度": f"{datetime.datetime.now().year-1911}年"} for item in d['inv']]
                if inv_to_append: append_dicts(ws_inv, inv_to_append)
                
                # 寫入年度購買
                pur_to_append = []
                now_str = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
                for p in d['pur']:
                    ext = os.path.splitext(p["file"].name)[1]
                    f_name = f"{d['campus']}_{d['dept']}_{d['mgr']}_{p['鋼瓶氣體種類']}_{p['購買日期']}_{p['年度氣體鋼瓶購買量(公斤)']}{ext}"
                    drive_link = upload_to_drive(p["file"], f_name)
                    pur_to_append.append({"填報時間": now_str, "系所": d['dept'], "實驗室老師": d['mgr'], "校區": d['campus'], "氣體鋼瓶所在位置實驗室門牌": d['room'], "鋼瓶氣體種類": p["鋼瓶氣體種類"], "購買日期": p["購買日期"], "年度氣體鋼瓶購買量(公斤)": p["年度氣體鋼瓶購買量(公斤)"], "購買單據連結": drive_link})
                if pur_to_append: append_dicts(ws_pur, pur_to_append)
                
                # 更新進度表
                ws_trk.update_cell(st.session_state.row_idx, 11, "已回報")
                ws_trk.update_cell(st.session_state.row_idx, 12, now_str)
                st.session_state.status = "已回報"

        with col_btn2:
            if st.download_button("📥 正式申報並下載申報檔留存", data=word_buf, file_name=f"申報檔_{d['dept']}_{d['mgr']}.docx", use_container_width=True):
                process_submission(); st.success("✅ 申報完成！"); time.sleep(2); st.rerun()
                
        with col_btn3:
            # 🟢 這裡的按鈕也會受到 CSS 橘色按鈕限制
            if st.button("🚀 正式申報不下載檔案留存", key="btn_submit_no_dl", use_container_width=True):
                process_submission(); st.success("✅ 申報完成！"); st.balloons(); time.sleep(2); st.rerun()

if __name__ == "__main__":
    main()