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

# ================= 系統與全域參數設定 =================
st.set_page_config(page_title="實驗室氣體鋼瓶資料回報", page_icon="💨", layout="wide")

SHEET_ID = '1Hw4rXo4ww7O9YXTwoUJeWioO5ZzM_bivRcLLpOl26DY'
FOLDER_ID = '1cpj6l-zTgqLwphMFzLvvr53-Alf60idN'

GAS_TYPES = ["二氧化碳", "甲烷", "乙炔", "一氧化二氮(笑氣)"]
# 🚀 精準修正：補齊缺失的變數，徹底解決 NameError 與 Missing Submit Button 的骨牌錯誤
CAMPUS_OPTS = ["蘭潭校區", "民雄校區", "新民校區", "林森校區"]

# ================= 莫蘭迪風格 CSS & UI =================
def apply_morandi_theme():
    st.markdown("""
    <style>
        .title-box { background-color: #5C6B73; color: #FFFFFF; padding: 20px; border-radius: 10px; text-align: center; margin-bottom: 20px; }
        .section-header { background-color: #9DB4AB; color: #FFFFFF; padding: 10px 15px; border-radius: 8px; font-size: 20px; font-weight: bold; margin-top: 25px; margin-bottom: 15px; }
        .card { background-color: #FDFBF7; border: 1px solid #DCE4E3; padding: 15px 20px; border-radius: 8px; margin-bottom: 15px; }
        button[kind="primary"] { background-color: #27AE60 !important; color: #FFFFFF !important; font-weight: bold !important; font-size: 16px !important; padding: 10px 20px !important;}
        .upload-area { background-color: #E3F2FD; padding: 15px; border-radius: 8px; border: 1px dashed #90CAF9; }
        [data-testid="stSidebar"], [data-testid="collapsedControl"], [data-testid="stHeader"] { display: none !important; }
        .block-container { padding-top: 2rem !important; }
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
def upload_to_drive(uploaded_file, file_name):
    creds = get_credentials()
    drive_service = build('drive', 'v3', credentials=creds)
    file_metadata = {'name': file_name, 'parents': [FOLDER_ID]}
    media = MediaIoBaseUpload(io.BytesIO(uploaded_file.getvalue()), mimetype=uploaded_file.type, resumable=True)
    file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    return f"https://drive.google.com/file/d/{file.get('id')}/view?usp=sharing"

# ================= 產生 Word 申報表 =================
def create_report_docx(data):
    doc = Document()
    title = doc.add_heading('溫室氣體盤查 - 實驗室氣體鋼瓶申報表', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f"填報日期：{datetime.datetime.now().strftime('%Y/%m/%d')}")
    
    table = doc.add_table(rows=6, cols=2)
    table.style = 'Table Grid' # 標準框線表格
    
    rows_data = [
        ("系所 / 老師", f"{data['dept']} / {data['mgr']}"),
        ("校區 / 門牌", f"{data['campus']} / {data['room']}"),
        ("聯絡分機", data['ext']),
        ("氣體種類", data['gas']),
        ("購買狀況", data['status']),
        ("購買量(kg) / 日期", f"{data['qty']} kg / {data['date']}" if data['status'] == '有購買' else "無")
    ]
    for i, (label, val) in enumerate(rows_data):
        table.cell(i, 0).text = label
        table.cell(i, 1).text = str(val)
        
    doc.add_paragraph("\n申報聲明：本人確認以上資料屬實，並依規定完成申報。")
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

# ================= 主程式 =================
def main():
    apply_morandi_theme()
    
    # 📢 宣導標語
    st.markdown('<div style="background-color: #F8D7DA; padding: 15px; border-radius: 10px; border: 1px solid #F5C2C7; text-align: center; font-weight: bold; color: #842029; font-size: 20px; margin-bottom: 20px;">📢 請「誠實申報」，以保障單位及自身權益！</div>', unsafe_allow_html=True)

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
        
        df_inv = pd.DataFrame(fetch_main_inventory())
        if not df_inv.empty:
            mask = (df_inv['系所'].astype(str) == st.session_state.dept) & \
                   (df_inv['實驗室老師'].astype(str) == st.session_state.manager)
            st.session_state.inventory = df_inv[mask].to_dict('records')
        else: st.session_state.inventory = []
        
        st.session_state.step = "input"

    is_readonly = (st.session_state.status == "已回報")
    
    st.markdown(f"""
    <div class="title-box">
        <h2 style="margin: 0;">國立嘉義大學 氣體鋼瓶年度盤查與回報系統</h2>
    </div>
    """, unsafe_allow_html=True)
    
    if is_readonly: 
        st.success("✅ 貴實驗室已完成本期回報，目前為唯讀檢視模式。")
        st.stop()

    # 取出預設值
    inv_data = st.session_state.inventory[0] if st.session_state.inventory else {}
    def_room = inv_data.get("氣體鋼瓶所在位置實驗室門牌", "")
    def_gas = inv_data.get("鋼瓶氣體種類", "請選擇")
    def_ext = inv_data.get("分機", "")
    def_mail = inv_data.get("電子郵件", st.session_state.email)

    # ================= 填報階段 =================
    if st.session_state.step == "input":
        st.markdown('<div class="section-header">🏭 實驗室基本資訊</div>', unsafe_allow_html=True)
        with st.form("input_form"):
            c1, c2, c3 = st.columns(3)
            with c1: e_dept = st.text_input("系所", value=st.session_state.dept)
            with c2: e_mgr = st.text_input("實驗室老師", value=st.session_state.manager)
            with c3: e_campus = st.selectbox("校區", CAMPUS_OPTS, index=CAMPUS_OPTS.index(st.session_state.campus) if st.session_state.campus in CAMPUS_OPTS else 0)
            
            c4, c5 = st.columns(2)
            with c4: e_mail = st.text_input("電子郵件", value=def_mail)
            with c5: e_ext = st.text_input("分機", value=def_ext)

            st.markdown('<div class="section-header">🧾 溫室氣體盤查範疇之氣體鋼瓶購買情形申報</div>', unsafe_allow_html=True)
            
            g1, g2 = st.columns(2)
            with g1: 
                gas_idx = GAS_TYPES.index(def_gas) if def_gas in GAS_TYPES else 0
                e_gas = st.selectbox("鋼瓶氣體種類", GAS_TYPES, index=gas_idx)
            with g2: e_room = st.text_input("氣體鋼瓶所在位置實驗室門牌", value=def_room, placeholder="例如：A32-101")
            
            st.markdown("---")
            e_buy_status = st.radio("調查期間有無購買", ["有購買", "無購買"], horizontal=True)
            
            buy_date = "-"; buy_qty = 0.0; uploaded_file = None
            
            if e_buy_status == "有購買":
                bc1, bc2 = st.columns(2)
                with bc1: 
                    st.markdown("購買日期 <span style='color:gray; font-size:14px;'>(💡 請以發票日期為主)</span>", unsafe_allow_html=True)
                    buy_date = st.date_input("", label_visibility="collapsed").strftime("%Y-%m-%d")
                with bc2:
                    st.markdown('年度氣體鋼瓶購買量 <span style="color:red; font-weight:bold;">(公斤)</span>', unsafe_allow_html=True)
                    buy_qty = st.number_input("", min_value=0.0, step=0.1, label_visibility="collapsed")
                
                st.markdown('<div class="upload-area">', unsafe_allow_html=True)
                uploaded_file = st.file_uploader("📎 上傳購買單據 (限 JPG, PNG, PDF)", type=['png', 'jpg', 'jpeg', 'pdf'])
                st.markdown('</div>', unsafe_allow_html=True)
            else:
                st.checkbox("☑️ 盤查期間無購買氣體鋼瓶", value=True, disabled=True)

            submitted = st.form_submit_button("下一步：預覽確認資料", type="primary", use_container_width=True)
            if submitted:
                if e_buy_status == "有購買" and (not uploaded_file or buy_qty <= 0):
                    st.error("❌ 您選擇了「有購買」，請務必填寫購買量並上傳單據！")
                else:
                    st.session_state.form_data = {
                        "dept": e_dept, "mgr": e_mgr, "campus": e_campus, "mail": e_mail, "ext": e_ext,
                        "gas": e_gas, "room": e_room, "status": e_buy_status,
                        "date": buy_date, "qty": buy_qty, "file": uploaded_file
                    }
                    st.session_state.step = "preview"
                    st.rerun()

    # ================= 預覽確認與送出階段 =================
    elif st.session_state.step == "preview":
        st.markdown('<div class="section-header">🔍 申報內容預覽確認</div>', unsafe_allow_html=True)
        d = st.session_state.form_data
        
        st.markdown(f"""
        <div class="card">
            <h4>基本資訊</h4>
            <p><b>系所：</b>{d['dept']} | <b>老師：</b>{d['mgr']} | <b>校區：</b>{d['campus']}</p>
            <p><b>電郵：</b>{d['mail']} | <b>分機：</b>{d['ext']} | <b>門牌：</b>{d['room']}</p>
            <hr>
            <h4>盤查資訊</h4>
            <p><b>氣體種類：</b>{d['gas']}</p>
            <p><b>購買狀況：</b>{d['status']}</p>
            <p><b>購買日期：</b>{d['date']}</p>
            <p><b>購買量：</b>{d['qty']} kg</p>
        </div>
        """, unsafe_allow_html=True)
        
        if d['status'] == "有購買" and d['file']:
            st.info(f"📎 附檔：已準備上傳單據 ({d['file'].name})")
            if d['file'].type in ['image/jpeg', 'image/png']: st.image(d['file'], width=300)

        st.markdown("<br>", unsafe_allow_html=True)
        col_btn1, col_btn2, col_btn3 = st.columns(3)
        
        with col_btn1:
            if st.button("⬅️ 資料修正", use_container_width=True):
                st.session_state.step = "input"; st.rerun()
                
        word_buf = create_report_docx(d)
        
        # 定義送出邏輯 (精準更新座標)
        def process_submission():
            with st.spinner("資料同步中，請稍候..."):
                gc = get_gc()
                sh_main = gc.open_by_key(SHEET_ID)
                ws_inv = sh_main.worksheet('氣體鋼瓶資料庫')
                ws_pur = sh_main.worksheet('氣體鋼瓶年度使用紀錄')
                ws_trk = sh_main.worksheet('發送紀錄與金鑰')
                
                # 1. 精準更新庫存 (只改有變動的欄位，不刪除重建)
                all_inv = ws_inv.get_all_records()
                target_idx = next((i + 2 for i, r in enumerate(all_inv) if str(r.get('系所')) == st.session_state.dept and str(r.get('實驗室老師')) == st.session_state.manager), None)
                
                if target_idx:
                    # 使用 batch_update 減少 API 請求
                    updates = [
                        {'range': f'D{target_idx}', 'values': [[d['room']]]},
                        {'range': f'E{target_idx}', 'values': [[d['mail']]]},
                        {'range': f'F{target_idx}', 'values': [[d['ext']]]},
                        {'range': f'G{target_idx}', 'values': [[d['gas']]]}
                    ]
                    ws_inv.batch_update(updates)
                else:
                    # 若為全新無資料，則附加上去
                    ws_inv.append_row([d['dept'], d['mgr'], d['campus'], d['room'], d['mail'], d['ext'], d['gas'], 1, f"{datetime.datetime.now().year-1911}年"])

                # 2. 寫入購買紀錄
                drive_link = ""
                if d['status'] == '有購買' and d['file']:
                    ext = os.path.splitext(d['file'].name)[1]
                    f_name = f"{d['campus']}_{d['dept']}_{d['mgr']}_{d['gas']}_{d['qty']}_{d['date']}{ext}"
                    drive_link = upload_to_drive(d['file'], f_name)
                    ws_pur.append_row([
                        datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S"),
                        d['dept'], d['mgr'], d['campus'], d['room'], d['gas'], d['date'], d['qty'], drive_link
                    ])
                
                # 3. 更新追蹤表狀態 (精準更新)
                ws_trk.update_cell(st.session_state.row_idx, 11, "已回報")
                ws_trk.update_cell(st.session_state.row_idx, 12, datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S"))
                
                st.session_state.status = "已回報"

        with col_btn2:
            if st.download_button("📥 正式申報並下載申報檔留存", data=word_buf, file_name=f"申報檔_{d['dept']}_{d['mgr']}.docx", use_container_width=True):
                process_submission()
                st.success("✅ 申報完成！")
                time.sleep(2); st.rerun()
                
        with col_btn3:
            if st.button("🚀 正式申報不下載檔案留存", type="primary", use_container_width=True):
                process_submission()
                st.success("✅ 申報完成！")
                st.balloons()
                time.sleep(2); st.rerun()

if __name__ == "__main__":
    main()