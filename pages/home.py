import streamlit as st

st.markdown("""
<style>
    .info-box {
        background-color: #EBF5FB; border-left: 5px solid #3498DB; 
        padding: 20px; border-radius: 5px; margin-bottom: 20px;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="info-box">
    <h4>👋 您好！歡迎來到本校「溫室氣體盤查填報系統」</h4>
    <p>請從 <strong>👈 左側選單</strong> 選擇您要執行的業務：</p>
    <ul>
        <li><strong>燃油填報作業</strong>：公務車輛、發電機、除草機、農用機具等油料使用申報。</li>
        <li><strong>冷媒填報作業</strong>：冷氣機、冷凍/藏設備等冷媒填充維修申報。</li>
        <li><strong>AI 設備油單辨識</strong>：上傳加油發票，由 AI 自動判讀金額與油品。</li>
    </ul>
</div>
""", unsafe_allow_html=True)

st.info("💡 提示：如果左側沒有看到選單，請點擊左上角的箭頭展開側邊欄 (Sidebar)。")

# 這裡未來還可以放一些總覽圖表，例如全校目前燃油總量進度條等等