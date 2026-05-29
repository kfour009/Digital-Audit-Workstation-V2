import streamlit as st
import os

# ==========================================
# 1. 全域 UX 優化：懸浮置中側邊欄開關 (CSS 覆寫)
# ==========================================
# 這段 CSS 會強制把 Streamlit 隱藏在最上方的 < / > 按鈕，直接拉到螢幕垂直正中央，並保持懸浮
floating_sidebar_css = """
<style>
    /* 側邊欄收合按鈕 (X) 垂直置中懸浮 */
    [data-testid="stSidebarCollapseButton"] {
        position: fixed !important;
        top: 50vh !important;
        transform: translateY(-50%) !important;
        z-index: 99999 !important;
        background-color: #f8fafc !important;
        border-radius: 50% !important;
        box-shadow: 0 2px 6px rgba(0,0,0,0.15) !important;
    }
    /* 側邊欄展開按鈕 (>) 垂直置中懸浮 */
    [data-testid="collapsedControl"] {
        top: 50vh !important;
        transform: translateY(-50%) !important;
        z-index: 99999 !important;
        background-color: #ffffff !important;
        border-radius: 0 50% 50% 0 !important;
        box-shadow: 3px 0 8px rgba(0,0,0,0.15) !important;
    }
</style>
"""

# ==========================================
# 2. 頁面清單定義 (次世代導覽架構)
# ==========================================
pg_home = st.Page("app_v2.py", title="🏠 數位轉型戰情室", default=True)

# 核心工具組
pg1 = st.Page("pages/1_Word圖片提取.py", title="📄 Word 圖片提取", icon="📝")
pg2 = st.Page("pages/2_PDF圖片提取.py", title="🖼️ PDF 圖片提取", icon="📸")
pg3 = st.Page("pages/3_PDF與圖檔合併器.py", title="🔗 PDF 頁面整併", icon="📎")

# 轉檔引擎
pg4 = st.Page("pages/4_萬能格式轉PDF_直式.py", title="📏 直式底稿轉 PDF", icon="👔")
pg5 = st.Page("pages/5_萬能格式轉PDF_橫式.py", title="📐 橫式底稿轉 PDF", icon="💼")

# 🔥 核心：資料整併樞紐
pg6 = st.Page("pages/6_CSA智能整併系統.py", title="📊 萬用表單智能整併", icon="📈") # 定位為通用工具

# 🔥 AI 核心產線 (一條龍精煉版 3 步驟)
pg8 = st.Page("pages/8_管理辦法解析_測試版.py", title="🧠 規章解析 Agent", icon="⚡")
pg9 = st.Page("pages/9_智能底稿比對_測試版.py", title="🕵️ 智能底稿比對", icon="🔍") 
pg7 = st.Page("pages/7_智能一頁式摘要生成器.py", title="🤖 戰情報告生成", icon="✨")

# 建立左側導覽列選單 (將 CSA 移出 AI 產線)
pg = st.navigation({
    "管理中心": [pg_home],
    "AI 智慧稽核產線": [pg8, pg9, pg7],
    "通用資料與行政工具箱": [pg6, pg4, pg5, pg3, pg1, pg2]
})

# ==========================================
# 3. 戰情中心版面渲染 
# ==========================================
if pg == pg_home:
    st.set_page_config(page_title="數位稽核戰情室", layout="wide", initial_sidebar_state="expanded")
    st.markdown(floating_sidebar_css, unsafe_allow_html=True) # 注入懸浮側邊欄魔法
    
    st.title("⚡ 內部稽核：數位轉型戰情中心")
    st.markdown("整合 LLM 大語言模型與 RPA 自動化流程，打造新世代資料驅動稽核生態系。")
    
    st.subheader("🎯 數位轉型效益追蹤 (動態預估值)")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric(label="本月節省工時", value="150 hrs", delta="↑ 22% 較上月", delta_color="normal")
    m2.metric(label="底稿比對自動化率", value="65%", delta="Fieldwork 引擎上線", delta_color="normal")
    m3.metric(label="規章與佐證解析量", value="128 件", delta="↑ 產能滿載")
    m4.metric(label="系統妥善率", value="100%", delta="穩定運行中")
    
    st.divider()

    col_left, col_right = st.columns([1, 1], gap="large")
    
    with col_left:
        st.header("🧠 啟動 AI 稽核一條龍")
        st.info("透過 Gemini 3.1 Flash Lite 引擎，完成從規劃到報告的智能自動化。")
        
        # 產線精煉為 3 步驟
        if st.button("① 規劃階段 ➔ 規章解析 (產出 Audit Program)", use_container_width=True, type="primary"): st.switch_page(pg8)
        if st.button("② 執行階段 ➔ 智能底稿比對 (Fieldwork 驗證)", use_container_width=True, type="primary"): st.switch_page(pg9)
        if st.button("③ 報告階段 ➔ 智能一頁式戰情報告", use_container_width=True, type="primary"): st.switch_page(pg7)

    with col_right:
        st.header("🛠️ 通用行政與資料工具")
        st.success("取代傳統人工作業，將繁瑣的整併、轉檔與素材提取全自動化。")
        
        # 將 CSA 獨立置頂於工具箱
        if st.button("📊 萬用表單智能整併系統 (自評內控 / 跨檔抓取)", use_container_width=True, type="secondary"): st.switch_page(pg6)
        
        c1, c2 = st.columns(2)
        with c1:
            if st.button("📝 Word 圖片提取", use_container_width=True): st.switch_page(pg1)
            if st.button("📸 PDF 圖片提取", use_container_width=True): st.switch_page(pg2)
            if st.button("📎 PDF 頁面合併", use_container_width=True): st.switch_page(pg3)
        with c2:
            if st.button("📏 直式底稿轉 PDF", use_container_width=True): st.switch_page(pg4)
            if st.button("📐 橫式底稿轉 PDF", use_container_width=True): st.switch_page(pg5)

    st.sidebar.divider()
    if os.name == 'nt':
        st.sidebar.success("🟢 本機全速運轉模式")
    else:
        st.sidebar.info("☁️ 雲端輕量防護模式")
else:
    pg.run()