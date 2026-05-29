import streamlit as st
import google.generativeai as genai
import json
import os
import glob
import pandas as pd

# ==========================================
# 網頁基礎配置
# ==========================================
st.set_page_config(page_title="智能一頁式摘要生成器", layout="wide")

# ==========================================
# 暫存狀態 (Session State) 初始化
# ==========================================
if 'ai_generated_data' not in st.session_state:
    st.session_state.ai_generated_data = None
if 'report_generated' not in st.session_state:
    st.session_state.report_generated = False
if 'trigger_api_call' not in st.session_state:
    st.session_state.trigger_api_call = False

if 'auto_pass_text' not in st.session_state: st.session_state.auto_pass_text = ""
if 'auto_issue_text' not in st.session_state: st.session_state.auto_issue_text = ""
if 'init_period' not in st.session_state: st.session_state.init_period = ""
if 'init_target' not in st.session_state: st.session_state.init_target = ""
if 'init_item' not in st.session_state: st.session_state.init_item = ""

def clean_html(html_str):
    return "\n".join([line.strip() for line in html_str.split("\n")])

# ==========================================
# 側邊欄：BYOK 金鑰收納與使用指南
# ==========================================
with st.sidebar:
    st.header("🔑 系統連線設定")
    api_key = st.text_input(
        "請輸入 Google Gemini API Key：", 
        type="password", 
        placeholder="AIza...",
        help="金鑰僅於當前瀏覽器記憶體暫存，關閉網頁即自動銷毀，絕對安全。"
    )
    st.divider()
    with st.expander("💡 第一次使用？點此查看如何免費取得 API Key"):
        st.markdown("""
        本工作站串接特化大腦 **Gemini 3.1 Flash Lite**，您的免費 Google 帳號每日預設享有高達 **500 次** 呼叫配額！
        
        **取得金鑰 5 步驟：**
        1. 前往 👉 **[Google AI Studio 控制台](https://aistudio.google.com/)** 登入。
        2. 點擊左上角的 **「Get API key」**。
        3. 點擊 **「Create API key」** ➔ **「Create API key in new project」**。
        4. 複製以 **`AIza`** 開頭的字串。
        5. 貼入上方的密碼欄位中即可解鎖全自動排版引擎！
        """)

# ==========================================
# 頂級主標題招牌
# ==========================================
st.title("🤖 智能 AI 稽核助理：一頁式視覺戰情報告")
st.write("全自動生成頂級視覺底稿。支援流水線資料自動拋轉，完美達成「規劃 ➔ 執行 ➔ 報告」一氣呵成的數位稽核產線。")

st.markdown("---")

# ==========================================
# 📂 💡 優化：第一步 ➔ 專案流水線整合區 (置頂最直觀)
# ==========================================
st.subheader("📂 第一步：載入專案流水線 (Pipeline Integration)")
existing_json_files = glob.glob("Audit_Program_*.json")

if existing_json_files:
    project_options = [os.path.basename(f).replace("Audit_Program_", "").replace(".json", "") for f in existing_json_files]
    
    col_p1, col_p2 = st.columns([3, 1], gap="medium")
    with col_p1:
        selected_project = st.selectbox("📥 選擇 Fieldwork 階段成果 (將自動載入標題與缺失分流)：", ["-- 保持獨立手動輸入模式 --"] + project_options)
    
    with col_p2:
        st.write("") 
        if selected_project != "-- 保持獨立手動輸入模式 --":
            if st.button("🔄 提取專案數據", use_container_width=True, type="secondary"):
                # 💡 強制洗腦：清空所有歷史 AI 生成與手動編輯快取，防止專案交替污染
                st.session_state.ai_generated_data = None
                st.session_state.report_generated = False
                for k in list(st.session_state.keys()):
                    if k.startswith("edit_h_") or k.startswith("edit_i_"):
                        del st.session_state[k]

                with open(f"Audit_Program_{selected_project}.json", "r", encoding="utf-8") as f:
                    loaded_pack = json.load(f)
                
                meta_info = loaded_pack.get("meta", {})
                data_list = loaded_pack.get("data", [])
                
                # 自動拋轉表頭欄位
                st.session_state.init_period = meta_info.get("查核期間", "")
                st.session_state.init_target = meta_info.get("受查單位", "")
                st.session_state.init_item = meta_info.get("循環別", "")
                
                pass_accumulated = []
                issue_accumulated = []
                pass_idx = 1
                issue_idx = 1
                
                for row in data_list:
                    if not row.get('選擇', True): continue
                    if str(row.get('Audit Step', '')).strip() == "": continue
                    
                    status = row.get('有效控制', '未檢測')
                    wp = row.get('W/P', '無編碼')
                    risk = row.get('風險項目', '通用風險')
                    step = row.get('Audit Step', '')
                    findings = row.get('Findings', '無查核紀錄')
                    
                    if status == "OK":
                        packet = f"► [合規項目 {pass_idx}] (底稿索引: {wp})\n【風險控制點】：{risk}\n【查核程序標準】：{step}\n{findings}"
                        pass_accumulated.append(packet)
                        pass_idx += 1
                    elif status in ["部分有效", "失效"]:
                        packet = f"► [異常項目 {issue_idx}] (底稿索引: {wp})\n【風險控制點】：{risk}\n【查核程序標準】：{step}\n{findings}"
                        issue_accumulated.append(packet)
                        issue_idx += 1
                
                st.session_state.auto_pass_text = "\n\n".join(pass_accumulated)
                st.session_state.auto_issue_text = "\n\n".join(issue_accumulated)
                st.success(f"🎉 專案 【{selected_project}】 數據已精準就位！表單已全面更新。")
                st.rerun()
else:
    st.caption("ℹ️ 目前目錄下無暫存專案 JSON 檔，系統自動鎖定為「獨立手動輸入模式」。")

st.markdown("---")

# ==========================================
# 📌 第二步：報告標題設定 (接收流水線資料)
# ==========================================
st.subheader("📌 第二步：報告表頭資訊確認")
col_t1, col_t2, col_t3 = st.columns(3)
with col_t1:
    audit_period = st.text_input("📅 查核期間", value=st.session_state.init_period, placeholder="範例：2026Q1")
with col_t2:
    audit_target = st.text_input("🏢 受查單位", value=st.session_state.init_target, placeholder="範例：運籌採購總處")
with col_t3:
    audit_item = st.text_input("📑 查核項目", value=st.session_state.init_item, placeholder="範例：採購與付款循環控制作業")

st.markdown("---")

# ==========================================
# 📝 第三步：文字預覽與編修區 
# ==========================================
st.subheader("📝 第三步：查核結果檢閱與語氣設定")
col_left, col_right = st.columns([1, 1])

with col_left:
    st.markdown("**🟢 無缺失事項 (執行查核事項)**")
    pass_text = st.text_area(
        "請貼上或由流水線引入無異常之查核程序：", 
        value=st.session_state.auto_pass_text, 
        placeholder="背景範例貼入格式：\n1. 隨機抽核本季度25份供應商建檔審查表，確認均經權責主管核准並檢附完整資格證明文件。", 
        height=300,
        label_visibility="collapsed"
    )

with col_right:
    st.markdown("**🔴 有缺失事項 (實務觀察與優化建議)**")
    issue_text = st.text_area(
        "請貼上或由流水線引入發現之異常描述：", 
        value=st.session_state.auto_issue_text, 
        placeholder="背景範例貼入格式 (若無缺失請保持空白)：\n1. 經查有部分特定原物料採購案，因業務端緊急需求未於事前完成請購單申請...", 
        height=300,
        label_visibility="collapsed"
    )
        
st.markdown("<br>", unsafe_allow_html=True)
st.info("⚙️ 報告呈現方向設定 (動態 Prompt 注入)")
col_s1, col_s2 = st.columns(2)
with col_s1:
    tone_setting = st.selectbox("🎯 呈報對象語氣要求", [
        "精煉大局觀 (公司整體管理層級，強調大局衝擊與治理維度)", 
        "建設性導向 (友善受查單位溝通，強調流程優化與共好防禦)"
    ])
with col_s2:
    focus_setting = st.selectbox("🚩 核心側重點", [
        "著重合規與內控防禦深度 (強化控制點嚴謹度/舉證軌跡)",
        "著重營運持續與流程效率 (優化作業瓶頸/降本增效)"
    ])

# ==========================================
# 🚀 啟動 AI 呼叫引擎 (內建智慧整併原則)
# ==========================================
st.markdown("<br>", unsafe_allow_html=True)
if st.button("🚀 啟動 AI 深度生成視覺戰情報告", type="primary", use_container_width=True):
    if not api_key:
        st.error("❌ 系統攔截：請先於左側欄位貼入您的 Google Gemini API Key！")
    elif not audit_period.strip() or not audit_target.strip() or not audit_item.strip():
        st.warning("⚠️ 系統攔截：請完整填寫上方『第二步』的三個表頭欄位！")
    elif not pass_text.strip() and not issue_text.strip():
        st.warning("⚠️ 系統攔截：請至少於『第三步』填入『無缺失事項』或『有缺失事項』！")
    else:
        st.session_state.trigger_api_call = True

if st.session_state.trigger_api_call:
    with st.spinner("✨ 智能助理正調用極速 3.1 Flash Lite 引擎進行排版，請稍候..."):
        try:
            for k in list(st.session_state.keys()):
                if k.startswith("edit_h_") or k.startswith("edit_i_"):
                    del st.session_state[k]

            genai.configure(api_key=api_key)
            model = genai.GenerativeModel(
                'gemini-3.1-flash-lite',
                generation_config={"response_mime_type": "application/json", "temperature": 0.1}
            )

            eff_pass = pass_text.strip()
            eff_iss = issue_text.strip()

            prompt = f"""
            請身為資深企業內部稽核主管，嚴格依循 JSON 輸出審查結果。
            語氣：【{tone_setting}】 / 側重點：【{focus_setting}】。
            
            🚨 【最高審查準則與項目智能整併指令】
            1. 絕對禁止憑空捏造事實！若輸入為空，回傳空陣列 `[]`。
            2. 【項目獨立與合併原則】：輸入的文本中可能包含了多個獨立的查核項目（已標註如 [合規項目 1]、底稿索引 A-1 等）。
               - 請敏銳辨識不同的底稿索引。
               - 若不同索引的缺失或合規事實屬於「不同業務性質」，請將其『分列為獨立的要點 (highlights/issues)』。
               - 若多個索引反映的是「同一類型的系統性問題或合規成就」，允許你發揮主管高度，將其『合併歸納』為一點，以展現總結視角。
            3. 🚨防幻覺鎖：若使用者提供的描述極度簡短，你的優化建議必須基於「通用內控健全原則」，嚴禁編造法規條號、系統名稱或具體數字！

            【任務一輸入文本 (執行查核事項 / 無缺失)】
            {eff_pass if eff_pass else "無輸入內容。"}

            【任務二輸入文本 (有缺失事項 / 待改善議題)】
            {eff_iss if eff_iss else "無輸入內容。"}

            請產出標準 JSON 結構：
            - highlights 陣列：1. 主旨大標題(精煉不超過12字) 2. 合規描述(約30-45字)。
            - issues 陣列：1. 議題標題(精煉不超過14字) 2. 實務觀察(提煉核心事實) 3. 優化建議(提出建設性對策)。

            嚴格輸出格式：
            {{"highlights": [{{"title": "大標題", "desc": "合規描述"}}], "issues": [{{"title": "議題標題", "observation": "觀察內容", "recommendation": "建議內容"}}]}}
            """
            
            res = model.generate_content(prompt)
            st.session_state.ai_generated_data = json.loads(res.text)
            st.session_state.report_generated = True
            
        except Exception as e:
            st.error(f"❌ 呼叫異常：{str(e)}")
        finally:
            st.session_state.trigger_api_call = False

# ==========================================
# 📊 報告優先展示與自適應編修區 
# ==========================================
if st.session_state.report_generated and st.session_state.ai_generated_data:
    st.divider()
    data = st.session_state.ai_generated_data
    
    p_val = audit_period.strip()
    t_val = audit_target.strip()
    i_val = audit_item.strip()
    full_title = f"{p_val} {t_val} {i_val}：整體查核總覽"

    icons = ["📋", "🔍", "🛡️", "📊", "⚙️", "💡", "📑", "📌"]
    highlights = data.get("highlights", [])
    issues = data.get("issues", [])

    total_items = len(highlights) + len(issues)
    is_compact = total_items > 6

    card_p_top = "8px" if is_compact else "10px"
    card_p_side = "12px" if is_compact else "14px"
    font_title = "16px" if is_compact else "18px"
    font_body = "14px" if is_compact else "15px"
    margin_b = "12px" if is_compact else "24px"
    issue_p = "14px" if is_compact else "18px"
    issue_mb = "8px" if is_compact else "12px"

    highlights_html = ""
    if highlights:
        for idx, h in enumerate(highlights):
            icon = icons[idx % len(icons)]
            h_title = st.session_state.get(f"edit_h_t_{idx}", h.get('title', ''))
            h_desc = st.session_state.get(f"edit_h_d_{idx}", h.get('desc', ''))
            
            highlights_html += f"""
            <div style="flex: 1; min-width: 200px; border: 1px solid #cbd5e1; border-radius: 6px; overflow: hidden; background: #ffffff; box-shadow: 0 1px 2px rgba(0,0,0,0.05); display: flex; flex-direction: column;">
                <div style="background-color: #0284c7; color: #ffffff; padding: {card_p_top} {card_p_side}; font-weight: 700; font-size: {font_title}; display: flex; align-items: center; gap: 6px;">
                    <span>{icon}</span> <span>{h_title}</span>
                </div>
                <div style="padding: {card_p_side}; flex-grow: 1;">
                    <p style="color: #334155; font-size: {font_body}; line-height: 1.5; margin: 0;">{h_desc}</p>
                </div>
            </div>
            """
    else:
        highlights_html = """
        <div style="background-color: #f8fafc; border: 1px dashed #cbd5e1; padding: 12px 20px; border-radius: 6px; width: 100%; color: #64748b; font-size: 15px;">
            ℹ️ 本次報告未記錄常規遵循事項。
        </div>
        """
        
    issues_html = ""
    if issues:
        for idx, item in enumerate(issues):
            i_title = st.session_state.get(f"edit_i_t_{idx}", item.get('title', ''))
            i_obs = st.session_state.get(f"edit_i_o_{idx}", item.get('observation', ''))
            i_rec = st.session_state.get(f"edit_i_r_{idx}", item.get('recommendation', ''))
            
            issues_html += f"""
            <div style="flex: 1; min-width: 300px; border: 1px solid #fde68a; border-radius: 8px; background-color: #fffbeb; padding: {issue_p}; margin-bottom: {issue_mb}; box-shadow: 0 2px 4px rgba(217,119,6,0.05);">
                <div style="border-bottom: 2px solid #f59e0b; padding-bottom: 8px; margin-bottom: 10px; display: flex; align-items: center; gap: 6px;">
                    <span style="font-size: 18px;">💡</span>
                    <span style="font-size: {font_title}; font-weight: 700; color: #92400e;">{i_title}</span>
                </div>
                <div style="background-color: #ffffff; border-left: 4px solid #f59e0b; padding: 10px 12px; margin-bottom: 10px; border-radius: 0 4px 4px 0;">
                    <span style="font-size: {font_body}; font-weight: 700; color: #d97706; display: block; margin-bottom: 4px;">【實務觀察】</span>
                    <p style="color: #451a03; font-size: {font_body}; line-height: 1.5; margin: 0;">{i_obs}</p>
                </div>
                <div style="background-color: #ffffff; border-left: 4px solid #0d9488; padding: 10px 12px; border-radius: 0 4px 4px 0;">
                    <span style="font-size: {font_body}; font-weight: 700; color: #0f766e; display: block; margin-bottom: 4px;">【優化建議】</span>
                    <p style="color: #134e5e; font-size: {font_body}; line-height: 1.5; margin: 0;">{i_rec}</p>
                </div>
            </div>
            """
        
        issues_section = f"""
        <div style="font-size: 17px; font-weight: 700; color: #b45309; margin-bottom: 10px; border-bottom: 2px solid #d97706; padding-bottom: 4px;">
            實務觀察與優化建議
        </div>
        <div style="display: flex; gap: 12px; flex-wrap: wrap;">
            {issues_html}
        </div>
        """
    else:
        issues_section = f"""
        <div style="font-size: 17px; font-weight: 700; color: #0f766e; margin-bottom: 10px; border-bottom: 2px solid #0d9488; padding-bottom: 4px;">
            實務觀察與優化建議
        </div>
        <div style="background-color: #f0fdf4; border-left: 4px solid #16a34a; padding: 14px 20px; border-radius: 0 8px 8px 0; margin-top: 10px;">
            <span style="color: #16a34a; font-weight: 700; font-size: 16px;">🎉 本次專案深度查核結果：合規良好</span>
            <p style="color: #15803d; font-size: {font_body}; margin: 4px 0 0 0;">經實質抽核與程序檢視，受查標的於查核期間內未發現重大內部控制程序缺陷或異常發現，運作機制嚴謹順暢。</p>
        </div>
        """

    raw_html = f"""
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #1e293b; background: #ffffff; padding: 5px 10px;">
        <div class="gradient-header" style="background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 100%); border-radius: 8px; padding: 12px 18px; margin-bottom: 16px;">
            <h1 style="color: #ffffff; font-size: 22px; font-weight: 700; margin: 0; letter-spacing: 0.5px;">{full_title}</h1>
        </div>
        
        <div style="font-size: 17px; font-weight: 700; color: #0f766e; margin-bottom: 10px; border-bottom: 2px solid #0d9488; padding-bottom: 4px;">
            執行查核事項
        </div>
        <div style="display: flex; gap: 10px; margin-bottom: {margin_b}; flex-wrap: wrap;">
            {highlights_html}
        </div>
        
        {issues_section}
    </div>
    """

    st.markdown(clean_html(raw_html), unsafe_allow_html=True)
    st.write("")

    pdf_html_source = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>{full_title}</title>
        <style>
            @page {{ size: A4 landscape; margin: 10mm; }}
            body, html {{ 
                font-family: 'Microsoft JhengHei', 'Segoe UI', sans-serif; 
                background-color: #ffffff; padding: 0; margin: 0; 
                -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important;
            }}
            .gradient-header {{
                background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 100%) !important;
                -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important;
            }}
        </style>
    </head>
    <body>
        {raw_html}
    </body>
    </html>
    """

    col_btn, col_empty = st.columns([1, 2])
    with col_btn:
        st.download_button(
            label="📥 匯出成高畫質單頁戰報檔 (.html)",
            data=clean_html(pdf_html_source).encode('utf-8'), 
            file_name=f"{full_title}.html", 
            mime="text/html", 
            type="primary", 
            use_container_width=True,
            help="點擊開啟後按 Ctrl+P，即可完美另存 PDF 報告。"
        )

    st.markdown("---")

    with st.expander("🛠️ 報告內容文字微調與手動覆核 (若需調整字眼請點此展開)"):
        st.caption("前端快取防護啟用中：在此處敲擊鍵盤修改任何文字皆 0 消耗後台 API 配額。")
        
        if highlights:
            st.markdown("**🔹 執行查核事項微調**")
            for idx, h in enumerate(highlights):
                c1, c2 = st.columns([1, 3])
                with c1:
                    h["title"] = st.text_input(f"標題 #{idx+1}", value=h.get("title", ""), key=f"edit_h_t_{idx}")
                with c2:
                    h["desc"] = st.text_input(f"內文遵循描述 #{idx+1}", value=h.get("desc", ""), key=f"edit_h_d_{idx}")
                
        if issues:
            st.markdown("<br>**🔸 實務觀察與優化建議微調**", unsafe_allow_html=True)
            for idx, item in enumerate(issues):
                st.markdown(f"**議題 #{idx+1}**")
                item["title"] = st.text_input(f"議題標題 #{idx+1}", value=item.get("title", ""), key=f"edit_i_t_{idx}")
                c_obs, c_rec = st.columns(2)
                with c_obs:
                    item["observation"] = st.text_area(f"實務觀察 #{idx+1}", value=item.get("observation", ""), height=70, key=f"edit_i_o_{idx}")
                with c_rec:
                    item["recommendation"] = st.text_area(f"優化建議 #{idx+1}", value=item.get("recommendation", ""), height=70, key=f"edit_i_r_{idx}")
        else:
            st.success("✨ 本次查核未產生任何待改善議題，無需進行下方文字微調！")

        st.session_state.ai_generated_data = data
