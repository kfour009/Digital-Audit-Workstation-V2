import streamlit as st
import google.generativeai as genai
import PyPDF2
import pandas as pd
import json
import re
import io
import os
import glob

# ==========================================
# 網頁基礎配置
# ==========================================
st.set_page_config(page_title="管理辦法解析 Agent", page_icon="🧠", layout="wide")
st.title("🧠 智能查核程序生成器 (企業級一條龍版)")
st.write("整合 Maker-Checker 雙重 AI 驗證機制，內建智慧 W/P 索引編碼與 JSON 專案持久化暫存倉庫。")

# ==========================================
# 側邊欄：API 設定
# ==========================================
with st.sidebar:
    st.header("🔑 系統連線設定")
    api_key = st.text_input("請輸入 Google Gemini API Key:", type="password")
    st.info("💡 系統指定使用 Gemini Flash 模型，具備百萬級 Token 窗口，完美勝任長文本雙重覆核。")

# ==========================================
# 核心自動化函數：智慧 W/P 編碼索引引擎
# ==========================================
def apply_smart_indexing(df, cycle_name):
    letters = "BCDEFGHIJKLMNOPQRSTUVWXYZ"
    current_letter_idx = 0
    risk_group_mapping = {}
    group_counters = {}
    indices = []
    wp_labels = []
    numbers = []
    current_num = 1 
    
    for i in range(len(df)):
        row = df.iloc[i]
        is_selected = row.get('選擇', True)
        
        # 💡 修復 2：只要未勾選，立刻重置為「尚未覆核」
        if not is_selected:
            df.at[df.index[i], 'AI可信度評分'] = "尚未覆核"
            df.at[df.index[i], 'AI覆核意見'] = "請點擊下方按鈕啟動雙重驗證"
            indices.append("")
            wp_labels.append("")
            numbers.append("")
            continue
            
        if i == 0:
            indices.append('A')
            wp_labels.append('A-1')
            numbers.append(current_num)
            current_num += 1
            continue
            
        risk_item = row['風險項目']
        if not risk_item or pd.isna(risk_item):
            risk_item = "未分類風險控制"
            
        if risk_item not in risk_group_mapping:
            if current_letter_idx < len(letters):
                risk_group_mapping[risk_item] = letters[current_letter_idx]
                current_letter_idx += 1
            else:
                risk_group_mapping[risk_item] = "Z"
            group_counters[risk_item] = 1
        else:
            group_counters[risk_item] += 1
            
        letter = risk_group_mapping[risk_item]
        count = group_counters[risk_item]
        
        indices.append(letter)
        wp_labels.append(f"{letter}-{count}")
        numbers.append(current_num)
        current_num += 1
        
    df['編號'] = indices
    df['W/P'] = wp_labels
    df['No'] = numbers
    df['循環別'] = cycle_name 
    return df

# ==========================================
# 核心函數：PDF 萃取與 JSON 清洗
# ==========================================
def extract_text_from_pdfs(uploaded_files):
    all_text = ""
    for file in uploaded_files:
        try:
            reader = PyPDF2.PdfReader(io.BytesIO(file.read()))
            text = f"--- 以下為【{file.name}】的內容 ---\n"
            for page in reader.pages:
                text += page.extract_text() + "\n"
            all_text += text + "\n\n"
        except Exception as e:
            st.error(f"❌ 讀取 {file.name} 失敗: {str(e)}")
    return all_text

def clean_json_string(raw_response):
    cleaned_str = re.sub(r'```json\s*', '', raw_response)
    cleaned_str = re.sub(r'\s*```', '', cleaned_str)
    return cleaned_str.strip()

# ==========================================
# 核心函數：企業級 Excel 匯出引擎 
# ==========================================
def to_excel(df, meta_data):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    
    export_df = df.drop(columns=['選擇'])
    export_df.to_excel(writer, sheet_name='Audit Program', startrow=10, index=False, header=False)
    
    workbook = writer.book
    worksheet = writer.sheets['Audit Program']
    
    title_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter'})
    meta_format = workbook.add_format({'font_size': 11})
    header_format = workbook.add_format({'bold': True, 'bg_color': '#9BC2E6', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    cell_format = workbook.add_format({'border': 1, 'valign': 'top', 'text_wrap': True})
    
    worksheet.merge_range('A1:N1', f"{meta_data['受查單位']} Audit Program", title_format)
    worksheet.merge_range('A2:N2', f"{meta_data['專案編號']} {meta_data['循環別']}", title_format)
    
    worksheet.write('A3', f"受查單位：{meta_data['受查單位']}", meta_format)
    worksheet.write('A4', f"查核期間：{meta_data['查核期間']}", meta_format)
    worksheet.write('A5', f"查核範圍：{meta_data['查核範圍']}", meta_format)
    worksheet.write('A6', f"報告日期：{meta_data['報告日期']}", meta_format)
    worksheet.write('A7', f"稽核人員：{meta_data['稽核人員']}", meta_format)
    
    worksheet.merge_range('A9:N9', 'Audit Program', header_format)
    columns = export_df.columns.tolist()
    for col_num, value in enumerate(columns):
        worksheet.write(9, col_num, value, header_format)
        
    for row_num in range(len(export_df)):
        for col_num in range(len(columns)):
            value = export_df.iloc[row_num, col_num]
            if pd.isna(value): value = ""
            worksheet.write(row_num + 10, col_num, value, cell_format)
            
    worksheet.set_column('A:A', 10)  
    worksheet.set_column('B:B', 15)  
    worksheet.set_column('C:C', 20)  
    worksheet.set_column('D:D', 6)   
    worksheet.set_column('E:E', 40)  
    worksheet.set_column('F:F', 30)  
    worksheet.set_column('G:H', 10)  
    worksheet.set_column('I:J', 20)  
    worksheet.set_column('K:K', 30)  
    worksheet.set_column('L:L', 10)  
    worksheet.set_column('M:M', 15)  
    worksheet.set_column('N:N', 35)  

    writer.close()
    return output.getvalue()

# ==========================================
# 💾 模組：JSON 檔案專案持久化與歷史讀取機制
# ==========================================
def save_project_to_json(project_id, meta_data, dataframe):
    filename = f"Audit_Program_{project_id}.json"
    save_pack = {
        "meta": meta_data,
        "data": dataframe.to_dict(orient='records')
    }
    with open(filename, "w", encoding="utf-8") as f:
        json.dump(save_pack, f, ensure_ascii=False, indent=4)

existing_json_files = glob.glob("Audit_Program_*.json")

if existing_json_files:
    with st.expander("📂 📥 歷史專案回溯倉庫 (發現已存檔專案！)", expanded=False):
        project_options = [os.path.basename(f).replace("Audit_Program_", "").replace(".json", "") for f in existing_json_files]
        selected_project = st.selectbox("請選擇要還原的歷史稽核專案：", ["-- 請選擇 --"] + project_options)
        
        if selected_project != "-- 請選擇 --" and st.button("🔄 載入歷史專案資料與底稿狀態"):
            with open(f"Audit_Program_{selected_project}.json", "r", encoding="utf-8") as f:
                loaded_pack = json.load(f)
            st.session_state['loaded_meta'] = loaded_pack['meta']
            st.session_state['audit_program_df'] = pd.DataFrame(loaded_pack['data'])
            st.success(f"🎉 專案 【{selected_project}】 暫存資料已完美還原！")
            st.rerun()

# ==========================================
# 介面區：專案基本設定 (Meta Data)
# ==========================================
lm = st.session_state.get('loaded_meta', {})

with st.expander("📋 第一步：填寫專案與底稿表頭資訊 (必填)", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        ui_company = st.text_input("受查單位", value=lm.get("受查單位", ""), placeholder="例： A公司")
        ui_project = st.text_input("專案編號", value=lm.get("專案編號", ""), placeholder="例：2026-W10")
        ui_cycle = st.text_input("循環別", value=lm.get("循環別", ""), placeholder="例：印信管理")
    with col2:
        ui_period = st.text_input("查核期間", value=lm.get("查核期間", ""), placeholder="例：2026/06/01-2026/06/30")
        ui_scope = st.text_input("查核範圍", value=lm.get("查核範圍", ""), placeholder="例：2025/01/01-2026/04/30")
        ui_auditor = st.text_input("稽核人員", value=lm.get("稽核人員", "Andy Pan"))
        ui_date = st.text_input("報告日期", value=lm.get("報告日期", ""), placeholder="例：2026Q2")

meta_data = {
    "受查單位": ui_company, "專案編號": ui_project, "循環別": ui_cycle,
    "查核期間": ui_period, "查核範圍": ui_scope, "稽核人員": ui_auditor, "報告日期": ui_date
}

# ==========================================
# 介面區：檔案上傳與 Maker Agent 解析
# ==========================================
st.markdown("### 📂 第二步：上傳管理辦法並解析")
uploaded_files = st.file_uploader("請上傳規章 (限 PDF 格式，支援多檔)", type=['pdf'], accept_multiple_files=True)

if uploaded_files:
    st.session_state['document_text'] = extract_text_from_pdfs(uploaded_files)
else:
    st.session_state['document_text'] = "" 

if uploaded_files and st.button("🚀 啟智能解析 (產出 Audit Program)", type="primary"):
    if not api_key:
        st.warning("⚠️ 請先於左側欄位輸入 Gemini API Key！")
        st.stop()
        
    with st.spinner("🤖 Maker Agent 正在全面窮盡解析控制點... (約需 15~25 秒)"):
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-3.1-flash-lite', generation_config={"response_mime_type": "application/json"}) 
        
        system_prompt = f"""
        您是一位資深的內部稽核長。我將提供公司的管理辦法內容，請為我解析並萃取出核心查核程序。
        
        分析邏輯與要求：
        1. 極致窮盡與細化顆粒度：請逐條仔細閱讀，拆解至「最細的作業顆粒度」。
        2. 全生命週期覆蓋：確保涵蓋事前預防、事中控制、事後偵測。
        3. 降冪排序：根據「風險程度」由高至低進行排序 (高 -> 中 -> 低)。
        4. 針對每個控制點，編制稽核程序、對應辦法條文、需求清單、應注意重點、風險程度。
        
        【系統強制輸出規範】
        請務必以 JSON 陣列 (JSON Array) 格式回覆。JSON 的 key 必須嚴格命名如下：
        [
            {{"核心控制點機制": "...", "稽核程序": "...", "對應辦法條文": "...", "需求清單": "...", "應注意重點": "...", "風險程度": "..."}}
        ]
        
        以下為管理辦法內容：
        {st.session_state['document_text']}
        """
        
        try:
            response = model.generate_content(system_prompt)
            json_str = clean_json_string(response.text)
            parsed_data = json.loads(json_str)
            
            raw_df = pd.DataFrame(parsed_data)
            excel_df = pd.DataFrame(index=raw_df.index) 
            
            excel_df['選擇'] = True
            excel_df['編號'] = ""  
            excel_df['循環別'] = ui_cycle
            excel_df['風險項目'] = raw_df.get('核心控制點機制', '')
            excel_df['No'] = ""
            excel_df['Audit Step'] = raw_df.get('稽核程序', '')
            excel_df['應取得之查核文件'] = raw_df.get('需求清單', '')
            excel_df['W/P'] = ""
            excel_df['有效控制'] = ""
            excel_df['Findings'] = "" 
            excel_df['對應辦法條文'] = raw_df.get('對應辦法條文', '') 
            excel_df['應注意重點'] = raw_df.get('應注意重點', '') 
            excel_df['風險程度'] = raw_df.get('風險程度', '') 
            excel_df['AI可信度評分'] = "尚未覆核" 
            excel_df['AI覆核意見'] = "請點擊下方按鈕啟動雙重驗證"
            
            sop_row = pd.DataFrame([{
                '選擇': True, '編號': 'A', '循環別': ui_cycle, '風險項目': 'SOP程序及管理範圍', 'No': 1,
                'Audit Step': '取得最新SOP及內部控制制度，確認相關程序是否更新並達有效控制。',
                '應取得之查核文件': 'SOP/內部控制制度', 'W/P': 'A-1', '有效控制': '', 'Findings': '',
                '對應辦法條文': '', '應注意重點': '', '風險程度': '中', 'AI可信度評分': '⭐⭐⭐⭐⭐', 'AI覆核意見': '常設固定程序，豁免驗證'
            }])
            
            final_df = pd.concat([sop_row, excel_df], ignore_index=True)
            final_df = apply_smart_indexing(final_df, ui_cycle)
            
            st.session_state['audit_program_df'] = final_df
            st.success(f"🎉 第一階段解析完成！AI 已為您窮盡列出 {len(final_df)-1} 項風險點 (由高至低排序)。")
            
            if ui_project:
                save_project_to_json(ui_project, meta_data, final_df)
            
        except Exception as e:
            st.error(f"❌ 解析過程發生錯誤: {str(e)}")

# ==========================================
# 人機協作 (HITL) 編輯區
# ==========================================
if 'audit_program_df' in st.session_state:
    st.markdown("### 📝 第三步：人工勾選與微調 (Human-In-The-Loop)")
    st.info("💡 請利用最左側「選擇」勾選您本次想要查核的程序。取消勾選後點擊下方重新整理按鈕，未勾選項目將自動洗白。")
    
    col_btn1, col_btn2 = st.columns([1, 2])
    with col_btn1:
        if st.button("🔄 重新整理並計算 W/P 索引編碼"):
            st.session_state['audit_program_df'] = apply_smart_indexing(st.session_state['audit_program_df'], ui_cycle)
            st.rerun()
    with col_btn2:
        if st.button("💾 手動儲存當前進度 (同步更新至歷史倉庫)"):
            if not ui_project:
                st.warning("⚠️ 請先在最上方填寫「專案編號」，系統才知道要存成什麼檔名！")
            else:
                save_project_to_json(ui_project, meta_data, st.session_state['audit_program_df'])
                st.success(f"✅ 專案 【{ui_project}】 編輯狀態已成功儲存！")

    edited_df = st.data_editor(
        st.session_state['audit_program_df'],
        num_rows="dynamic", use_container_width=True, height=350,
        column_config={"選擇": st.column_config.CheckboxColumn("是否匯出?", default=True)}
    )
    st.session_state['audit_program_df'] = edited_df

    # ==========================================
    # 🕵️ 兩段式 Checker Agent 雙重驗證引擎
    # ==========================================
    selected_df = edited_df[edited_df['選擇'] == True].copy()
    
    st.markdown("### 🛡️ 第四步：以 AI 驗證 AI (Maker-Checker 覆核防線)")
    
    if st.button("🕵️ 啟動雙重驗證 (僅針對上方已勾選的程序)", type="secondary"):
        if 'document_text' not in st.session_state or not st.session_state['document_text']:
            st.error("❌ 找不到原始管理辦法文本文檔！請先在上方『第二步』上傳您【本次要比對的】規章 PDF。")
            st.stop()
            
        if not api_key:
            st.error("❌ 請先輸入 API Key 才能呼叫 Checker 覆核員！")
            st.stop()
            
        with st.spinner("🔍 首席品質覆核員正在嚴格核對您最新的表格內容與新規章，請稍候..."):
            
            master_df = st.session_state['audit_program_df'].copy()
            target_nos = selected_df['No'].tolist()
            
            # 💡 修復 2 同步加強：驗證前，強制確保所有「未勾選」的保持重置狀態
            for idx in master_df.index:
                if not master_df.at[idx, '選擇']:
                    master_df.at[idx, 'AI可信度評分'] = "尚未覆核"
                    master_df.at[idx, 'AI覆核意見'] = "請點擊下方按鈕啟動雙重驗證"

            # 💡 修復 1：無堅不摧的「強制轉型比對法」！先破壞，再重建。
            for t_no in target_nos:
                if pd.isna(t_no) or t_no == "": continue
                t_no_int = int(float(t_no))
                # 尋找對應列並破壞
                for idx in master_df.index:
                    val = master_df.at[idx, 'No']
                    if pd.notna(val) and val != "":
                        if int(float(val)) == t_no_int:
                            master_df.at[idx, 'AI可信度評分'] = "⚠️ 驗證失敗"
                            master_df.at[idx, 'AI覆核意見'] = "系統判定異常：本次上傳之管理辦法與查核項目內容不匹配！"
            
            st.session_state['audit_program_df'] = master_df
            
            verification_target = selected_df[['No', '風險項目', 'Audit Step', '對應辦法條文']].to_json(orient='records', force_ascii=False)
            
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-3.1-flash-lite', generation_config={"response_mime_type": "application/json"}) 
            
            checker_prompt = f"""
            您現在是一位資深犀利的「內部稽核品質覆核員 (QA)」。
            請根據【本次最新上傳的管理辦法】，嚴格審查下方的【查核程序清單】。
            
            🚨 【最高防護指令：跨領域防禦與拒絕腦補】🚨
            1. 請先判斷「管理辦法」與「查核程序清單」是否為同一業務領域。如果發現毫不相干，請你【必須】對所有項目給予 ⭐ 1星，並在意見寫明「嚴重異常：上傳之規章與本查核程序完全無關！」。
            2. 你必須在規章中找到確切的文字證據。絕對不允許腦補、假裝找到、或因為它看起來合理就給高分。
            
            本次上傳之管理辦法內容：
            {st.session_state['document_text']}
            
            使用者最新修改的查核程序清單 (JSON 格式)：
            {verification_target}
            
            【系統強制輸出規範】
            請務必回傳 JSON 陣列。JSON 的 key 必須與輸入的 No 精準對應，嚴格命名如下：
            [
                {{
                    "No": 2,
                    "AI可信度評分": "⭐⭐⭐⭐⭐",
                    "AI覆核意見": "你的具體判斷理由"
                }}
            ]
            """
            try:
                checker_res = model.generate_content(checker_prompt)
                checker_data = json.loads(checker_res.text) 
                
                # 💡 修復 1 同步：強制轉型比對法 (重建覆寫)
                for review in checker_data:
                    if not review.get('No'): continue
                    r_no_int = int(float(review.get('No')))
                    r_score = review.get('AI可信度評分', '⭐⭐⭐')
                    r_comment = review.get('AI覆核意見', '已完成比對')
                    
                    for idx in master_df.index:
                        val = master_df.at[idx, 'No']
                        if pd.notna(val) and val != "":
                            if int(float(val)) == r_no_int:
                                master_df.at[idx, 'AI可信度評分'] = r_score
                                master_df.at[idx, 'AI覆核意見'] = r_comment
                
                st.session_state['audit_program_df'] = master_df
                st.success("🎯 雙重覆核完成！已根據您本次的修改與新規章，成功覆寫驗證結果。")
                
            except Exception as e:
                st.error("❌ 覆核過程遭遇異常 (規章嚴重不匹配或 API 拒絕驗證)，已中斷覆寫。")
            
            if ui_project:
                save_project_to_json(ui_project, meta_data, st.session_state['audit_program_df'])
                st.toast("💾 專案歷程已自動同步持久化至本地端 JSON 庫！")
            st.rerun()
                
    st.divider()
    
    # ==========================================
    # 最終匯出 Excel 按鈕 
    # ==========================================
    refreshed_df = apply_smart_indexing(st.session_state['audit_program_df'].copy(), ui_cycle)
    final_download_set = refreshed_df[refreshed_df['選擇'] == True].copy()
    excel_data = to_excel(final_download_set, meta_data)
    
    st.download_button(
        label=f"📥 下載企業版 Audit Program ({len(final_download_set)}項認證程序)",
        data=excel_data,
        file_name=f"Audit Program_{ui_project}.xlsx" if ui_project else "Audit_Program_Draft.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )