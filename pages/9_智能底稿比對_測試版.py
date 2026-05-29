import streamlit as st
import google.generativeai as genai
import PyPDF2
import pandas as pd
import json
import re
import io
import os
import glob
import tempfile
from PIL import Image, ImageDraw

# 嘗試載入 PyMuPDF (fitz) 用於 PDF 渲染
try:
    import fitz 
    HAS_FITZ = True
except ImportError:
    HAS_FITZ = False

# ==========================================
# 網頁基礎配置
# ==========================================
st.set_page_config(page_title="智能底稿比對 Agent", page_icon="🕵️", layout="wide")
st.title("🕵️ 智能底稿比對 Agent (人機協作完稿版)")
st.write("全面支援 PDF、圖檔、Word、Excel。內建「分頁導航預覽器」，完美結合 AI 語意分析與稽核專業判斷。")

# ==========================================
# 側邊欄：API 設定
# ==========================================
with st.sidebar:
    st.header("🔑 系統連線設定")
    api_key = st.text_input("請輸入 Google Gemini API Key:", type="password")
    st.info("💡 系統指定使用 Gemini 3.1 Flash Lite 引擎，發揮高性價比多模態能力。")
    if not HAS_FITZ:
        st.error("⚠️ 缺少 PyMuPDF 套件，將無法預覽 PDF。請於終端機執行 `pip install PyMuPDF`。")

# ==========================================
# 核心函數工具箱
# ==========================================
def clean_json_string(raw_response):
    cleaned_str = re.sub(r'```json\s*', '', raw_response)
    cleaned_str = re.sub(r'\s*```', '', cleaned_str)
    return cleaned_str.strip()

def save_project_to_json(project_id, meta_data, dataframe, mapping_res, page_res):
    filename = f"Audit_Program_{project_id}.json"
    save_pack = {
        "meta": meta_data,
        "data": dataframe.to_dict(orient='records'),
        "mapping": mapping_res,
        "pages": page_res
    }
    with open(filename, "w", encoding="utf-8") as f:
        json.dump(save_pack, f, ensure_ascii=False, indent=4)

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
    
    worksheet.merge_range('A1:N1', f"{meta_data.get('受查單位', '')} Audit Program", title_format)
    worksheet.merge_range('A2:N2', f"{meta_data.get('專案編號', '')} {meta_data.get('循環別', '')} - 現場查核結果", title_format)
    worksheet.write('A3', f"受查單位：{meta_data.get('受查單位', '')}", meta_format)
    worksheet.write('A4', f"查核期間：{meta_data.get('查核期間', '')}", meta_format)
    worksheet.write('A5', f"查核範圍：{meta_data.get('查核範圍', '')}", meta_format)
    worksheet.write('A6', f"報告日期：{meta_data.get('報告日期', '')}", meta_format)
    worksheet.write('A7', f"稽核人員：{meta_data.get('稽核人員', '')}", meta_format)
    
    worksheet.merge_range('A9:N9', 'Audit Program (Fieldwork)', header_format)
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
    worksheet.set_column('G:G', 10)  
    worksheet.set_column('H:H', 12)  
    worksheet.set_column('I:I', 55)  
    worksheet.set_column('J:J', 20)  
    worksheet.set_column('K:K', 30)  
    worksheet.set_column('L:L', 10)  
    worksheet.set_column('M:M', 15)  
    worksheet.set_column('N:N', 30)  
    writer.close()
    return output.getvalue()

def render_pdf_page(pdf_bytes, page_num):
    if not HAS_FITZ:
        return None, 0
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        total_pages = len(doc)
        target_page_idx = max(0, min(page_num - 1, total_pages - 1))
        page = doc[target_page_idx]
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        return pix.tobytes("png"), total_pages
    except Exception:
        return None, 0

# ==========================================
# 📂 第一步：載入已規劃之專案 JSON
# ==========================================
st.markdown("### 📂 第一步：載入規劃階段專案 (Loading Project)")
existing_json_files = glob.glob("Audit_Program_*.json")

if not existing_json_files:
    st.warning("⚠️ 在本地目錄中找不到任何專案存檔！請先至「第 8 支程式」完成規劃。")
    st.stop()

project_options = [os.path.basename(f).replace("Audit_Program_", "").replace(".json", "") for f in existing_json_files]
selected_project = st.selectbox("請選擇要執行查核的專案：", ["-- 請選擇 --"] + project_options)

if selected_project != "-- 請選擇 --":
    if st.button("📥 載入此專案底稿"):
        st.session_state['evidence_dict'] = {}
        st.session_state['processed_meta'] = []
        
        with open(f"Audit_Program_{selected_project}.json", "r", encoding="utf-8") as f:
            loaded_pack = json.load(f)
        st.session_state['fw_meta'] = loaded_pack['meta']
        st.session_state['fw_df'] = pd.DataFrame(loaded_pack['data'])
        st.session_state['mapping_results'] = loaded_pack.get('mapping', {})
        st.session_state['page_results'] = loaded_pack.get('pages', {})
        st.success(f"🎉 專案 【{selected_project}】 載入成功！")
        st.rerun()

if 'fw_df' not in st.session_state:
    st.stop()

meta = st.session_state['fw_meta']
df = st.session_state['fw_df']

with st.expander(f"📋 專案資訊確認：{meta.get('專案編號', '')} {meta.get('循環別', '')}", expanded=True):
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown(f"**🏢 受查單位:** {meta.get('受查單位', '')}")
        st.markdown(f"**🏷️ 專案編號:** {meta.get('專案編號', '')}")
    with col2:
        st.markdown(f"**📅 查核期間:** {meta.get('查核期間', '')}")
        st.markdown(f"**🎯 查核範圍:** {meta.get('查核範圍', '')}")
    with col3:
        st.markdown(f"**👤 稽核人員:** {meta.get('稽核人員', 'Andy Pan')}")
        st.markdown(f"**📆 報告日期:** {meta.get('報告日期', '')}")

# ==========================================
# 📦 第二步：多元素材庫與視覺引擎 
# ==========================================
st.markdown("### 📦 第二步：上傳受查單位佐證資料 (多模態素材庫)")
st.info("💡 支援 PDF、圖片、Word、Excel。我們將啟用 Gemini 原生視覺引擎直接判讀整份文件！")
uploaded_evidence = st.file_uploader(
    "上傳佐證附件 (支援 .pdf, .jpg, .png, .docx, .xlsx)", 
    type=['pdf', 'jpg', 'jpeg', 'png', 'docx', 'xlsx'], 
    accept_multiple_files=True
)

if uploaded_evidence:
    if not api_key:
        st.error("⚠️ 請先於左側輸入 API Key！")
        st.stop()
        
    current_files_meta = [(f.name, f.size) for f in uploaded_evidence]
    
    if 'evidence_dict' not in st.session_state or st.session_state.get('processed_meta') != current_files_meta:
        with st.spinner("👁️ 正在啟動 Gemini 視覺解析引擎處理您的佐證資料，請稍候..."):
            genai.configure(api_key=api_key)
            evidence_dict = {}
            temp_dir = tempfile.mkdtemp()
            
            for file in uploaded_evidence:
                ext = file.name.split('.')[-1].lower()
                try:
                    if ext in ['pdf', 'jpg', 'jpeg', 'png']:
                        file_bytes = file.read()
                        temp_path = os.path.join(temp_dir, file.name)
                        with open(temp_path, "wb") as f:
                            f.write(file_bytes)
                        uploaded_gfile = genai.upload_file(temp_path)
                        evidence_dict[file.name] = {"type": "vision_file", "ext": ext, "data": uploaded_gfile, "bytes": file_bytes}
                    
                    elif ext == 'docx':
                        import docx
                        doc = docx.Document(io.BytesIO(file.read()))
                        text = "\n".join([para.text for para in doc.paragraphs])
                        evidence_dict[file.name] = {"type": "text", "ext": ext, "data": f"--- {file.name} 內容 ---\n{text}"}
                    
                    elif ext in ['xlsx', 'xls']:
                        df_dict = pd.read_excel(io.BytesIO(file.read()), sheet_name=None)
                        text = f"--- {file.name} 內容 ---\n"
                        for sheet_name, dframe in df_dict.items():
                            text += f"【工作表: {sheet_name}】\n{dframe.to_string(index=False)}\n\n"
                        evidence_dict[file.name] = {"type": "text", "ext": ext, "data": text}
                        
                except Exception as e:
                    st.error(f"❌ 讀取 {file.name} 失敗: {str(e)}")
                    
            st.session_state['evidence_dict'] = evidence_dict
            st.session_state['processed_meta'] = current_files_meta
            st.success(f"✅ 成功載入並解析 {len(uploaded_evidence)} 份佐證資料！")
else:
    if 'evidence_dict' not in st.session_state:
        st.session_state['evidence_dict'] = {}
    st.session_state['processed_meta'] = []

# ==========================================
# 🔗 第三步：查核程序配對 (Mapping UI)
# ==========================================
st.markdown("### 🔗 第三步：智能底稿配對 (Mapping)")
active_df = df[df['選擇'] == True].copy()
mapping_results = st.session_state.get('mapping_results', {})

for index, row in active_df.iterrows():
    r_no = str(row['No']) 
    wp_code = row.get('W/P', '無編碼')
    risk_item = row.get('風險項目', '無風險項目')
    audit_step = row.get('Audit Step', '')
    req_docs = row.get('應取得之查核文件', '')
    
    with st.container(border=True):
        st.markdown(f"**📌 索引 [{wp_code}] - {risk_item}**")
        st.caption(f"**查核程序：** {audit_step}")
        st.caption(f"**應取得文件：** {req_docs}")
        
        if r_no not in mapping_results:
            mapping_results[r_no] = []
            
        with st.expander("📂 點擊展開清單，勾選欲分配的佐證資料 (可多選)"):
            evidence_keys = list(st.session_state.get('evidence_dict', {}).keys())
            if not evidence_keys:
                st.info("尚無可用的佐證資料，請先於第二步上傳檔案。")
            else:
                for fname in evidence_keys:
                    is_checked = fname in mapping_results[r_no]
                    if st.checkbox(fname, value=is_checked, key=f"chk_{r_no}_{fname}"):
                        if fname not in mapping_results[r_no]: mapping_results[r_no].append(fname)
                    else:
                        if fname in mapping_results[r_no]: mapping_results[r_no].remove(fname)

st.session_state['mapping_results'] = mapping_results

# ==========================================
# 🕵️ 第四步：啟動 AI 智能比對 (強制寫入進化版)
# ==========================================
st.markdown("### 🕵️ 第四步：AI 視覺自動比對與產出 Findings")

if st.button("🚀 啟動多模態智能比對 (生成 Fieldwork 底稿)", use_container_width=True, type="primary"):
    if not api_key:
        st.error("❌ 請先於左側欄位輸入 API Key！")
        st.stop()
        
    master_df = st.session_state['fw_df'].copy()
    page_results = {} 
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    total_steps = len(active_df)
    current_step = 0
    
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-3.1-flash-lite', generation_config={"response_mime_type": "application/json"})
    
    for index, row in active_df.iterrows():
        r_no = str(row['No'])
        wp_code = row.get('W/P', f"No.{r_no}")
        raw_mapped_filenames = mapping_results.get(r_no, [])
        mapped_filenames = [f for f in raw_mapped_filenames if f in st.session_state.get('evidence_dict', {})]
        
        current_step += 1
        status_text.markdown(f"🔍 **首席稽核機器人比對中 ({current_step}/{total_steps})**：正在解析底稿索引 `[{wp_code}]`...")
        progress_bar.progress(current_step / total_steps)
        
        # 💡 終極防護：改用 .loc[index, ...] 以資料表的「絕對座標」強制覆寫，徹底解決型別跑版寫入失敗的問題！
        if not mapped_filenames:
            master_df.loc[index, '有效控制'] = "未檢測"
            master_df.loc[index, 'Findings'] = "【結論】：受查單位未提供相關佐證資料（或尚未上傳配對），無法執行比對。\n【發現事實】：無。\n【違反之管理辦法條文】：無。\n【可能造成的風險影響】：無。"
            continue
            
        master_df.loc[index, '有效控制'] = "⚠️ 異常"
        master_df.loc[index, 'Findings'] = "系統比對過程發生錯誤或中斷。"
        
        content_payload = []
        
        compare_prompt_text = f"""
        您是一位極度嚴謹的「主辦內部稽核 (Lead Auditor)」。
        請嚴格審查我附上的【佐證資料】並與下方【查核程序標準】比對。

        🚨 【稽核查核順序與判定防呆原則】(請務必依序檢驗)
        第一關：【關聯性審查 (Domain Match)】
        - 檢查「佐證資料的本質」是否符合「Audit Step 的要求」。
        - 絕對禁止「牛頭不對馬嘴」！若查核程序在查「採購與供應商」，但佐證卻是「印信管理辦法」等完全無關之文件，請立刻判定為「失效」，並於結論與事實中直接寫明：「提供之佐證資料性質與查核程序完全無關」，終止後續檢查。

        第二關：【時間邊界審查】
        - 若第一關通過，請檢查佐證資料日期是否落於法定查核範圍：【{meta.get('查核範圍', '未指定')}】。
        - 若不在範圍內，請判定為「失效」，指出發生日期不符查核區間。

        第三關：【實質控制點審查】
        - 若前兩關皆通過，才仔細檢驗文件中是否有滿足 Audit Step 要求的簽名、權限與防弊控制點。

        🚨 【PDF 定位要求】
        若資料包含多頁 PDF，且您判定為「失效」或「部分有效」，請指出導致您判定異常的證據發生在該檔案的「第幾頁」。若為整份文件皆不合規（如牛頭不對馬嘴），請填 1。

        【查核程序標準】
        - 應執行的 Audit Step：{row.get('Audit Step', '')}
        - 應注意重點：{row.get('應注意重點', '')}
        - 對應條文：{row.get('對應辦法條文', '')}

        【系統強制 JSON 輸出規範】
        {{
            "有效控制": "OK" 或 "部分有效" 或 "失效",
            "結論": "直接呈述該查核程序是否遵循或違反規定。",
            "發現事實": "若文件無關請直接點破。若有關則具體描述您從資料中看到的事實。",
            "違反之管理辦法條文": "若有違反請指明條文，否則填寫『無』。",
            "可能造成的風險影響": "根據缺失推導風險，若為無關文件請說明『無法驗證該控制點之有效性』。",
            "異常頁數": 1
        }}
        """
        
        content_payload.append(compare_prompt_text)
        
        for fname in mapped_filenames:
            ev_obj = st.session_state['evidence_dict'][fname]
            if ev_obj['type'] == 'vision_file':
                content_payload.append(ev_obj['data']) 
            else:
                content_payload.append(ev_obj['data']) 
        
        try:
            res = model.generate_content(content_payload)
            res_data = json.loads(res.text)
            if isinstance(res_data, list):
                res_data = res_data[0] if len(res_data) > 0 else {}
            
            final_findings_str = (
                f"【結論】：{res_data.get('結論', '')} ({res_data.get('有效控制', '')})\n"
                f"【發現事實】：{res_data.get('發現事實', '')}\n"
                f"【違反之管理辦法條文】：{res_data.get('違反之管理辦法條文', '')}\n"
                f"【可能造成的風險影響】：{res_data.get('可能造成的風險影響', '')}\n"
                f"--- 📎 參照附件：{', '.join(mapped_filenames)} ---"
            )
            
            # 強制寫入絕對座標，絕不跑版
            master_df.loc[index, '有效控制'] = res_data.get('有效控制', '異常')
            master_df.loc[index, 'Findings'] = final_findings_str
            page_results[r_no] = max(1, int(res_data.get('異常頁數', 1)))
            
        except Exception as e:
            # 萬一遇到 API 異常，現在錯誤訊息也能精準寫進表格中，不會再被吃掉了！
            master_df.loc[index, 'Findings'] = f"系統比對過程發生錯誤。(錯誤細節: {str(e)})"
            continue
            
    status_text.markdown("✅ **比對任務全數完成！**")
    st.session_state['fw_df'] = master_df
    st.session_state['page_results'] = page_results
    
    save_project_to_json(selected_project, meta, master_df, mapping_results, page_results)
    st.success("🎯 全案多模態比對完成！")
    st.rerun()

st.divider()

# ==========================================
# 📊 第五步：垂直佈局大面板 (終極人機協作)
# ==========================================
if 'fw_df' in st.session_state:
    st.markdown("### 📊 第五步：人工檢閱與匯出完稿 (人機協作現場)")
    
    display_df = st.session_state['fw_df'][st.session_state['fw_df']['選擇'] == True].copy()
    display_df = display_df[display_df['Audit Step'].astype(str).str.strip() != ""]
    
    # 【上半部】：全寬度表格
    edited_display_df = st.data_editor(
        display_df,
        hide_index=True, 
        num_rows="dynamic", 
        use_container_width=True, 
        height=320,
        column_config={"選擇": st.column_config.CheckboxColumn("是否匯出?", default=True)}
    )
    
    col_b1, col_b2 = st.columns(2)
    with col_b1:
        if st.button("💾 儲存底稿修改", key="save_fw", use_container_width=True):
            master_df = st.session_state['fw_df'].copy()
            master_df.update(edited_display_df)
            st.session_state['fw_df'] = master_df
            save_project_to_json(selected_project, st.session_state['fw_meta'], master_df, st.session_state['mapping_results'], st.session_state.get('page_results', {}))
            st.toast("✅ 手動修改已成功寫入專案庫！")

    with col_b2:
        final_download_set = edited_display_df[edited_display_df['選擇'] == True].copy()
        excel_data = to_excel(final_download_set, st.session_state['fw_meta'])
        st.download_button(
            label=f"📥 下載 Fieldwork 查核底稿完稿版 (Excel)",
            data=excel_data,
            file_name=f"Audit Program_{selected_project}_AI Review.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

    st.markdown("---")
    
    # 【下半部】：完美左資右圖佈局 (4:6)
    st.markdown("#### 👁️ 數位稽核佐證調閱面板 (多頁導航版)")
    
    col_info, col_viewer = st.columns([4, 6], gap="large")
    
    wp_options = edited_display_df['W/P'].tolist()
    
    if wp_options:
        with col_info:
            selected_wp = st.selectbox("🔍 請挑選欲調閱影像的程序索引 (W/P)：", options=wp_options)
            
            selected_row_data = edited_display_df[edited_display_df['W/P'] == selected_wp].iloc[0]
            selected_no = str(selected_row_data['No'])
            control_status = selected_row_data['有效控制']
            findings_text = selected_row_data['Findings']
            
            if control_status in ["失效", "部分有效"]:
                st.error(f"⚠️ 稽核結論：【{control_status}】")
            elif control_status == "OK":
                st.success(f"✅ 稽核結論：【{control_status}】")
            elif control_status == "未檢測":
                st.warning(f"💡 稽核結論：【{control_status}】 (尚未配對上傳之佐證檔案)")
            else:
                st.warning(f"💡 稽核結論：【{control_status}】")
                
            st.markdown("##### 📝 系統查核紀錄")
            st.info(findings_text.replace('\n', '\n\n'))

        with col_viewer:
            raw_mapped_files = st.session_state.get('mapping_results', {}).get(selected_no, [])
            mapped_files = [f for f in raw_mapped_files if f in st.session_state.get('evidence_dict', {})]
            
            if not mapped_files:
                if raw_mapped_files:
                    st.warning("⚠️ 查核程序曾配對過檔案，但您目前尚未將檔案上傳至第二步的素材庫。")
                else:
                    st.info("該查核程序未配對任何佐證文件。")
            else:
                tabs = st.tabs(mapped_files)
                for i, fname in enumerate(mapped_files):
                    with tabs[i]:
                        ev_obj = st.session_state.get('evidence_dict', {}).get(fname)
                        if not ev_obj:
                            continue
                            
                        target_page = st.session_state.get('page_results', {}).get(selected_no, 1)
                        
                        if ev_obj['ext'] == 'pdf':
                            if HAS_FITZ:
                                doc_temp = fitz.open(stream=ev_obj['bytes'], filetype="pdf")
                                max_pages = len(doc_temp)
                                
                                if max_pages > 1:
                                    safe_target = max(1, min(target_page, max_pages))
                                    user_selected_page = st.slider(f"📄 瀏覽 {fname} 的頁碼", min_value=1, max_value=max_pages, value=safe_target, key=f"slider_{selected_no}_{fname}")
                                else:
                                    user_selected_page = 1
                                
                                rendered_pdf_img, _ = render_pdf_page(ev_obj['bytes'], user_selected_page)
                                
                                if rendered_pdf_img:
                                    st.image(rendered_pdf_img, use_container_width=True)
                            else:
                                st.warning("請安裝 PyMuPDF 套件以啟用 PDF 渲染。")
                                
                        elif ev_obj['ext'] in ['jpg', 'jpeg', 'png']:
                            st.image(ev_obj['bytes'], use_container_width=True)
                        else:
                            st.info("此檔案格式暫不支援影像渲染，請參閱左側文字紀錄。")
