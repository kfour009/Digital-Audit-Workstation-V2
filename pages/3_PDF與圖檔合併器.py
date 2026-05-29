import streamlit as st
import os
import shutil
import fitz
import time
import pandas as pd

st.title("📑 PDF 與圖檔合併器 (防鎖定版)")
st.write("將零散的圖片與 PDF 文件合併為單一底稿。您可以直接在下方表格手動修改順序。")

files = st.file_uploader("📂 請一次選取所有要合併的檔案", type=['pdf', 'png', 'jpg', 'jpeg'], accept_multiple_files=True)

if files:
    # 建立一個供使用者編輯排序的 DataFrame
    file_names = [f.name for f in files]
    df = pd.DataFrame({
        "順序 (點擊修改)": range(1, len(file_names) + 1),
        "檔名": file_names
    })

    st.info("📌 **手動調整合併順序：**\n請直接點擊下方表格的『順序』欄位修改數字。合併時系統會自動依數字由小到大進行裝訂。")
    
    # 顯示可編輯表格 (唯讀檔名，只能改順序)
    edited_df = st.data_editor(df, hide_index=True, use_container_width=True, disabled=["檔名"])
    
    # 依據使用者輸入的數字重新排序
    sorted_df = edited_df.sort_values(by="順序 (點擊修改)")
    ordered_names = sorted_df["檔名"].tolist()
    
    # 將排序後的檔名對應回真實的檔案物件
    file_dict = {f.name: f for f in files}
    sorted_files = [file_dict[name] for name in ordered_names]

    st.divider()
    out_name = st.text_input("✏️ 請設定輸出檔名：", "最終合併底稿.pdf")
    if not out_name.endswith('.pdf'):
        out_name += ".pdf"
    
    if st.button("🔗 執行合併", type="primary"):
        task_id = time.strftime("%Y%m%d_%H%M%S")
        WORK_DIR = f"temp_merge_{task_id}"
        os.makedirs(WORK_DIR, exist_ok=True)
        
        merged_pdf = fitz.open()
        
        for f in sorted_files:
            f_bytes = f.read()
            if f.name.lower().endswith('.pdf'):
                doc = fitz.open("pdf", f_bytes)
                merged_pdf.insert_pdf(doc)
                doc.close()
            else:
                img = fitz.open("img", f_bytes)
                pdf_bytes = img.convert_to_pdf()
                img_pdf = fitz.open("pdf", pdf_bytes)
                merged_pdf.insert_pdf(img_pdf)
                img_pdf.close()
                img.close()
        
        out_path = os.path.join(WORK_DIR, out_name)
        merged_pdf.save(out_path)
        merged_pdf.close()
        
        st.success(f"🎉 裝訂成功！檔案已依照您的自訂順序合併。")
        
        with open(out_path, "rb") as final:
            st.download_button(f"📥 下載 {out_name}", final, file_name=out_name, type="primary")
            
        shutil.rmtree(WORK_DIR, ignore_errors=True)