import streamlit as st
import os
import shutil
import fitz  # PyMuPDF
import time

def extract_images_from_pdf(pdf_path, output_folder, log_container):
    try:
        doc = fitz.open(pdf_path)
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        total = 0

        for page_num in range(len(doc)):
            page = doc[page_num]
            image_list = page.get_images(full=True)
            
            for img_idx, img in enumerate(image_list):
                xref = img[0]
                pix = fitz.Pixmap(doc, xref)
                
                # 過濾過小的碎片
                if pix.width < 50 or pix.height < 50: continue
                
                if pix.n - pix.alpha >= 4: pix = fitz.Pixmap(fitz.csRGB, pix)
                
                name = f"{base_name}_p{page_num+1}_img{img_idx+1}.png"
                pix.save(os.path.join(output_folder, name))
                log_container.text(f"✅ 成功擷取: {name}")
                total += 1
        
        doc.close()
        return total
    except Exception as e:
        log_container.error(f"❌ 錯誤: {str(e)}")
        return 0

st.title("📄 PDF 圖片提取 (防鎖定版)")
st.write("請上傳 PDF 檔案，系統將自動提取出內含的圖片（會自動過濾微小排版碎片）。")

files = st.file_uploader("📂 拖曳或點擊此處上傳 PDF 檔案 (.pdf)", type=['pdf'], accept_multiple_files=True)

if files and st.button("🚀 開始處理"):
    task_id = time.strftime("%Y%m%d_%H%M%S")
    WORK_DIR = f"temp_pdf_{task_id}"
    os.makedirs(WORK_DIR, exist_ok=True)
    log = st.container()

    total_images_extracted = 0

    for f in files:
        path = os.path.join(WORK_DIR, f.name)
        with open(path, "wb") as tmp: tmp.write(f.getbuffer())
        total_images_extracted += extract_images_from_pdf(path, WORK_DIR, log)

    st.success(f"🎉 處理完成！本次共成功提取了 **{total_images_extracted}** 張圖片。")

    zip_name = f"PDF_Result_{task_id}"
    shutil.make_archive(zip_name, 'zip', WORK_DIR)
    
    with open(f"{zip_name}.zip", "rb") as z:
        st.download_button("📥 下載提取結果 (ZIP)", z, file_name="PDF圖片提取結果.zip", type="primary")
    
    shutil.rmtree(WORK_DIR, ignore_errors=True)