import streamlit as st
import os
import zipfile
import subprocess
import shutil
import io
import time
from bs4 import BeautifulSoup
from PIL import Image

def convert_doc_to_docx(doc_path, output_dir):
    try:
        command = ['libreoffice', '--headless', '--convert-to', 'docx', doc_path, '--outdir', output_dir]
        subprocess.run(command, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        base_name = os.path.splitext(os.path.basename(doc_path))[0]
        return os.path.join(output_dir, f"{base_name}.docx")
    except:
        return None

def extract_images_in_order(docx_path, output_folder, original_name=None, log_container=None, max_name_len=15):
    original_base_name = original_name if original_name else os.path.splitext(os.path.basename(docx_path))[0]
    base_name = original_base_name[:max_name_len] + "_短" if len(original_base_name) > max_name_len else original_base_name
    extracted_count = 0

    try:
        with zipfile.ZipFile(docx_path, 'r') as docx_zip:
            rels_xml = docx_zip.read('word/_rels/document.xml.rels')
            rels_soup = BeautifulSoup(rels_xml, 'xml')
            rel_map = {rel.get('Id'): rel.get('Target') for rel in rels_soup.find_all('Relationship')}
            doc_xml = docx_zip.read('word/document.xml')
            doc_soup = BeautifulSoup(doc_xml, 'xml')
            
            image_rids = []
            for tag in doc_soup.find_all(['blip', 'imagedata']):
                rid = tag.get('r:embed') or tag.get('r:id')
                if rid and rid in rel_map and rel_map[rid].startswith('media/'):
                    image_rids.append(rid)

            count = 1
            for rid in image_rids:
                target_path = f"word/{rel_map[rid]}"
                if target_path in docx_zip.namelist():
                    image_data = docx_zip.read(target_path)
                    new_name = f"{base_name}_{count:02d}.png"
                    out_path = os.path.join(output_folder, new_name)
                    
                    try:
                        image = Image.open(io.BytesIO(image_data))
                        if image.mode not in ('RGB', 'RGBA'): image = image.convert('RGBA')
                        image.save(out_path, format="PNG")
                        if log_container: log_container.text(f"✅ 成功擷取: {new_name}")
                        extracted_count += 1
                    except:
                        ext = target_path.split('.')[-1]
                        with open(os.path.join(output_folder, f"{base_name}_{count:02d}.{ext}"), 'wb') as f:
                            f.write(image_data)
                        extracted_count += 1
                    count += 1
    except Exception as e:
        if log_container: log_container.error(f"❌ 錯誤: {str(e)}")
        
    return extracted_count

st.title("📝 Word 圖片提取 (防鎖定版)")
st.write("請上傳需要提取圖片的 Word 檔案。系統將自動按頁面順序擷取原始圖片，並打包成 ZIP 壓縮檔供您下載。")

files = st.file_uploader("📂 拖曳或點擊此處上傳 Word 檔案 (.doc, .docx)", type=['doc', 'docx'], accept_multiple_files=True)

if files and st.button("🚀 開始提取"):
    task_id = time.strftime("%Y%m%d_%H%M%S")
    WORK_DIR = f"temp_word_{task_id}"
    os.makedirs(WORK_DIR, exist_ok=True)
    log = st.container()
    
    total_images_extracted = 0

    for f in files:
        file_path = os.path.join(WORK_DIR, f.name)
        with open(file_path, "wb") as tmp: tmp.write(f.getbuffer())
        
        if f.name.endswith('.doc'):
            docx_p = convert_doc_to_docx(file_path, WORK_DIR)
            if docx_p: 
                total_images_extracted += extract_images_in_order(docx_p, WORK_DIR, original_name=os.path.splitext(f.name)[0], log_container=log)
        else:
            total_images_extracted += extract_images_in_order(file_path, WORK_DIR, log_container=log)

    st.success(f"🎉 處理完成！本次共成功提取了 **{total_images_extracted}** 張圖片。")

    zip_file = f"Word_Result_{task_id}"
    shutil.make_archive(zip_file, 'zip', WORK_DIR)
    
    with open(f"{zip_file}.zip", "rb") as z:
        st.download_button("📥 下載提取結果 (ZIP)", z, file_name="Word圖片提取結果.zip", type="primary")
    
    shutil.rmtree(WORK_DIR, ignore_errors=True)