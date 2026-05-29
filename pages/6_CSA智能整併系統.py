import streamlit as st
import pandas as pd
import os

# ==========================================
# 核心引擎：純淨化與智能整併 (穩定追蹤版)
# ==========================================
def sanitize_key(series):
    """強制轉為純文字、去除頭尾空白、並消除 Excel 浮點數尾數"""
    return series.astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

def smart_merge_csa(master_df, response_files, id_col, target_cols, kw_input, check_kw):
    final_df = master_df.copy()
    final_df[id_col] = sanitize_key(final_df[id_col])
    
    # 目標接收欄位強制轉為通用型態 object，避免寫入型態衝突
    for col in target_cols:
        final_df[col] = final_df[col].astype(object)
        
    source_col = "資料來源檔案"
    final_df[source_col] = "⚠️ 尚未回填"
    final_df = final_df.set_index(id_col)
    
    file_summary_list = []
    
    for f in response_files:
        try:
            rdf = pd.read_excel(f)
            if id_col not in rdf.columns:
                st.warning(f"⚠️ 檔案 {f.name} 找不到欄位『{id_col}』，已跳過。")
                continue
                
            rdf_subset = rdf[[id_col] + target_cols].copy()
            rdf_subset[id_col] = sanitize_key(rdf_subset[id_col])
            rdf_subset = rdf_subset[rdf_subset[id_col] != 'nan']
            
            for col in target_cols:
                rdf_subset[col] = rdf_subset[col].astype(object)
                
            rdf_subset = rdf_subset.set_index(id_col)
            
            # --- 運算該檔案的實質貢獻度與品質 ---
            valid_mask = pd.Series(False, index=rdf_subset.index)
            anomaly_in_file = 0
            normal_in_file = 0
            
            for col in target_cols:
                not_empty = rdf_subset[col].notna() & (rdf_subset[col].astype(str).str.strip() != '') & (rdf_subset[col].astype(str).str.lower() != 'nan')
                valid_mask = valid_mask | not_empty
                
            if valid_mask.any():
                filled_rows = rdf_subset[valid_mask]
                file_anomaly_mask = pd.Series(False, index=filled_rows.index)
                
                if check_kw and kw_input:
                    for col in target_cols:
                        series_str = filled_rows[col].astype(str).str.strip()
                        has_kw = series_str.str.contains(kw_input, na=False)
                        not_false_alarm = ~series_str.str.contains("無異常", na=False)
                        file_anomaly_mask = file_anomaly_mask | (has_kw & not_false_alarm)
                        
                anomaly_in_file = file_anomaly_mask.sum()
                normal_in_file = len(filled_rows) - anomaly_in_file
                
            clean_fname = os.path.splitext(f.name)[0]
            
            file_summary_list.append({
                "檔案名稱": clean_fname,
                "實質回填筆數": int(valid_mask.sum()),
                "✅ 正常筆數": int(normal_in_file),
                "🚩 觸發異常筆數": int(anomaly_in_file)
            })
            
            # 執行總表覆蓋更新
            updated_keys = rdf_subset[valid_mask].index
            final_df.update(rdf_subset)
            final_df.loc[final_df.index.isin(updated_keys), source_col] = clean_fname
            
        except Exception as e:
            st.error(f"❌ 讀取檔案 {f.name} 時發生錯誤: {str(e)}")
            
    return final_df.reset_index(), file_summary_list, source_col

# ==========================================
# 網頁 UI 設計
# ==========================================
st.set_page_config(page_title="CSA 智能整併系統", layout="wide")
st.title("📊 內控自評 (CSA) 多檔智能整併與檢核系統")
st.write("透過指定『唯一 KEY 值』，全自動將數十個部門回填檔的指定欄位精準彙整至總表，並具備自訂異常篩選引擎。")

# --- 步驟 1 ---
st.subheader("步驟 1：載入總表骨架與欄位對應")
master_file = st.file_uploader("📥 請上傳『原始乾淨總表範本』", type=['xlsx'])

if master_file:
    master_df = pd.read_excel(master_file)
    all_cols = master_df.columns.tolist()
    
    col_k, col_t = st.columns([1, 2])
    with col_k:
        id_col = st.selectbox("🔑 選擇『唯一識別碼 (KEY)』欄位：", all_cols)
    with col_t:
        # 💡 介面修改為單選 selectbox，預設自動選擇最後一個欄位 (通常是說明欄)
        target_col = st.selectbox("🎯 選擇『需從分表抓回覆蓋的欄位』：", all_cols, index=len(all_cols)-1 if len(all_cols)>0 else 0)
        # 背景自動包裝成單一元素的清單，讓核心覆寫引擎完美運作
        target_cols = [target_col]

    st.divider()

    # --- 步驟 2 ---
    st.subheader("步驟 2：上傳各部門回填檔與設定檢核邏輯")
    col_file, col_rule = st.columns([3, 2])
    
    with col_file:
        response_files = st.file_uploader("📂 拖曳所有收回的 Excel 分表至此 (可一次全選)", type=['xlsx'], accept_multiple_files=True)
        
    with col_rule:
        st.info("⚙️ **異常檢核觸發條件設定**")
        check_null = st.checkbox("🔍 偵測『未填寫 (空值)』", value=True, help="目標欄位為完全空白或無資料時觸發")
        check_kw = st.checkbox("關鍵字偵測 (自動排除 '無異常')", value=True)
        kw_input = st.text_input("自訂異常關鍵字 (多個用 | 分開)：", "異常|不符合|待改善|延遲")

    if response_files and target_cols:
        if st.button("🚀 啟動智能彙整與檢核", type="primary"):
            with st.spinner("資料對齊與防呆引擎運作中..."):
                result_df, file_summary_list, source_col = smart_merge_csa(master_df, response_files, id_col, target_cols, kw_input, check_kw)
                
                # --- 總體異常運算 ---
                anomaly_mask = pd.Series(False, index=result_df.index)
                for col in target_cols:
                    if check_null:
                        is_empty = result_df[col].isna() | (result_df[col].astype(str).str.strip() == '') | (result_df[col].astype(str).str.lower() == 'nan')
                        anomaly_mask = anomaly_mask | is_empty
                        
                    if check_kw and kw_input:
                        series_str = result_df[col].astype(str).str.strip()
                        has_kw = series_str.str.contains(kw_input, na=False)
                        not_false_alarm = ~series_str.str.contains("無異常", na=False)
                        anomaly_mask = anomaly_mask | (has_kw & not_false_alarm)

                anomaly_df = result_df[anomaly_mask]
                missing_df = result_df[result_df[source_col] == "⚠️ 尚未回填"]
                
                # --- 步驟 3：儀表板呈報 ---
                st.divider()
                st.success(f"🎊 整併完成！本次共成功解析了 {len(file_summary_list)} 份檔案。")
                
                m1, m2, m3 = st.columns(3)
                m1.metric("📌 總查核項目數", len(result_df))
                m2.metric("⚠️ 觸發檢核項目數 (空值/異常)", len(anomaly_df))
                
                filled_count = len(result_df) - len(missing_df)
                fill_rate = round((filled_count / len(result_df)) * 100, 1)
                m3.metric("📈 整體資料回填率", f"{fill_rate}%")

                st.markdown("---")
                st.subheader("📁 各部門檔案回填品質追蹤")
                summary_df = pd.DataFrame(file_summary_list)
                summary_df = summary_df.rename(columns={"有效填答數": "實質回填筆數", "正常數": "✅ 正常筆數", "異常數": "🚩 觸發異常筆數"})
                st.dataframe(summary_df, use_container_width=True, hide_index=True)

                st.markdown("---")
                
                tab1, tab2, tab3 = st.tabs(["🚩 異常與待確認清單", f"⏳ 尚未回填追蹤 ({len(missing_df)} 筆)", "📑 完整彙整總表預覽"])
                
                with tab1:
                    st.caption(f"📌 本頁面過濾出共 **{len(anomaly_df)}** 筆觸發空值或異常關鍵字的控制項：")
                    if not anomaly_df.empty:
                        st.dataframe(anomaly_df, use_container_width=True, hide_index=True)
                    else:
                        st.success("🎉 完美！所有分表皆已完整回填，且未觸發任何異常關鍵字。")
                        
                with tab2:
                    st.caption("📌 以下項目在所有上傳的分表中均無對應資料更新，請確認對應的【權責單位】是否漏交：")
                    if not missing_df.empty:
                        st.dataframe(missing_df, use_container_width=True, hide_index=True)
                    else:
                        st.success("🎉 太棒了！所有查核項目均已成功匹配到來源檔案。")
                        
                with tab3:
                    st.dataframe(result_df, use_container_width=True, hide_index=True)

                # --- 步驟 4：下載 ---
                st.divider()
                out_file = "CSA_Consolidated_Master.xlsx"
                result_df.to_excel(out_file, index=False)
                with open(out_file, "rb") as final_excel:
                    st.download_button("📥 下載整併完成總表 (.xlsx)", final_excel, file_name="內控自評彙整底稿.xlsx", type="primary")

else:
    st.info("💡 操作提示：請先上傳原始乾淨總表，系統將自動啟動後續分析模組。")