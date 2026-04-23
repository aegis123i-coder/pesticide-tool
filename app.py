import streamlit as st
import pdfplumber
import pandas as pd
import io

# ================= 1. 介面基礎設定 =================
st.set_page_config(page_title="PDF 轉 Excel 神器", page_icon="📄", layout="centered")

# 標題與說明
st.title("📄 PDF 農藥容許量自動萃取工具")
st.markdown("這是一個自動化工具。請上傳您的 PDF 檔案，系統會自動為您萃取表格、清洗資料，並轉換為乾淨的 Excel 格式。")

# ================= 2. 檔案上傳區塊 =================
uploaded_file = st.file_uploader("請將 PDF 檔案拖曳至此，或點擊選擇檔案", type="pdf")

# 當使用者上傳檔案後才顯示處理按鈕
if uploaded_file is not None:
    st.info(f"已讀取檔案：{uploaded_file.name}")
    
    # 點擊按鈕開始處理
    if st.button("🚀 開始萃取資料", type="primary", use_container_width=True):
        
        # ================= 3. UI 元件：進度條與狀態文字 =================
        progress_bar = st.progress(0) # 初始化進度條 (0%)
        status_text = st.empty()      # 建立一個文字佔位符，用來隨時更新狀態
        
        all_table_data = []
        
        try:
            # ================= 4. 讀取 PDF 與進度更新 =================
            with pdfplumber.open(uploaded_file) as pdf:
                total_pages = len(pdf.pages)
                
                # 逐頁讀取，並同步更新 UI 進度條
                for i, page in enumerate(pdf.pages):
                    status_text.text(f"⏳ 正在處理第 {i+1} / {total_pages} 頁，請稍候...")
                    progress_bar.progress((i + 1) / total_pages) # 更新進度條比例
                    
                    table = page.extract_table()
                    if table:
                        all_table_data.extend(table)
            
            if not all_table_data:
                st.error("❌ 抱歉，在這份 PDF 中找不到可以辨識的表格結構！")
            else:
                # ================= 5. 資料清洗邏輯 =================
                status_text.text("🔄 正在清洗與整理資料 (合併儲存格處理)...")
                
                df = pd.DataFrame(all_table_data)
                df = df.replace('\n', '', regex=True)
                
                # 取前 7 欄並設定標準欄位名稱
                df_selected = df.iloc[:, 0:7].copy()
                df_selected.columns = [
                    "項次", "(農藥項次) 國際普通名稱", "普通名稱", 
                    "作物類別", "作物", "修正後容許量(ppm)", "修正前容許量(ppm)"
                ]
                
                # 處理合併儲存格造成的空白
                df_selected.replace('', pd.NA, inplace=True)
                df_selected.replace(r'^\s*$', pd.NA, regex=True, inplace=True)
                cols_to_ffill = ["項次", "(農藥項次) 國際普通名稱", "普通名稱", "作物類別"]
                df_selected[cols_to_ffill] = df_selected[cols_to_ffill].fillna(method='ffill')
                
                # 過濾標題列與空白列
                df_selected = df_selected.dropna(subset=['作物', '修正後容許量(ppm)'])
                df_selected = df_selected[df_selected['項次'] != '項次']
                
                # ================= 6. 準備匯出檔案 =================
                status_text.text("✅ 處理完成！準備產生下載檔案...")
                
                # 將整理好的表格展示在畫面上，讓你可以預覽前幾筆
                st.markdown("### 📊 資料預覽 (前 5 筆)")
                st.dataframe(df_selected.head())
                
                # 將 Pandas DataFrame 轉成 Excel 檔案格式存入記憶體
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_selected.to_excel(writer, index=False, sheet_name='農藥容許量')
                processed_data = output.getvalue()
                
                st.success("🎉 轉換成功！請點擊下方按鈕下載。")
                
                # 下載按鈕
                st.download_button(
                    label="📥 下載整理好的 Excel 檔案",
                    data=processed_data,
                    file_name="農藥容許量自動整理.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
        except Exception as e:
            st.error(f"❌ 處理過程中發生錯誤：{e}")