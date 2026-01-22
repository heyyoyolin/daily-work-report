import streamlit as st
import pandas as pd
from docx import Document
import io
import zipfile
from datetime import datetime

# --- 設定頁面標題 ---
st.set_page_config(page_title="銀行驗收單自動生成器", page_icon="🖨️")
st.title("🖨️ 銀行驗收單自動生成系統")
st.markdown("### 步驟：上傳檔案 -> 篩選資料 -> 下載 Word")

# --- 1. 檔案上傳區 ---
st.sidebar.header("📂 1. 請上傳檔案")
uploaded_excel = st.sidebar.file_uploader("上傳 Excel 清單 (.xlsx)", type=['xlsx'])
uploaded_word = st.sidebar.file_uploader("上傳 Word 範本 (.docx)", type=['docx'])

# --- 函式：替換段落文字 ---
def replace_text_in_paragraph(paragraph, replacements):
    for key, value in replacements.items():
        if key in paragraph.text:
            inline = paragraph.runs
            for i in range(len(inline)):
                if key in inline[i].text:
                    text = inline[i].text.replace(key, str(value))
                    inline[i].text = text

# --- 主程式邏輯 ---
if uploaded_excel and uploaded_word:
    try:
        # 讀取 Excel 資料
        # dtype=str 確保所有欄位都當作文字處理 (避免電話/SIM卡變成科學記號)
        df = pd.read_excel(uploaded_excel, dtype=str)
        
        # 資料清理：移除欄位名稱前後空白
        df.columns = df.columns.str.strip()
        
        # 確認是否有必要的欄位
        required_columns = ['工程師', '汰換日期', '機號', '站點名稱', '4G']
        if not all(col in df.columns for col in required_columns):
            st.error(f"Excel 缺少必要欄位，請檢查是否包含：{required_columns}")
            st.stop()

        # 日期格式處理 (轉換為 datetime 物件以便篩選)
        # 假設 Excel 日期格式可能為 "2025-11-18 00:00:00" 或 "2025-11-18"
        df['日期物件'] = pd.to_datetime(df['汰換日期'], errors='coerce')
        
        # 去除無日期的無效資料
        df = df.dropna(subset=['日期物件'])

        st.success(f"✅ 檔案讀取成功！共載入 {len(df)} 筆資料。")
        st.divider()

        # --- 2. 篩選條件區 ---
        st.header("🔍 2. 設定篩選條件")
        
        col1, col2 = st.columns(2)
        
        with col1:
            # 取得所有工程師名單
            unique_engineers = df['工程師'].unique().tolist()
            selected_engineers = st.multiselect(
                "選擇工程師 (可多選):",
                options=unique_engineers,
                default=unique_engineers
            )

        with col2:
            # 取得資料中的最小與最大日期作為預設值
            min_date = df['日期物件'].min().date()
            max_date = df['日期物件'].max().date()
            
            start_date = st.date_input("開始日期", min_date)
            end_date = st.date_input("結束日期", max_date)

        # 執行篩選
        mask = (
            (df['工程師'].isin(selected_engineers)) & 
            (df['日期物件'].dt.date >= start_date) & 
            (df['日期物件'].dt.date <= end_date)
        )
        filtered_df = df[mask]

        st.info(f"📊 根據篩選條件，即將產出 **{len(filtered_df)}** 份文件。")

        # --- 3. 產出與下載 ---
        if st.button("🚀 開始生成驗收單", type="primary"):
            if len(filtered_df) == 0:
                st.warning("⚠️ 沒有符合條件的資料，請重新調整篩選條件。")
            else:
                # 準備一個記憶體內的 ZIP 檔案
                zip_buffer = io.BytesIO()
                
                # 顯示進度條
                progress_bar = st.progress(0)
                
                with zipfile.ZipFile(zip_buffer, "w") as zf:
                    total_files = len(filtered_df)
                    
                    for i, (index, row) in enumerate(filtered_df.iterrows()):
                        # 讀取 Word 範本 (每次都要重新讀取原始檔)
                        uploaded_word.seek(0)
                        doc = Document(uploaded_word)
                        
                        # --- 邏輯判斷 (需求 5) ---
                        # 判斷 4G 欄位決定型號
                        is_4g = str(row.get('4G', '')).strip()
                        if '含4G' in is_4g:
                            model_name = "FortiGate40F 3G/4G"
                        else:
                            model_name = "FortiGate40F"

                        # --- 定義對應變數 ---
                        # 這裡移除了 {{Engineer}}，並加入了 {{Model}}
                        replacements = {
                            '{{Date}}': str(row.get('汰換日期', '')).split()[0], # 只取日期部分
                            '{{Station}}': row.get('站點名稱', ''),
                            '{{MachineID}}': row.get('機號', ''),
                            '{{Address}}': row.get('地址', ''),
                            '{{SN}}': row.get('機器序號', ''),
                            '{{AssetID}}': row.get('CUB財編', ''),
                            '{{SIM}}': row.get('SIM卡編號', ''),
                            '{{IP}}': row.get('SIM卡IP', ''),
                            '{{Model}}': model_name,  # 這裡填入自動判斷後的型號
                        }

                        # --- 執行替換 ---
                        # 替換段落
                        for paragraph in doc.paragraphs:
                            replace_text_in_paragraph(paragraph, replacements)

                        # 替換表格
                        for table in doc.tables:
                            for row_cell in table.rows:
                                for cell in row_cell.cells:
                                    for paragraph in cell.paragraphs:
                                        replace_text_in_paragraph(paragraph, replacements)

                        # --- 存入 ZIP ---
                        # 建立檔名：日期_工程師_機號_站點.docx
                        date_str = str(row.get('汰換日期', '')).split()[0]
                        eng_name = row.get('工程師', 'Unknown')
                        station_name = str(row.get('站點名稱', '')).replace('/', '_') # 避免檔名錯誤
                        file_name = f"{date_str}_{eng_name}_{row.get('機號', '')}_{station_name}.docx"
                        
                        # 將 Word 存到記憶體
                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        
                        # 寫入 ZIP
                        zf.writestr(file_name, doc_io.getvalue())
                        
                        # 更新進度條
                        progress_bar.progress((i + 1) / total_files)

                # 下載按鈕
                st.success("🎉 生成完成！請點擊下方按鈕下載。")
                st.download_button(
                    label="📥 下載所有驗收單 (ZIP壓縮檔)",
                    data=zip_buffer.getvalue(),
                    file_name="已產出驗收單.zip",
                    mime="application/zip"
                )

    except Exception as e:
        st.error(f"發生錯誤: {e}")
else:
    st.info("請先在左側欄位上傳 Excel 和 Word 範本檔案。")