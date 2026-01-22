import streamlit as st
import pandas as pd
from docx import Document
import io
import zipfile
from datetime import datetime

# --- 頁面設定 ---
st.set_page_config(page_title="銀行驗收單生成器", page_icon="🏦", layout="wide")
st.title("🏦 銀行驗收單自動生成系統")
st.info("請上傳 Word 範本與 Excel 清單，系統將自動根據篩選條件產出對應的驗收單。")

# --- 函式：替換文字 (保留格式) ---
def replace_text_in_document(doc, replacements):
    # 替換段落
    for paragraph in doc.paragraphs:
        for key, value in replacements.items():
            if key in paragraph.text:
                for run in paragraph.runs:
                    if key in run.text:
                        run.text = run.text.replace(key, str(value))
    
    # 替換表格內容
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for key, value in replacements.items():
                        if key in paragraph.text:
                            for run in paragraph.runs:
                                if key in run.text:
                                    run.text = run.text.replace(key, str(value))

# --- 側邊欄：檔案上傳 ---
st.sidebar.header("📁 檔案上傳")
uploaded_excel = st.sidebar.file_uploader("1. 上傳 Excel 清單 (.xlsx)", type=['xlsx'])
uploaded_word = st.sidebar.file_uploader("2. 上傳 Word 範本 (.docx)", type=['docx'])

if uploaded_excel and uploaded_word:
    # 讀取 Excel
    df = pd.read_excel(uploaded_excel, dtype=str)
    df.columns = df.columns.str.strip() # 去除欄位名稱空格
    
    # 日期預處理 (嘗試轉換 Excel 的日期)
    df['日期物件'] = pd.to_datetime(df['汰換日期'], errors='coerce')
    df = df.dropna(subset=['日期物件']) # 排除無日期資料
    
    # --- 篩選介面 ---
    st.header("⚙️ 篩選與產出設定")
    col1, col2 = st.columns(2)
    
    with col1:
        # 工程師篩選 (需求 2)
        all_engineers = df['工程師'].unique().tolist()
        selected_engineers = st.multiselect("選擇工程師：", options=all_engineers, default=all_engineers)
        
    with col2:
        # 日期區間篩選 (需求 3)
        min_date = df['日期物件'].min().date()
        max_date = df['日期物件'].max().date()
        date_range = st.date_input("選擇日期區間：", [min_date, max_date])
    
    # 執行資料過濾
    if len(date_range) == 2:
        start_date, end_date = date_range
        mask = (df['工程師'].isin(selected_engineers)) & \
               (df['日期物件'].dt.date >= start_date) & \
               (df['日期物件'].dt.date <= end_date)
        final_df = df[mask]
    else:
        final_df = pd.DataFrame()

    st.write(f"📊 目前篩選條件下共有 **{len(final_df)}** 筆資料。")

    # --- 執行產出 ---
    if st.button("🚀 開始生成並打包檔案"):
        if final_df.empty:
            st.error("目前篩選結果為空，請調整篩選條件。")
        else:
            zip_buffer = io.BytesIO()
            progress_bar = st.progress(0)
            
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for i, (idx, row) in enumerate(final_df.iterrows()):
                    # 重新載入範本
                    uploaded_word.seek(0)
                    doc = Document(uploaded_word)
                    
                    # 邏輯判斷：4G 型號 (需求 5)
                    is_4g_val = str(row.get('4G', '')).strip()
                    model_text = "FortiGate40F 3G/4G" if "含4G" in is_4g_val else "FortiGate40F"
                    
                    # 定義取代字典 (排除需求 4 的工程師變數)
                    replacements = {
                        "{{Date}}": str(row.get('汰換日期', '')).split(' ')[0],
                        "{{Station}}": str(row.get('站點名稱', '')),
                        "{{MachineID}}": str(row.get('機號', '')),
                        "{{Address}}": str(row.get('地址', '')),
                        "{{SN}}": str(row.get('機器序號', '')),
                        "{{AssetID}}": str(row.get('CUB財編', '')),
                        "{{SIM}}": str(row.get('SIM卡編號', '')),
                        "{{IP}}": str(row.get('SIM卡IP', '')),
                        "{{Model}}": model_text
                    }
                    
                    # 執行替換
                    replace_text_in_document(doc, replacements)
                    
                    # 產出檔案到記憶體
                    doc_io = io.BytesIO()
                    doc.save(doc_io)
                    doc_io.seek(0)
                    
                    # 檔名命名規則
                    safe_station = str(row.get('站點名稱', '')).replace('/', '_')
                    file_name = f"{replacements['{{Date}}']}_{row.get('機號', '')}_{safe_station}.docx"
                    
                    # 寫入 ZIP
                    zip_file.writestr(file_name, doc_io.getvalue())
                    progress_bar.progress((i + 1) / len(final_df))
            
            st.success("✅ 產出完成！")
            st.download_button(
                label="📥 下載所有 Word 檔案 (ZIP)",
                data=zip_buffer.getvalue(),
                file_name=f"驗收單產出_{datetime.now().strftime('%Y%m%d')}.zip",
                mime="application/zip"
            )
else:
    st.warning("請先在左側上傳必要的 Excel 與 Word 檔案。")