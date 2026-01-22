import streamlit as st
import pandas as pd
from docx import Document
import io
import zipfile
from datetime import datetime

# --- 頁面設定 ---
st.set_page_config(page_title="銀行驗收單生成器-優化版", page_icon="🏦", layout="wide")
st.title("🏦 銀行驗收單自動生成系統 (v2.0)")

# --- 函式：替換文字 (進階強化版) ---
def replace_text_in_document(doc, replacements):
    # 遍歷所有段落
    for paragraph in doc.paragraphs:
        for key, value in replacements.items():
            if key in paragraph.text:
                # 遍歷 run 以保持格式，但需處理變數被拆分在不同 run 的情況
                for run in paragraph.runs:
                    if key in run.text:
                        run.text = run.text.replace(key, str(value))
    
    # 遍歷所有表格
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
    # 讀取 Excel (強制轉字串避免科學記號)
    df = pd.read_excel(uploaded_excel, dtype=str)
    df.columns = df.columns.str.strip() 
    
    # 日期預處理
    df['日期物件'] = pd.to_datetime(df['汰換日期'], errors='coerce')
    df = df.dropna(subset=['日期物件'])
    
    # --- 篩選介面 ---
    st.header("⚙️ 篩選與產出設定")
    col1, col2 = st.columns(2)
    
    with col1:
        # 工程師篩選 (需求 5: 預設不勾選)
        all_engineers = sorted(df['工程師'].unique().tolist())
        selected_engineers = st.multiselect("選擇需要的工程師 (請至少選一個)：", options=all_engineers, default=[])
        
    with col2:
        # 日期區間篩選
        min_date = df['日期物件'].min().date()
        max_date = df['日期物件'].max().date()
        date_range = st.date_input("選擇日期區間：", [min_date, max_date])
    
    # 執行過濾
    if len(date_range) == 2 and selected_engineers:
        start_date, end_date = date_range
        mask = (df['工程師'].isin(selected_engineers)) & \
               (df['日期物件'].dt.date >= start_date) & \
               (df['日期物件'].dt.date <= end_date)
        final_df = df[mask]
    else:
        final_df = pd.DataFrame()
        if not selected_engineers:
            st.warning("⚠️ 請從上方選單選擇工程師名字以開始產出。")

    st.write(f"📊 目前篩選條件下共有 **{len(final_df)}** 筆資料。")

    # --- 執行產出 ---
    if st.button("🚀 開始生成並打包檔案") and not final_df.empty:
        zip_buffer = io.BytesIO()
        progress_bar = st.progress(0)
        
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for i, (idx, row) in enumerate(final_df.iterrows()):
                # 重新載入範本
                uploaded_word.seek(0)
                doc = Document(uploaded_word)
                
                # 1. 處理機號與財編 (需求 1, 2, 3)
                m_id = str(row.get('機號', '')).strip()
                if not m_id or m_id == 'nan': m_id = "(缺機號)"
                
                a_id = str(row.get('CUB財編', '')).strip()
                if not a_id or a_id == 'nan': a_id = "(缺財編)"

                # 2. 處理日期格式 (需求 4)
                raw_date = row.get('日期物件')
                formatted_date = raw_date.strftime("%Y年%m月%d日") if not pd.isna(raw_date) else "日期錯誤"
                
                # 3. 邏輯判斷：4G 型號
                is_4g_val = str(row.get('4G', '')).strip()
                model_text = "FortiGate40F 3G/4G" if "含4G" in is_4g_val else "FortiGate40F"
                
                # 定義取代字典
                replacements = {
                    "{{Date}}": formatted_date,
                    "{{Station}}": str(row.get('站點名稱', '')),
                    "{{MachineID}}": m_id,
                    "{{Address}}": str(row.get('地址', '')),
                    "{{SN}}": str(row.get('機器序號', '')),
                    "{{AssetID}}": a_id,
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
                
                # 檔名命名規則 (需求 1: 空值標記 Error)
                error_tag = "[Error]" if "(缺" in m_id or "(缺" in a_id else ""
                file_date = raw_date.strftime("%Y%m%d")
                safe_station = str(row.get('站點名稱', '')).replace('/', '_')
                file_name = f"{error_tag}{file_date}_{m_id}_{safe_station}.docx"
                
                # 寫入 ZIP
                zip_file.writestr(file_name, doc_io.getvalue())
                progress_bar.progress((i + 1) / len(final_df))
        
        st.success("✅ 產出完成！")
        st.download_button(
            label="📥 下載所有 Word 檔案 (ZIP)",
            data=zip_buffer.getvalue(),
            file_name=f"驗收單_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
            mime="application/zip"
        )
else:
    st.info("請上傳檔案以繼續。")
