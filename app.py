import streamlit as st
import pandas as pd
from docx import Document
import io
import zipfile
from datetime import datetime

# --- 頁面設定 ---
st.set_page_config(page_title="銀行驗收單生成器-修復版", page_icon="🏦", layout="wide")
st.title("🏦 銀行驗收單自動生成系統 (v2.1)")

# --- 函式：替換文字 (進階強化版) ---
def replace_text_in_document(doc, replacements):
    # 遍歷所有段落
    for paragraph in doc.paragraphs:
        for key, value in replacements.items():
            if key in paragraph.text:
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
    # 這裡不過濾掉日期空值，但在產出時處理
    
    # --- 篩選介面 ---
    st.header("⚙️ 篩選與產出設定")
    col1, col2 = st.columns(2)
    
    with col1:
        # --- 關鍵修復點：處理工程師欄位的空值與型別 ---
        # 1. 轉為字串 2. 過濾掉 'nan' 或空字串 3. 排序
        engineer_list = df['工程師'].astype(str).unique().tolist()
        all_engineers = sorted([eng for eng in engineer_list if eng.lower() != 'nan' and eng.strip() != ''])
        
        selected_engineers = st.multiselect("選擇需要的工程師 (請至少選一個)：", options=all_engineers, default=[])
        
    with col2:
        # 日期區間篩選 (處理日期可能全部為空的情況)
        valid_dates = df['日期物件'].dropna()
        if not valid_dates.empty:
            min_date = valid_dates.min().date()
            max_date = valid_dates.max().date()
            date_range = st.date_input("選擇日期區間：", [min_date, max_date])
        else:
            st.error("Excel 中找不到有效的『汰換日期』資料。")
            date_range = []
    
    # 執行過濾
    if len(date_range) == 2 and selected_engineers:
        start_date, end_date = date_range
        mask = (df['工程師'].astype(str).isin(selected_engineers)) & \
               (df['日期物件'].dt.date >= start_date) & \
               (df['日期物件'].dt.date <= end_date)
        final_df = df[mask]
    else:
        final_df = pd.DataFrame()
        if not selected_engineers:
            st.warning("⚠️ 請從左上方選單選擇工程師名字。")

    st.write(f"📊 目前篩選條件下共有 **{len(final_df)}** 筆資料。")

    # --- 執行產出 ---
    if st.button("🚀 開始生成並打包檔案") and not final_df.empty:
        zip_buffer = io.BytesIO()
        progress_bar = st.progress(0)
        
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for i, (idx, row) in enumerate(final_df.iterrows()):
                uploaded_word.seek(0)
                doc = Document(uploaded_word)
                
                # 1. 處理機號與財編 (排除 nan)
                m_id = str(row.get('機號', '')).strip()
                if m_id.lower() == 'nan' or not m_id: m_id = "(缺機號)"
                
                a_id = str(row.get('CUB財編', '')).strip()
                if a_id.lower() == 'nan' or not a_id: a_id = "(缺財編)"

                # 2. 處理日期格式 (需求：YYYY年MM月DD日)
                raw_date = row.get('日期物件')
                try:
                    formatted_date = raw_date.strftime("%Y年%m月%d日") if not pd.isna(raw_date) else "(日期缺失)"
                except:
                    formatted_date = "(日期格式錯誤)"
                
                # 3. 邏輯判斷：4G 型號
                is_4g_val = str(row.get('4G', '')).strip()
                model_text = "FortiGate40F 3G/4G" if "含4G" in is_4g_val else "FortiGate40F"
                
                # 定義取代字典 (確保所有 Key 與 Word 標籤一致)
                replacements = {
                    "{{Date}}": formatted_date,
                    "{{Station}}": str(row.get('站點名稱', '')).replace('nan', '(缺站點)'),
                    "{{MachineID}}": m_id,
                    "{{Address}}": str(row.get('地址', '')).replace('nan', '(缺地址)'),
                    "{{SN}}": str(row.get('機器序號', '')).replace('nan', '(缺序號)'),
                    "{{AssetID}}": a_id,
                    "{{SIM}}": str(row.get('SIM卡編號', '')).replace('nan', '(缺SIM)'),
                    "{{IP}}": str(row.get('SIM卡IP', '')).replace('nan', '(缺IP)'),
                    "{{Model}}": model_text
                }
                
                # 執行替換
                replace_text_in_document(doc, replacements)
                
                doc_io = io.BytesIO()
                doc.save(doc_io)
                doc_io.seek(0)
                
                # 檔名命名規則 (加上錯誤標記)
                error_tag = "[Error]" if "(缺" in m_id or "(缺" in a_id or "缺失" in formatted_date else ""
                file_date = raw_date.strftime("%Y%m%d") if not pd.isna(raw_date) else "NoDate"
                safe_station = str(row.get('站點名稱', 'Unknown')).replace('/', '_').replace('nan', '')
                file_name = f"{error_tag}{file_date}_{m_id}_{safe_station}.docx"
                
                zip_file.writestr(file_name, doc_io.getvalue())
                progress_bar.progress((i + 1) / len(final_df))
        
        st.success("✅ 產出完成！")
        st.download_button(
            label="📥 下載所有 Word 檔案 (ZIP)",
            data=zip_buffer.getvalue(),
            file_name=f"驗收單產出_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
            mime="application/zip"
        )
else:
    st.info("請在左側上傳 Excel (.xlsx) 與 Word (.docx) 檔案。")
