import streamlit as st
import pandas as pd
from docx import Document
import io
import zipfile
from datetime import datetime

# --- 頁面設定 ---
st.set_page_config(page_title="銀行驗收單生成器-最終修復版", page_icon="🏦", layout="wide")
st.title("🏦 銀行驗收單自動生成系統 (v2.2)")

# --- 函式：強化版文字替換 (解決 Word 變數斷裂問題) ---
def replace_text_in_paragraph(paragraph, key, value):
    if key in paragraph.text:
        # 這是最穩定的替換方法：先在 paragraph 層級替換
        # 但為了保留格式，我們需要一些技巧
        full_text = paragraph.text.replace(key, str(value))
        # 覆蓋掉原本的 runs
        for run in paragraph.runs:
            run.text = ""
        paragraph.runs[0].text = full_text

def replace_text_in_document(doc, replacements):
    # 處理所有段落
    for paragraph in doc.paragraphs:
        for key, value in replacements.items():
            if key in paragraph.text:
                # 使用簡單覆蓋法，這對標籤替換最有效
                inline = paragraph.runs
                for i in range(len(inline)):
                    if key in inline[i].text:
                        inline[i].text = inline[i].text.replace(key, str(value))
                    # 處理標籤被切斷在不同 run 的情況
                    elif "{{" in paragraph.text and "}}" in paragraph.text:
                        # 如果段落中有標籤但 run 沒抓到，強行合併處理
                        paragraph.text = paragraph.text.replace(key, str(value))
    
    # 處理所有表格
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for key, value in replacements.items():
                        if key in paragraph.text:
                            # 強制直接替換 cell 內的段落文字
                            paragraph.text = paragraph.text.replace(key, str(value))

# --- 側邊欄：檔案上傳 ---
st.sidebar.header("📁 檔案上傳")
uploaded_excel = st.sidebar.file_uploader("1. 上傳 Excel 清單 (.xlsx)", type=['xlsx'])
uploaded_word = st.sidebar.file_uploader("2. 上傳 Word 範本 (.docx)", type=['docx'])

if uploaded_excel and uploaded_word:
    # 讀取 Excel (強制轉字串)
    df = pd.read_excel(uploaded_excel, dtype=str)
    df.columns = df.columns.str.strip() 
    
    # 日期預處理
    df['日期物件'] = pd.to_datetime(df['汰換日期'], errors='coerce')
    
    st.header("⚙️ 篩選與產出設定")
    col1, col2 = st.columns(2)
    
    with col1:
        engineer_list = df['工程師'].astype(str).unique().tolist()
        all_engineers = sorted([eng for eng in engineer_list if eng.lower() != 'nan' and eng.strip() != ''])
        selected_engineers = st.multiselect("選擇需要的工程師：", options=all_engineers, default=[])
        
    with col2:
        valid_dates = df['日期物件'].dropna()
        if not valid_dates.empty:
            min_date = valid_dates.min().date()
            max_date = valid_dates.max().date()
            date_range = st.date_input("選擇日期區間：", [min_date, max_date])
        else:
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

    st.write(f"📊 目前篩選條件下共有 **{len(final_df)}** 筆資料。")

    if st.button("🚀 開始生成並打包檔案") and not final_df.empty:
        zip_buffer = io.BytesIO()
        progress_bar = st.progress(0)
        
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for i, (idx, row) in enumerate(final_df.iterrows()):
                uploaded_word.seek(0)
                doc = Document(uploaded_word)
                
                # 1. 處理機號與財編 (需求 1, 2: 修正無法填入問題)
                m_id = str(row.get('機號', '')).strip()
                if m_id.lower() == 'nan' or not m_id: m_id = "(缺機號)"
                
                a_id = str(row.get('CUB財編', '')).strip()
                if a_id.lower() == 'nan' or not a_id: a_id = "(缺財編)"

                # 2. 處理日期格式 (需求 4)
                raw_date = row.get('日期物件')
                formatted_date = raw_date.strftime("%Y年%m月%d日") if not pd.isna(raw_date) else "(日期缺失)"
                
                # 3. 邏輯判斷：4G 欄位為「無」時，SIM 與 IP 保留空格 (需求 3, 5)
                is_4g_val = str(row.get('4G', '')).strip()
                if is_4g_val == "無":
                    sim_val = ""
                    ip_val = ""
                    model_text = "FortiGate40F"
                else:
                    sim_val = str(row.get('SIM卡編號', '')).replace('nan', '')
                    ip_val = str(row.get('SIM卡IP', '')).replace('nan', '')
                    model_text = "FortiGate40F 3G/4G"
                
                # 定義取代字典 (確保 Key 與 Word 內的標籤完全一致)
                replacements = {
                    "{{Date}}": formatted_date,
                    "{{Station}}": str(row.get('站點名稱', '')).replace('nan', ''),
                    "{{MachineID}}": m_id,
                    "{{Address}}": str(row.get('地址', '')).replace('nan', ''),
                    "{{SN}}": str(row.get('機器序號', '')).replace('nan', ''),
                    "{{AssetID}}": a_id,
                    "{{SIM}}": sim_val,
                    "{{IP}}": ip_val,
                    "{{Model}}": model_text
                }
                
                # 執行替換
                replace_text_in_document(doc, replacements)
                
                doc_io = io.BytesIO()
                doc.save(doc_io)
                doc_io.seek(0)
                
                # 檔名規則
                error_tag = "[Error]" if "(缺" in m_id or "(缺" in a_id else ""
                file_date = raw_date.strftime("%Y%m%d") if not pd.isna(raw_date) else "NoDate"
                safe_station = str(row.get('站點名稱', 'Unknown')).replace('/', '_').replace('nan', '')
                file_name = f"{error_tag}{file_date}_{m_id}_{safe_station}.docx"
                
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
    st.info("請上傳 Excel 與 Word 檔案以開始。")
