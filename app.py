import streamlit as st
import pandas as pd
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt
import io
import zipfile
from datetime import datetime

# --- 頁面設定 ---
st.set_page_config(page_title="銀行驗收單生成器-標楷體版", page_icon="🏦", layout="wide")
st.title("🏦 銀行驗收單自動生成系統 (v2.3)")

# --- 函式：設定字體為標楷體 ---
def set_font_kai(run):
    run.font.name = '標楷體'
    # 這是關鍵：必須強制指定東亞字體(East Asia)為標楷體
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

# --- 函式：替換文字並保留字體格式 ---
def replace_text_in_document(doc, replacements):
    # 處理段落
    for paragraph in doc.paragraphs:
        for key, value in replacements.items():
            if key in paragraph.text:
                # 為了避免變數被 Word 拆成多個 run，我們進行合併取代
                # 1. 紀錄原本段落中是否含有該變數
                # 2. 直接在段落層級取代文字
                full_text = paragraph.text.replace(key, str(value))
                # 3. 清空原本的 runs 並重新寫入，確保字體一致
                for run in paragraph.runs:
                    run.text = ""
                new_run = paragraph.add_run(full_text)
                set_font_kai(new_run)
    
    # 處理表格內容 (大部分驗收資料都在這)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for key, value in replacements.items():
                        if key in paragraph.text:
                            # 針對表格內的格子進行強制取代與字體設定
                            new_text = paragraph.text.replace(key, str(value))
                            # 清空原本 runs
                            for run in paragraph.runs:
                                run.text = ""
                            # 建立新的 run 並鎖定標楷體
                            new_run = paragraph.add_run(new_text)
                            set_font_kai(new_run)

# --- 側邊欄：檔案上傳 ---
st.sidebar.header("📁 檔案上傳")
uploaded_excel = st.sidebar.file_uploader("1. 上傳 Excel 清單 (.xlsx)", type=['xlsx'])
uploaded_word = st.sidebar.file_uploader("2. 上傳 Word 範本 (.docx)", type=['docx'])

if uploaded_excel and uploaded_word:
    # 讀取 Excel
    df = pd.read_excel(uploaded_excel, dtype=str)
    df.columns = df.columns.str.strip() 
    
    # 日期預處理
    df['日期物件'] = pd.to_datetime(df['汰換日期'], errors='coerce')
    
    st.header("⚙️ 篩選與產出設定")
    col1, col2 = st.columns(2)
    
    with col1:
        # 預設不勾選
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

    if len(date_range) == 2 and selected_engineers:
        start_date, end_date = date_range
        mask = (df['工程師'].astype(str).isin(selected_engineers)) & \
               (df['日期物件'].dt.date >= start_date) & \
               (df['日期物件'].dt.date <= end_date)
        final_df = df[mask]
    else:
        final_df = pd.DataFrame()

    st.write(f"📊 目前篩選條件下共有 **{len(final_df)}** 筆資料。")

    if st.button("🚀 開始生成標楷體驗收單") and not final_df.empty:
        zip_buffer = io.BytesIO()
        progress_bar = st.progress(0)
        
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for i, (idx, row) in enumerate(final_df.iterrows()):
                uploaded_word.seek(0)
                doc = Document(uploaded_word)
                
                # 1. 處理機號與財編 (確保填入)
                m_id = str(row.get('機號', '')).strip()
                if m_id.lower() == 'nan' or not m_id: m_id = "(缺機號)"
                
                a_id = str(row.get('CUB財編', '')).strip()
                if a_id.lower() == 'nan' or not a_id: a_id = "(缺財編)"

                # 2. 處理日期格式 (需求: YYYY年MM月DD日)
                raw_date = row.get('日期物件')
                formatted_date = raw_date.strftime("%Y年%m月%d日") if not pd.isna(raw_date) else "(日期缺失)"
                
                # 3. 邏輯判斷：4G 欄位為「無」時
                is_4g_val = str(row.get('4G', '')).strip()
                if is_4g_val == "無":
                    sim_val = ""
                    ip_val = ""
                    model_text = "FortiGate40F"
                else:
                    sim_val = str(row.get('SIM卡編號', '')).replace('nan', '')
                    ip_val = str(row.get('SIM卡IP', '')).replace('nan', '')
                    model_text = "FortiGate40F 3G/4G"
                
                # 定義取代字典
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
                
                # 執行替換 (此函式內已包含標楷體鎖定)
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
        
        st.success("✅ 全部標楷體文件已生成完成！")
        st.download_button(
            label="📥 下載所有標楷體 Word 檔案 (ZIP)",
            data=zip_buffer.getvalue(),
            file_name=f"標楷體驗收單_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
            mime="application/zip"
        )
else:
    st.info("請上傳 Excel 與 Word 檔案以開始操作。")
