import streamlit as st
import pandas as pd
from docx import Document
from docx.oxml.ns import qn
import io
import zipfile
from datetime import datetime
import re

# --- 頁面設定 ---
st.set_page_config(page_title="銀行驗收單生成器-日期檢核版", page_icon="🏦", layout="wide")
st.title("🏦 銀行驗收單自動生成系統 (v2.4)")

# --- 函式：設定字體為標楷體 ---
def set_font_kai(run):
    run.font.name = '標楷體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

# --- 函式：替換文字並保留字體格式 ---
def replace_text_in_document(doc, replacements):
    for paragraph in doc.paragraphs:
        for key, value in replacements.items():
            if key in paragraph.text:
                full_text = paragraph.text.replace(key, str(value))
                for run in paragraph.runs:
                    run.text = ""
                new_run = paragraph.add_run(full_text)
                set_font_kai(new_run)
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for key, value in replacements.items():
                        if key in paragraph.text:
                            new_text = paragraph.text.replace(key, str(value))
                            for run in paragraph.runs:
                                run.text = ""
                            new_run = paragraph.add_run(new_text)
                            set_font_kai(new_run)

# --- 函式：手動日期解析與檢核 ---
def parse_date(date_str):
    """
    支援多種日期格式解析：YYYY-MM-DD, YYYY/MM/DD, YYYYMMDD
    """
    date_str = str(date_str).strip()
    # 清除非數字字元以便嘗試解析
    clean_date = re.sub(r'[^0-9]', '', date_str)
    
    if len(clean_date) == 8:
        try:
            return datetime.strptime(clean_date, "%Y%m%d").date()
        except ValueError:
            return None
    
    # 嘗試一般格式
    for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(date_str, fmt).date()
        except ValueError:
            continue
    return None

# --- 側邊欄：檔案上傳 ---
st.sidebar.header("📁 檔案上傳")
uploaded_excel = st.sidebar.file_uploader("1. 上傳 Excel 清單 (.xlsx)", type=['xlsx'])
uploaded_word = st.sidebar.file_uploader("2. 上傳 Word 範本 (.docx)", type=['docx'])

if uploaded_excel and uploaded_word:
    df = pd.read_excel(uploaded_excel, dtype=str)
    df.columns = df.columns.str.strip() 
    df['日期物件'] = pd.to_datetime(df['汰換日期'], errors='coerce')
    
    st.header("⚙️ 篩選與產出設定")
    col1, col2 = st.columns(2)
    
    with col1:
        engineer_list = df['工程師'].astype(str).unique().tolist()
        all_engineers = sorted([eng for eng in engineer_list if eng.lower() != 'nan' and eng.strip() != ''])
        selected_engineers = st.multiselect("選擇工程師：", options=all_engineers, default=[])
        
    with col2:
        date_mode = st.radio("日期選擇方式：", ["日曆選擇器", "手動輸入區間"], horizontal=True)
        
        start_date, end_date = None, None
        
        if date_mode == "日曆選擇器":
            valid_dates = df['日期物件'].dropna()
            if not valid_dates.empty:
                dr = st.date_input("選擇日期區間：", [valid_dates.min().date(), valid_dates.max().date()])
                if len(dr) == 2:
                    start_date, end_date = dr
            else:
                st.error("Excel 中無有效日期。")
        
        else:
            c1, c2 = st.columns(2)
            with c1:
                s_input = st.text_input("開始日期 (例: 20251118)", "")
            with c2:
                e_input = st.text_input("結束日期 (例: 20251120)", "")
            
            if s_input and e_input:
                start_date = parse_date(s_input)
                end_date = parse_date(e_input)
                
                if not start_date or not end_date:
                    st.error("❌ 日期格式錯誤，請輸入 YYYYMMDD 或 YYYY-MM-DD")
                    start_date, end_date = None, None
                elif start_date > end_date:
                    st.error("❌ 開始日期不可大於結束日期")
                    start_date, end_date = None, None
                else:
                    st.success(f"✅ 已識別區間：{start_date} 至 {end_date}")

    # --- 資料過濾 ---
    if start_date and end_date and selected_engineers:
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
                
                # 欄位解析
                m_id = str(row.get('機號', '')).strip()
                if m_id.lower() == 'nan' or not m_id: m_id = "(缺機號)"
                
                a_id = str(row.get('CUB財編', '')).strip()
                if a_id.lower() == 'nan' or not a_id: a_id = "(缺財編)"

                raw_date = row.get('日期物件')
                formatted_date = raw_date.strftime("%Y年%m月%d日") if not pd.isna(raw_date) else "(日期缺失)"
                
                is_4g_val = str(row.get('4G', '')).strip()
                if is_4g_val == "無":
                    sim_val, ip_val, model_text = "", "", "FortiGate40F"
                else:
                    sim_val = str(row.get('SIM卡編號', '')).replace('nan', '')
                    ip_val = str(row.get('SIM卡IP', '')).replace('nan', '')
                    model_text = "FortiGate40F 3G/4G"
                
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
                
                replace_text_in_document(doc, replacements)
                
                doc_io = io.BytesIO()
                doc.save(doc_io)
                doc_io.seek(0)
                
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
            file_name=f"驗收單_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
            mime="application/zip"
        )
else:
    st.info("請上傳 Excel 與 Word 檔案以開始操作。")
