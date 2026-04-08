import streamlit as st
import pandas as pd
from docx import Document
from docx.oxml.ns import qn
import io
import zipfile
from datetime import datetime
import re

# --- 頁面設定 ---
st.set_page_config(page_title="銀行驗收單生成器-修復版", page_icon="🏦", layout="wide")
st.title("🏦 銀行驗收單自動生成系統 (v2.5)")

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
    date_str = str(date_str).strip()
    clean_date = re.sub(r'[^0-9]', '', date_str)
    if len(clean_date) == 8:
        try:
            return datetime.strptime(clean_date, "%Y%m%d").date()
        except ValueError:
            return None
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
    # 讀取時強制所有欄位為字串
    df = pd.read_excel(uploaded_excel, dtype=str)
    df.columns = df.columns.str.strip() 
    df['日期物件'] = pd.to_datetime(df['汰換日期'], errors='coerce')
    
    st.header("⚙️ 篩選與產出設定")
    col1, col2 = st.columns(2)
    
    with col1:
        # --- 【核心修復邏輯】 ---
        # 1. 先用 dropna() 去除真正的空值
        # 2. 用列表推導式確保每個元素都是 str 且排除 'nan' 字串
        raw_list = df['工程師'].dropna().unique().tolist()
        all_engineers = sorted([str(eng).strip() for eng in raw_list if str(eng).strip().lower() != 'nan' and str(eng).strip() != ''])
        
        selected_engineers = st.multiselect("選擇需要的工程師：", options=all_engineers, default=[])
        
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
                    st.error("❌ 日期格式錯誤")
                    start_date, end_date = None, None
                elif start_date > end_date:
                    st.error("❌ 開始日期不可大於結束日期")
                    start_date, end_date = None, None

    # --- 資料過濾 ---
    if start_date and end_date and selected_engineers:
        # 過濾時也確保工程師欄位是字串比對
        mask = (df['工程師'].astype(str).str.strip().isin(selected_engineers)) & \
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
                
                # 安全抓取資料
                def get_val(col_name):
                    v = row.get(col_name, '')
                    return "" if pd.isna(v) or str(v).lower() == 'nan' else str(v).strip()

                m_id = get_val('機號') or "(缺機號)"
                a_id = get_val('CUB財編') or "(缺財編)"
                raw_date = row.get('日期物件')
                formatted_date = raw_date.strftime("%Y年%m月%d日") if pd.notna(raw_date) else "(日期缺失)"
                
                is_4g_val = get_val('4G')
                if is_4g_val == "無":
                    sim_val, ip_val, model_text = "", "", "FortiGate40F"
                else:
                    sim_val = get_val('SIM卡編號')
                    ip_val = get_val('SIM卡IP')
                    model_text = "FortiGate40F 3G/4G"
                
                replacements = {
                    "{{Date}}": formatted_date,
                    "{{Station}}": get_val('站點名稱'),
                    "{{MachineID}}": m_id,
                    "{{Address}}": get_val('地址'),
                    "{{SN}}": get_val('機器序號'),
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
                file_date = raw_date.strftime("%Y%m%d") if pd.notna(raw_date) else "NoDate"
                safe_station = get_val('站點名稱').replace('/', '_') or "Unknown"
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
