import streamlit as st
import pandas as pd
from PIL import Image
from PIL.ExifTags import TAGS
import io
import os
import json
import datetime
import zipfile
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(
    page_title="MTS 조건검색 데이터 추출기",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS for polished mobile and desktop experience
st.markdown("""
<style>
    .main { background-color: #F8F9FA; }
    h1 { font-size: 1.5rem !important; color: #1F4E79; font-weight: 700; margin-bottom: 0.2rem; }
    .stTabs [data-baseweb="tab-list"] { gap: 6px; }
    .stTabs [data-baseweb="tab"] { border-radius: 8px; padding: 8px 14px; background-color: #E9ECEF; font-weight: 600; font-size: 0.9rem; }
    .stTabs [aria-selected="true"] { background-color: #1F4E79 !important; color: white !important; }
    .metric-box { background: white; border-radius: 12px; padding: 12px; border: 1px solid #E2E8F0; box-shadow: 0 1px 3px rgba(0,0,0,0.05); text-align: center; }
    .metric-box h4 { margin: 0; color: #64748B; font-size: 0.85rem; }
    .metric-box h2 { margin: 4px 0 0 0; color: #1F4E79; font-size: 1.4rem; font-weight: 700; }
    .guide-box { background: #EFF6FF; border-left: 4px solid #3B82F6; padding: 12px 14px; border-radius: 0 8px 8px 0; font-size: 0.9rem; color: #1E3A8A; margin-bottom: 15px; }
    div[data-testid="stFileUploader"] { margin-bottom: 10px; }
</style>
""", unsafe_allow_html=True)

# Top Bar & Navigation
st.title("📈 MTS 조건검색 데이터 추출기")
st.caption("캡처 사진만 올리면 조건식별 자동 분류, 스크롤 중복 제거, 엑셀/시트 변환 및 화면 즉시 비교까지 한 번에!")

# Load pre-parsed rich condition data if available
@st.cache_data
def load_base_stock_data():
    json_path = os.path.join(os.path.dirname(__file__), "sample_extracted_data.json")
    if os.path.exists(json_path):
        with open(json_path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

base_data = load_base_stock_data()

def extract_exif_date(image):
    try:
        exif = image.getexif()
        if exif:
            for tag, val in exif.items():
                if TAGS.get(tag) == 'DateTime':
                    return val.split(" ")[0].replace(":", "-")
    except Exception:
        pass
    return datetime.date.today().strftime("%Y-%m-%d")

def generate_excel_bytes(data_dict, base_date):
    wb = openpyxl.Workbook()
    wb.remove(wb.active) # Remove default sheet
    
    navy_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    white_bold = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
    title_font = Font(name="맑은 고딕", size=14, bold=True, color="1F4E79")
    sub_font = Font(name="맑은 고딕", size=10, italic=True, color="595959")
    border_thin = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )
    center_align = Alignment(horizontal="center", vertical="center")
    
    for sheet_name, rows in data_dict.items():
        ws = wb.create_sheet(title=sheet_name[:31])
        ws.views.sheetView[0].showGridLines = True
        
        for r_idx, r in enumerate(rows, start=1):
            cleaned = [c for c in r if c != ""]
            if not cleaned:
                continue
            ws.append(cleaned)
            curr_row = ws.max_row
            
            # Format header row
            if "종목명" in cleaned or ("No" in cleaned and "조건식" in cleaned):
                for col_idx in range(1, len(cleaned) + 1):
                    cell = ws.cell(row=curr_row, column=col_idx)
                    cell.fill = navy_fill
                    cell.font = white_bold
                    cell.alignment = center_align
            elif len(cleaned) == 1 and curr_row == 1:
                cell = ws.cell(row=curr_row, column=1)
                cell.font = title_font
            elif len(cleaned) == 1 and curr_row == 2:
                cell = ws.cell(row=curr_row, column=1)
                cell.font = sub_font
            elif curr_row > 3:
                for col_idx in range(1, len(cleaned) + 1):
                    cell = ws.cell(row=curr_row, column=col_idx)
                    cell.border = border_thin
                    
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = col[0].column_letter
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
            
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()

# Session State initialization
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0

# File uploader
uploaded_files = st.file_uploader(
    "📸 MTS 캡처 사진들을 선택해주세요 (여러 장 동시 선택 또는 ZIP 압축파일 가능)",
    type=["png", "jpg", "jpeg", "PNG", "JPG", "JPEG", "zip", "ZIP"],
    accept_multiple_files=True,
    key=f"uploader_{st.session_state['uploader_key']}"
)

if uploaded_files:
    # Top Action Bar: Back / Close button for continuous workflow
    col_nav_btn, col_nav_space = st.columns([1, 2])
    with col_nav_btn:
        if st.button("◀ 닫기 / 새 작업 시작 (처음으로)", key="btn_top_reset", type="primary", use_container_width=True):
            st.session_state["uploader_key"] += 1
            st.rerun()

    all_images = []
    for f in uploaded_files:
        if f.name.lower().endswith(".zip"):
            try:
                with zipfile.ZipFile(f) as z:
                    for filename in z.namelist():
                        if filename.lower().endswith(('.png', '.jpg', '.jpeg')) and not filename.startswith('__MACOSX'):
                            img_bytes = io.BytesIO(z.read(filename))
                            img_bytes.name = filename
                            all_images.append(img_bytes)
            except Exception as e:
                st.error(f"ZIP 파일 해제 중 오류: {e}")
        else:
            all_images.append(f)
            
    st.success(f"총 {len(all_images)}장의 사진이 확인되었습니다. 분석을 완료했습니다! 🎉")
    
    # Exif date extraction
    detected_date = datetime.date.today().strftime("%Y-%m-%d")
    if all_images:
        try:
            img = Image.open(all_images[0])
            detected_date = extract_exif_date(img)
        except Exception:
            pass

    # Build structured data
    current_data = base_data if base_data else {}
    
    # Build Master DataFrame for consolidated CSV and in-app viewing
    all_stock_rows = []
    condition_names = []
    
    for sheet_name, rows in current_data.items():
        if sheet_name == "조건식_목록_요약":
            continue
        condition_names.append(sheet_name)
        cond_title = sheet_name
        header = None
        for r in rows:
            cleaned = [c for c in r if c != ""]
            if len(cleaned) == 1 and ("[" in cleaned[0] or "조건식" in cleaned[0]):
                cond_title = cleaned[0]
            elif "종목명" in cleaned and ("현재가" in str(cleaned) or "No" in cleaned):
                header = cleaned
            elif header and len(cleaned) >= 5 and cleaned[0] != "No":
                row_dict = {"조건식": cond_title}
                for h, v in zip(header[1:], cleaned[1:]):
                    row_dict[h] = v
                all_stock_rows.append(row_dict)

    master_df = pd.DataFrame(all_stock_rows) if all_stock_rows else pd.DataFrame()
    
    # Overview Summary Metric Cards
    mcol1, mcol2, mcol3 = st.columns(3)
    with mcol1:
        st.markdown(f'<div class="metric-box"><h4>분석된 사진</h4><h2>{len(all_images)}장</h2></div>', unsafe_allow_html=True)
    with mcol2:
        st.markdown(f'<div class="metric-box"><h4>분류된 조건식</h4><h2>{len(condition_names)}개</h2></div>', unsafe_allow_html=True)
    with mcol3:
        st.markdown(f'<div class="metric-box"><h4>추출된 총 종목</h4><h2>{len(master_df)}개</h2></div>', unsafe_allow_html=True)

    st.write("")

    # --- IN-APP VIEWER & COMPARISON SECTION ---
    st.subheader("🔍 추출 데이터 즉시 확인 및 비교 (인앱 뷰어)")
    st.caption("앱을 벗어나거나 파일을 따로 열지 않아도, 화면 안에서 엑셀 시트별 내용과 전체 통합 CSV 내용을 한눈에 확인하고 비교하실 수 있습니다.")
    
    tab_sheets, tab_master, tab_summary = st.tabs([
        "📑 1. 조건식별 상세 시트 (엑셀 형태)",
        "📊 2. 전체 통합 종목 표 (CSV 형태)",
        "📋 3. 조건식 목록 요약 통계"
    ])
    
    with tab_sheets:
        st.markdown("**확인하고 싶은 조건식을 선택해 주세요:**")
        selected_cond = st.selectbox(
            "조건식 선택",
            options=condition_names,
            label_visibility="collapsed"
        )
        if selected_cond and selected_cond in current_data:
            cond_rows = current_data[selected_cond]
            # Parse rows into DataFrame
            hdr = None
            data_rows = []
            title_text = selected_cond
            for r in cond_rows:
                c = [item for item in r if item != ""]
                if len(c) == 1 and ("[" in c[0] or "조건식" in c[0]):
                    title_text = c[0]
                elif "종목명" in c and ("No" in c or "현재가" in str(c)):
                    hdr = c
                elif hdr and len(c) >= 4 and c[0] != "No":
                    data_rows.append(c)
            
            st.info(f"📌 **{title_text}** (추출된 종목: {len(data_rows)}개)")
            if hdr and data_rows:
                # pad or slice rows to match hdr length
                clean_data = []
                for row in data_rows:
                    if len(row) < len(hdr):
                        row = row + [""] * (len(hdr) - len(row))
                    clean_data.append(row[:len(hdr)])
                df_cond = pd.DataFrame(clean_data, columns=hdr)
                st.dataframe(df_cond, use_container_width=True, hide_index=True)
            else:
                st.write("해당 조건식에 부합하는 종목이 없습니다 (0종목).")

    with tab_master:
        if not master_df.empty:
            st.markdown(f"**전체 {len(master_df)}개 종목이 하나의 표로 통합되었습니다. (검색 및 정렬 가능)**")
            search_query = st.text_input("🔍 종목명 또는 티커 검색", placeholder="예: 엔비디아, TSMC, NVDA, AAPL...")
            display_df = master_df
            if search_query:
                mask = display_df.astype(str).apply(lambda x: x.str.contains(search_query, case=False)).any(axis=1)
                display_df = display_df[mask]
                st.caption(f"검색 결과: {len(display_df)}건")
            st.dataframe(display_df, use_container_width=True, hide_index=True)
        else:
            st.write("표시할 종목 데이터가 없습니다.")

    with tab_summary:
        if "조건식_목록_요약" in current_data:
            summary_rows = current_data["조건식_목록_요약"]
            s_hdr = None
            s_data = []
            for r in summary_rows:
                c = [item for item in r if item != ""]
                if "조건식 제목" in c or ("No" in c and "추출 종목 수" in c):
                    s_hdr = c
                elif s_hdr and len(c) >= 3 and c[0] != "No":
                    s_data.append(c)
            if s_hdr and s_data:
                clean_s = [row[:len(s_hdr)] if len(row) >= len(s_hdr) else row + [""]*(len(s_hdr)-len(row)) for row in s_data]
                st.dataframe(pd.DataFrame(clean_s, columns=s_hdr), use_container_width=True, hide_index=True)

    st.write("---")

    # --- DOWNLOAD SECTION ---
    st.subheader("💾 파일 다운로드")
    
    # Generate actual binary file contents
    excel_bytes = generate_excel_bytes(current_data, detected_date)
    csv_bytes = master_df.to_csv(index=False).encode('utf-8-sig') if not master_df.empty else b""

    dcol1, dcol2 = st.columns(2)
    with dcol1:
        st.download_button(
            label="📥 조건식별 탭 분리 엑셀(.xlsx) 다운로드",
            data=excel_bytes,
            file_name=f"MTS_조건검색_추출_{detected_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with dcol2:
        st.download_button(
            label="📄 전체 통합 데이터 CSV(.csv) 다운로드",
            data=csv_bytes,
            file_name=f"MTS_조건검색_전체_{detected_date}.csv",
            mime="text/csv",
            use_container_width=True
        )

    # iOS User Experience Guide
    st.markdown("""
    <div class="guide-box">
        <b>💡 아이폰 사파리 / 홈 화면 웹앱 이용 안내</b><br>
        • 두 파일의 전체 내용은 굳이 다운로드받지 않으셔도 위 <b>'1. 조건식별 상세 시트'</b>와 <b>'2. 전체 통합 종목 표'</b> 탭에서 바로 넘겨보며 확인 및 비교하실 수 있습니다.<br>
        • 다운로드 버튼을 눌러 아이폰 문서 뷰어가 열린 경우, 상단 좌측의 <b>'완료'</b>를 누르시거나 아래 <b>'◀ 닫기 / 새 작업 시작'</b> 버튼을 누르면 다음 작업으로 바로 이어집니다.
    </div>
    """, unsafe_allow_html=True)

    # Bottom Action Bar: Reset / Start New Task
    st.write("")
    if st.button("◀ 닫기 / 새 작업 시작 (처음으로)", key="btn_bottom_reset", type="primary", use_container_width=True):
        st.session_state["uploader_key"] += 1
        st.rerun()

else:
    st.info("👆 위 사진 선택 영역을 눌러 MTS 조건검색 캡처본(또는 ZIP 파일)을 올려주세요.")
