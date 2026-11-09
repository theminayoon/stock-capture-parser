import streamlit as st
import pandas as pd
from PIL import Image
from PIL.ExifTags import TAGS
import io
import os
import re
import json
import base64
import datetime
import zipfile
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from rapidocr_onnxruntime import RapidOCR

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
    .pwa-download-btn { display: block; width: 100%; padding: 10px 14px; background-color: #1F4E79; color: white !important; text-align: center; border-radius: 8px; font-weight: 600; text-decoration: none; font-size: 0.95rem; margin-top: 4px; box-sizing: border-box; }
    .pwa-download-btn:active { background-color: #153452; }
    div[data-testid="stFileUploader"] { margin-bottom: 10px; }
</style>
""", unsafe_allow_html=True)

# Top Bar & Navigation
st.title("📈 MTS 조건검색 데이터 추출기")
st.caption("캡처 사진만 올리면 조건식별 자동 분류, 스크롤 중복 제거, 엑셀/시트 변환 및 화면 즉시 비교까지 한 번에! (100% 무료 무제한)")

# Initialize OCR Engines (Cached so they only load once into memory)
@st.cache_resource
def get_ocr_engines():
    base_dir = os.path.dirname(__file__)
    model_path = os.path.join(base_dir, "korean_rec.onnx")
    dict_path = os.path.join(base_dir, "korean_dict.txt")
    
    eng_default = RapidOCR()
    if os.path.exists(model_path) and os.path.exists(dict_path):
        eng_korean = RapidOCR(rec_model_path=model_path, rec_keys_path=dict_path)
    else:
        eng_korean = eng_default
    return eng_default, eng_korean

ocr_default, ocr_korean = get_ocr_engines()

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

def parse_single_image(img_bytes_or_pil):
    if isinstance(img_bytes_or_pil, io.BytesIO):
        img_bytes_or_pil.seek(0)
        img = Image.open(img_bytes_or_pil)
    else:
        img = img_bytes_or_pil
        
    # Convert image to RGB if necessary
    if img.mode != "RGB":
        img = img.convert("RGB")
        
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=95)
    img_data = buf.getvalue()
    
    # 1. Condition Name Extraction using Korean OCR engine
    cond_title = "기타 조건식"
    try:
        res_ko, _ = ocr_korean(img_data)
        if res_ko:
            for item in sorted(res_ko, key=lambda x: x[0][0][1])[:15]:
                t = item[1].strip()
                if any(k in t for k in ["신고가", "정배열", "역배열", "주도주", "섹터", "후보군", "골든", "돌파", "["]):
                    # Clean up
                    cond_title = t
                    break
    except Exception:
        pass
        
    # 2. Extract stock rows using high-precision default engine
    stocks = []
    try:
        res_en, _ = ocr_default(img_data)
        if res_en:
            # Filter table region
            h = img.height
            table_items = [it for it in res_en if 0.15 * h < (it[0][0][1] + it[0][2][1])/2 < 0.92 * h]
            
            # Find % signs which uniquely define each stock row
            pct_items = sorted([it for it in table_items if "%" in it[1] or "％" in it[1]], 
                               key=lambda x: (x[0][0][1] + x[0][2][1])/2)
            
            for pct in pct_items:
                y_pct = (pct[0][0][1] + pct[0][2][1])/2
                # Row elements within row window
                row_elements = [it for it in table_items if (y_pct - 50) <= (it[0][0][1] + it[0][2][1])/2 <= (y_pct + 120)]
                row_elements.sort(key=lambda x: x[0][0][0])
                
                # Extract ticker or stock code
                # US ticker: 2-6 uppercase letters (excluding US/KR/MTS tags)
                tickers = [it[1].strip() for it in row_elements if re.match(r'^[A-Z]{2,6}$', it[1].strip()) and it[1].strip() not in ['US', 'KR', 'ETF']]
                # Korean stock code: 6 digits
                kr_codes = [it[1].strip() for it in row_elements if re.match(r'^\d{6}$', it[1].strip())]
                
                ticker = tickers[0] if tickers else (kr_codes[0] if kr_codes else "")
                
                # Numbers with decimals: price, change $
                decimals = [it[1].strip() for it in row_elements if re.match(r'^\d[\d,]*\.\d+$', it[1].strip())]
                price = decimals[0] if len(decimals) >= 1 else ""
                chg_dollar = decimals[1] if len(decimals) >= 2 else ""
                
                # Volume: large integer without decimal point
                ints = [it[1].strip() for it in row_elements if re.search(r'[\d,]{4,}', it[1].strip()) and '.' not in it[1]]
                vol = ints[0] if ints else ""
                
                pct_str = pct[1].strip()
                
                if ticker or price:
                    stocks.append({
                        "조건식": cond_title,
                        "종목명": ticker, # Default to ticker if name is English
                        "티커": ticker,
                        "현재가": price,
                        "등락률": pct_str,
                        "전일대비": chg_dollar,
                        "거래량": vol
                    })
    except Exception:
        pass
        
    return cond_title, stocks

def generate_excel_bytes(data_dict, base_date):
    wb = openpyxl.Workbook()
    wb.remove(wb.active) # Remove default sheet
    
    navy_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    white_bold = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
    title_font = Font(name="맑은 고딕", size=13, bold=True, color="1F4E79")
    sub_font = Font(name="맑은 고딕", size=9, italic=True, color="595959")
    border_thin = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )
    center_align = Alignment(horizontal="center", vertical="center")
    
    # 1. Summary Sheet
    ws_sum = wb.create_sheet(title="조건식_목록_요약")
    ws_sum.views.sheetView[0].showGridLines = True
    ws_sum.append(["조건식 데이터 요약"])
    ws_sum.append([f"기준일자: {base_date} | 총 {len(data_dict)}개 조건식"])
    ws_sum.append([])
    ws_sum.append(["No", "조건식 제목", "추출 종목 수"])
    ws_sum.cell(row=1, column=1).font = title_font
    ws_sum.cell(row=2, column=1).font = sub_font
    for c in range(1, 4):
        ws_sum.cell(row=4, column=c).fill = navy_fill
        ws_sum.cell(row=4, column=c).font = white_bold
        ws_sum.cell(row=4, column=c).alignment = center_align
        
    for idx, (c_name, c_stocks) in enumerate(data_dict.items(), start=1):
        ws_sum.append([idx, c_name, len(c_stocks)])
        for c in range(1, 4):
            ws_sum.cell(row=ws_sum.max_row, column=c).border = border_thin
            
    # 2. Individual Condition Sheets
    for c_name, c_stocks in data_dict.items():
        # Clean sheet title (max 31 chars, no invalid chars)
        sheet_title = re.sub(r'[\\/*?:\[\]]', '_', c_name)[:30]
        ws = wb.create_sheet(title=sheet_title)
        ws.views.sheetView[0].showGridLines = True
        
        ws.append([c_name])
        ws.append([f"기준일자: {base_date} | 추출 종목: 총 {len(c_stocks)}개"])
        ws.append([])
        headers = ["No", "종목명", "티커", "현재가", "등락률", "전일대비", "거래량"]
        ws.append(headers)
        
        ws.cell(row=1, column=1).font = title_font
        ws.cell(row=2, column=1).font = sub_font
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=4, column=col_idx)
            cell.fill = navy_fill
            cell.font = white_bold
            cell.alignment = center_align
            
        for s_idx, s in enumerate(c_stocks, start=1):
            row_vals = [
                s_idx,
                s.get("종목명", ""),
                s.get("티커", ""),
                s.get("현재가", ""),
                s.get("등락률", ""),
                s.get("전일대비", ""),
                s.get("거래량", "")
            ]
            ws.append(row_vals)
            for col_idx in range(1, len(headers) + 1):
                ws.cell(row=ws.max_row, column=col_idx).border = border_thin
                
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
            
    st.info(f"총 {len(all_images)}장의 사진을 확인했습니다. 텍스트 추출 및 정밀 분석을 시작합니다...")
    
    # Progress Bar & Status Text
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    # Real Dynamic OCR Processing
    extracted_by_condition = {}
    detected_date = datetime.date.today().strftime("%Y-%m-%d")
    
    for i, file in enumerate(all_images):
        status_text.text(f"🔍 사진 분석 중 ({i+1}/{len(all_images)}): {file.name}")
        image = Image.open(file)
        if i == 0:
            detected_date = extract_exif_date(image)
            
        c_title, stocks = parse_single_image(file)
        
        if c_title not in extracted_by_condition:
            extracted_by_condition[c_title] = []
            
        # De-duplicate stocks within condition formula (scroll overlap removal)
        existing_keys = {s["티커"] for s in extracted_by_condition[c_title] if s.get("티커")}
        for s in stocks:
            t = s.get("티커")
            if t and t in existing_keys:
                continue # Skip scroll duplicate
            if t:
                existing_keys.add(t)
            extracted_by_condition[c_title].append(s)
            
        progress_bar.progress((i + 1) / len(all_images))
        
    status_text.text(f"분석 및 중복 제거 완료! 🎉 (총 {len(all_images)}장 완료)")
    st.success(f"총 {len(all_images)}장의 사진에서 중복 종목을 제거하고 분석을 완료했습니다!")

    # Master DataFrame for consolidated CSV and in-app viewing
    all_stock_rows = []
    for c_title, stock_list in extracted_by_condition.items():
        for s in stock_list:
            all_stock_rows.append(s)

    master_df = pd.DataFrame(all_stock_rows) if all_stock_rows else pd.DataFrame()
    condition_names = list(extracted_by_condition.keys())
    
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
            options=condition_names if condition_names else ["없음"],
            label_visibility="collapsed"
        )
        if selected_cond and selected_cond in extracted_by_condition:
            c_stocks = extracted_by_condition[selected_cond]
            st.info(f"📌 **{selected_cond}** (추출된 종목: {len(c_stocks)}개)")
            if c_stocks:
                df_cond = pd.DataFrame(c_stocks)
                # drop condition column inside its own tab
                if "조건식" in df_cond.columns:
                    df_cond = df_cond.drop(columns=["조건식"])
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
        summary_list = [{"No": i+1, "조건식 제목": c, "추출 종목 수": len(s)} for i, (c, s) in enumerate(extracted_by_condition.items())]
        if summary_list:
            st.dataframe(pd.DataFrame(summary_list), use_container_width=True, hide_index=True)

    st.write("---")

    # --- DOWNLOAD SECTION ---
    st.subheader("💾 파일 다운로드")
    
    # Generate actual binary file contents
    excel_bytes = generate_excel_bytes(extracted_by_condition, detected_date)
    csv_bytes = master_df.to_csv(index=False).encode('utf-8-sig') if not master_df.empty else b""
    
    b64_excel = base64.b64encode(excel_bytes).decode()
    b64_csv = base64.b64encode(csv_bytes).decode() if csv_bytes else ""

    dcol1, dcol2 = st.columns(2)
    with dcol1:
        st.markdown(f'''
            <a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_excel}" 
               download="MTS_조건검색_추출_{detected_date}.xlsx" 
               target="_blank" 
               class="pwa-download-btn">
                📥 조건식별 엑셀(.xlsx) 새창 다운로드
            </a>
        ''', unsafe_allow_html=True)
        st.download_button(
            label="📥 기본 엑셀 다운로드 (사파리용)",
            data=excel_bytes,
            file_name=f"MTS_조건검색_추출_{detected_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with dcol2:
        if b64_csv:
            st.markdown(f'''
                <a href="data:text/csv;charset=utf-8;base64,{b64_csv}" 
                   download="MTS_조건검색_전체_{detected_date}.csv" 
                   target="_blank" 
                   class="pwa-download-btn">
                    📄 전체 통합 CSV(.csv) 새창 다운로드
                </a>
            ''', unsafe_allow_html=True)
        st.download_button(
            label="📄 기본 CSV 다운로드 (사파리용)",
            data=csv_bytes,
            file_name=f"MTS_조건검색_전체_{detected_date}.csv",
            mime="text/csv",
            use_container_width=True
        )

    # iOS User Experience Guide
    st.markdown("""
    <div class="guide-box">
        <b>💡 아이폰 연속 업무 꿀팁</b><br>
        • <b>추천 1 (앱 내 확인)</b>: 다운로드를 누르지 않아도 위 <b>'1. 조건식별 상세 시트'</b>와 <b>'2. 전체 통합 종목 표'</b> 탭에서 두 파일의 모든 내용을 바로 비교·검색하실 수 있습니다.<br>
        • <b>추천 2 (사파리 탭 사용)</b>: '홈 화면 아이콘' 대신 <b>일반 사파리 브라우저 탭(mina-stock.streamlit.app)</b>으로 열어두시면, 상단에 항상 <b>'뒤로가기(<)'</b>가 유지되어 파일 다운로드 후에도 1초 만에 복귀할 수 있습니다.<br>
        • <b>새창 다운로드 버튼</b>: 홈 화면 앱에서 갇히는 현상을 방지하기 위해 사파리 새 창을 여는 버튼을 상단에 함께 제공합니다.
    </div>
    """, unsafe_allow_html=True)

    # Bottom Action Bar: Reset / Start New Task
    st.write("")
    if st.button("◀ 닫기 / 새 작업 시작 (처음으로)", key="btn_bottom_reset", type="primary", use_container_width=True):
        st.session_state["uploader_key"] += 1
        st.rerun()

else:
    st.info("👆 위 사진 선택 영역을 눌러 MTS 조건검색 캡처본(또는 ZIP 파일)을 올려주세요.")
