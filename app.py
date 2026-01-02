import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# 페이지 기본 설정
st.set_page_config(page_title="납품대금 정리 프로그램 (수동 보정)", layout="wide")

def find_header_row_index_robust(df):
    """
    1. '발주번호'라는 정확한 단어가 포함된 셀이 있는지 전체 탐색
    2. 없으면 가장 키워드가 많은 행을 추측
    """
    # 1단계: '발주번호' 텍스트가 있는 셀의 행 번호 찾기 (가장 확실)
    for i in range(min(50, len(df))): # 상위 50줄 탐색
        row_values = [str(val).strip() for val in df.iloc[i].values]
        if any("발주번호" in v for v in row_values):
            return i
            
    # 2단계: 실패 시 키워드 매칭 (기존 방식)
    keywords = ['품명', '품번', '거래처', '단가', '수량', '금액']
    best_idx = 0
    max_matches = 0
    for i in range(min(20, len(df))):
        row_values = [str(val).strip() for val in df.iloc[i].values]
        matches = sum(1 for k in keywords if any(k in v for v in row_values))
        if matches > max_matches:
            max_matches = matches
            best_idx = i
            
    return best_idx

def load_data(uploaded_file, header_row_idx):
    """지정된 행(header_row_idx)을 헤더로 사용하여 데이터를 읽습니다."""
    try:
        if uploaded_file.name.endswith('.csv'):
            try:
                # 헤더 위치를 지정해서 읽기
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, header=header_row_idx)
            except UnicodeDecodeError:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, header=header_row_idx, encoding='cp949')
        else:
            uploaded_file.seek(0)
            df = pd.read_excel(uploaded_file, header=header_row_idx, engine='openpyxl')
            
        return df
    except Exception as e:
        return None

def process_dataframe(df):
    # 컬럼명 공백 제거
    df.columns = [str(col).strip() for col in df.columns]

    # 컬럼 매핑
    column_mapping = {
        '거래처': '업체', '발주번호': '발주번호', '품번': '품번', '품명': '품명',
        '단가': '납품단가', '납품수량': '납품수량', '금액': '납품금액(세전)',
        '부가세': '부가세', '금액계': '납품금액(세후)'
    }

    # 존재하는 컬럼만 선택
    valid_cols = [col for col in column_mapping.keys() if col in df.columns]
    
    # 데이터 추출
    if not valid_cols:
        return None, list(df.columns) # 실패 시 현재 컬럼명 반환

    df_result = df[valid_cols].copy()
    df_result.rename(columns=column_mapping, inplace=True)
    
    # 숫자 변환
    numeric_cols = ['납품단가', '납품수량', '납품금액(세전)', '부가세', '납품금액(세후)']
    for col in numeric_cols:
        if col in df_result.columns:
            df_result[col] = pd.to_numeric(
                df_result[col].astype(str).str.replace(',', ''), errors='coerce'
            ).fillna(0)

    # 추가 컬럼
    df_result['선금 지급일'] = ''
    df_result['선금 금액'] = 0
    df_result['잔여금액'] = 0 
    
    return df_result, None

def create_excel_file(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    
    output.seek(0)
    wb = load_workbook(output)
    ws = wb.active
    
    header_map = {str(cell.value).strip(): cell.col_idx for cell in ws[1]}
    
    try:
        if '납품금액(세후)' in header_map and '선금 금액' in header_map and '잔여금액' in header_map:
            col_total = get_column_letter(header_map['납품금액(세후)'])
            col_prepay = get_column_letter(header_map['선금 금액'])
            col_balance = get_column_letter(header_map['잔여금액'])
            
            row_count = ws.max_row
            for r in range(2, row_count + 1):
                ws[f"{col_balance}{r}"] = f"={col_total}{r}-{col_prepay}{r}"
                
                for col_name in ['납품단가', '납품금액(세전)', '부가세', '납품금액(세후)', '선금 금액', '잔여금액']:
                    if col_name in header_map:
                         ws[f"{get_column_letter(header_map[col_name])}{r}"].number_format = '#,##0'
    except Exception:
        pass

    final_output = BytesIO()
    wb.save(final_output)
    final_output.seek(0)
    return final_output

# --- UI 실행 ---
st.title("📊 납품대금 자동 정리기")

st.info("💡 파일 업로드 후, 미리보기가 이상하면 아래 **'헤더 위치 직접 지정'**을 조절해주세요.")

uploaded_file = st.file_uploader("파일 업로드 (xlsx, csv)", type=['xlsx', 'csv', 'xls'])

if uploaded_file:
    # 1. 일단 헤더 없이 읽어서 자동 감지 시도
    try:
        if uploaded_file.name.endswith('.csv'):
            try:
                df_raw = pd.read_csv(uploaded_file, header=None)
            except:
                uploaded_file.seek(0)
                df_raw = pd.read_csv(uploaded_file, header=None, encoding='cp949')
        else:
            df_raw = pd.read_excel(uploaded_file, header=None, engine='openpyxl')
            
        detected_header_idx = find_header_row_index_robust(df_raw)
        
    except Exception as e:
        st.error(f"파일 기본 읽기 실패: {e}")
        st.stop()

    # 2. 사용자 보정 컨트롤 (슬라이더)
    st.write("---")
    col1, col2 = st.columns([1, 2])
    with col1:
        header_row = st.number_input(
            "📌 헤더(제목) 행 번호 직접 지정 (0부터 시작)", 
            min_value=0, 
            max_value=30, 
            value=detected_header_idx,
            help="표의 제목(발주번호, 품명 등)이 있는 행 번호를 맞춰주세요."
        )
    
    with col2:
        st.caption(f"현재 **{header_row}행**을 제목으로 인식하고 변환을 시도합니다.")

    # 3. 선택된 헤더로 데이터 로드 및 변환
    df_loaded = load_data(uploaded_file, header_row)
    
    if df_loaded is not None:
        df_clean, error_cols = process_dataframe(df_loaded)
        
        if df_clean is not None:
            st.success("✅ 변환 성공!")
            
            st.dataframe(df_clean.head(10), use_container_width=True)
            
            excel_data = create_excel_file(df_clean)
            st.download_button(
                label="📥 엑셀 파일 다운로드",
                data=excel_data,
                file_name="납품대금_정리_완료.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
        else:
            st.warning(f"⚠️ **{header_row}행**을 제목으로 읽었으나 필요한 컬럼이 없습니다.")
            st.error(f"감지된 컬럼명: {error_cols}")
            st.info("👆 위 슬라이더 숫자를 1씩 변경해보세요. 제목 줄이 맞아야 합니다.")
            
            # 디버깅용 원본 데이터 표시
            st.write("▼ 원본 파일 데이터 (참고용)")
            st.dataframe(df_raw.head(10))
