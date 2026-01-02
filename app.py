import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="납품대금 집계 프로그램 (피벗 모드)", layout="wide")

def find_header_row_index_robust(df):
    """헤더(제목) 위치를 자동으로 찾는 함수"""
    # 1. '발주번호' 텍스트가 있는 셀 찾기
    scan_limit = min(50, len(df))
    for i in range(scan_limit):
        row_values = [str(val).strip() for val in df.iloc[i].values]
        if any("발주번호" in v for v in row_values):
            return i
            
    # 2. 실패 시 키워드 매칭
    keywords = ['품명', '품번', '거래처', '단가', '수량', '금액']
    best_idx = 0
    max_matches = 0
    for i in range(scan_limit):
        row_values = [str(val).strip() for val in df.iloc[i].values]
        matches = sum(1 for k in keywords if any(k in v for v in row_values))
        if matches > max_matches:
            max_matches = matches
            best_idx = i
    return best_idx

def load_and_aggregate_data(uploaded_file, header_row_idx):
    """데이터를 읽고 '발주번호/품번' 기준으로 집계(Sum)합니다."""
    try:
        # 파일 읽기
        if uploaded_file.name.endswith('.csv'):
            try:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, header=header_row_idx)
            except:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, header=header_row_idx, encoding='cp949')
        else:
            uploaded_file.seek(0)
            df = pd.read_excel(uploaded_file, header=header_row_idx, engine='openpyxl')
    except Exception as e:
        return None, f"파일 읽기 오류: {e}"

    # 컬럼명 공백 제거
    df.columns = [str(col).strip() for col in df.columns]

    # 필요한 컬럼 매핑
    column_mapping = {
        '거래처': '업체', '발주번호': '발주번호', '품번': '품번', '품명': '품명',
        '단가': '납품단가', '납품수량': '납품수량', '금액': '납품금액(세전)',
        '부가세': '부가세', '금액계': '납품금액(세후)'
    }

    valid_cols = [col for col in column_mapping.keys() if col in df.columns]
    if not valid_cols:
        return None, f"필수 컬럼을 찾을 수 없습니다. (감지된 컬럼: {list(df.columns)})"

    # 1. 데이터 추출
    df_extracted = df[valid_cols].copy()
    df_extracted.rename(columns=column_mapping, inplace=True)

    # 2. 숫자 데이터 변환 (집계를 위해 필수)
    numeric_cols = ['납품수량', '납품금액(세전)', '부가세', '납품금액(세후)']
    for col in numeric_cols:
        if col in df_extracted.columns:
            df_extracted[col] = pd.to_numeric(
                df_extracted[col].astype(str).str.replace(',', ''), errors='coerce'
            ).fillna(0)

    # 3. [핵심] 발주번호/품번 기준으로 집계 (Pivot 역할)
    # 업체와 품명은 그룹핑 키에 포함 (보통 동일하므로)
    group_keys = ['업체', '발주번호', '품번', '품명']
    # 실제 데이터에 존재하는 키만 사용
    real_keys = [k for k in group_keys if k in df_extracted.columns]
    
    if not real_keys:
        return None, "그룹화할 기준 컬럼(발주번호, 품번 등)이 없습니다."

    # 집계 수행 (수량, 금액은 합계)
    df_grouped = df_extracted.groupby(real_keys, as_index=False)[['납품수량', '납품금액(세전)', '부가세', '납품금액(세후)']].sum()

    # 4. 단가 재계산 (총 금액 / 총 수량) - 정확성을 위해
    if '납품금액(세전)' in df_grouped.columns and '납품수량' in df_grouped.columns:
        df_grouped['납품단가'] = df_grouped.apply(
            lambda x: x['납품금액(세전)'] / x['납품수량'] if x['납품수량'] != 0 else 0, axis=1
        )
    
    # 5. 컬럼 순서 정리
    desired_order = ['업체', '발주번호', '품번', '품명', '납품단가', '납품수량', '납품금액(세전)', '부가세', '납품금액(세후)']
    final_cols = [c for c in desired_order if c in df_grouped.columns]
    df_final = df_grouped[final_cols]

    # 6. 추가 관리 컬럼 생성
    df_final['선금 지급일'] = ''
    df_final['선금 금액'] = 0
    df_final['잔여금액'] = 0 
    
    return df_final, None

def create_excel_with_formula(df):
    """엑셀 파일 생성 및 수식 적용"""
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
                # 수식: 잔여금액 = 세후금액 - 선금
                ws[f"{col_balance}{r}"] = f"={col_total}{r}-{col_prepay}{r}"
                
                # 서식 적용
                for col_name in ['납품단가', '납품금액(세전)', '부가세', '납품금액(세후)', '선금 금액', '잔여금액']:
                    if col_name in header_map:
                         ws[f"{get_column_letter(header_map[col_name])}{r}"].number_format = '#,##0'
    except Exception:
        pass

    final_output = BytesIO()
    wb.save(final_output)
    final_output.seek(0)
    return final_output

# --- 메인 화면 ---
st.title("📊 납품대금 집계 프로그램 (피벗 모드)")
st.markdown("""
업로드한 내역을 **발주번호와 품번별로 자동으로 합쳐서(Sum)** 보여줍니다.  
(같은 품목이 여러 번 납품되었어도 한 줄로 요약됩니다.)
""")

uploaded_file = st.file_uploader("파일 업로드 (xlsx, csv)", type=['xlsx', 'csv', 'xls'])

if uploaded_file:
    # 1. 헤더 위치 자동 감지
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

    # 2. 헤더 위치 수동 보정 (필요시)
    st.write("---")
    col1, col2 = st.columns([1, 2])
    with col1:
        header_row = st.number_input(
            "📌 헤더(제목) 행 번호 확인", 
            min_value=0, 
            max_value=30, 
            value=detected_header_idx,
            help="표의 제목(발주번호 등)이 시작되는 행 번호입니다. 결과가 이상하면 조절하세요."
        )

    # 3. 집계 실행 및 결과 표시
    df_result, error_msg = load_and_aggregate_data(uploaded_file, header_row)
    
    if df_result is not None:
        st.success(f"✅ 집계 완료! 총 {len(df_result)}건으로 요약되었습니다.")
        
        # 미리보기
        st.dataframe(
            df_result, 
            use_container_width=True,
            column_config={
                "납품단가": st.column_config.NumberColumn(format="%d"),
                "납품수량": st.column_config.NumberColumn(format="%d"),
                "납품금액(세전)": st.column_config.NumberColumn(format="%d"),
                "납품금액(세후)": st.column_config.NumberColumn(format="%d")
            }
        )
        
        # 다운로드
        excel_data = create_excel_with_formula(df_result)
        st.download_button(
            label="📥 집계된 엑셀 파일 다운로드",
            data=excel_data,
            file_name="납품대금_집계표.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
    else:
        st.warning("⚠️ 데이터를 변환할 수 없습니다.")
        if error_msg:
            st.error(error_msg)
        st.info("💡 위 슬라이더의 숫자를 변경하여 헤더 위치를 맞춰보세요.")
        st.write("▼ 원본 파일 데이터 (참고용)")
        st.dataframe(df_raw.head(10))
