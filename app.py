import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# 페이지 설정
st.set_page_config(page_title="납품대금 집계 프로그램", layout="wide")

def load_and_aggregate_data(uploaded_file):
    """
    데이터를 읽고(헤더 1행 고정) '발주번호/품번' 기준으로 집계 후
    거래처 순으로 정렬합니다.
    """
    try:
        # 1. 파일 읽기 (헤더는 무조건 첫 번째 줄(0번 행)로 고정)
        if uploaded_file.name.endswith('.csv'):
            try:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, header=0)
            except:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, header=0, encoding='cp949')
        else:
            uploaded_file.seek(0)
            df = pd.read_excel(uploaded_file, header=0, engine='openpyxl')
    except Exception as e:
        return None, f"파일을 읽을 수 없습니다: {e}"

    # 컬럼명 공백 제거
    df.columns = [str(col).strip() for col in df.columns]

    # 2. 필요한 컬럼 매핑
    column_mapping = {
        '거래처': '업체', '발주번호': '발주번호', '품번': '품번', '품명': '품명',
        '단가': '납품단가', '납품수량': '납품수량', '금액': '납품금액(세전)',
        '부가세': '부가세', '금액계': '납품금액(세후)'
    }

    # 필수 컬럼 확인
    valid_cols = [col for col in column_mapping.keys() if col in df.columns]
    if not valid_cols:
        return None, f"파일 첫 줄에 필요한 제목(거래처, 발주번호 등)이 없습니다. (감지된 제목: {list(df.columns)})"

    # 데이터 추출 및 컬럼명 변경
    df_extracted = df[valid_cols].copy()
    df_extracted.rename(columns=column_mapping, inplace=True)

    # 3. 숫자 데이터 변환 (콤마 제거 후 숫자로)
    numeric_cols = ['납품수량', '납품금액(세전)', '부가세', '납품금액(세후)']
    for col in numeric_cols:
        if col in df_extracted.columns:
            df_extracted[col] = pd.to_numeric(
                df_extracted[col].astype(str).str.replace(',', ''), errors='coerce'
            ).fillna(0)

    # 4. 집계 (GroupBy) - 업체, 발주번호, 품번, 품명 기준
    group_keys = ['업체', '발주번호', '품번', '품명']
    real_keys = [k for k in group_keys if k in df_extracted.columns]
    
    if not real_keys:
        return None, "그룹화할 기준 컬럼이 없습니다."

    # 합계 계산
    df_grouped = df_extracted.groupby(real_keys, as_index=False)[['납품수량', '납품금액(세전)', '부가세', '납품금액(세후)']].sum()

    # 5. 단가 재계산
    if '납품금액(세전)' in df_grouped.columns and '납품수량' in df_grouped.columns:
        df_grouped['납품단가'] = df_grouped.apply(
            lambda x: x['납품금액(세전)'] / x['납품수량'] if x['납품수량'] != 0 else 0, axis=1
        )

    # 6. 거래처(업체) 순으로 정렬
    if '업체' in df_grouped.columns:
        df_grouped = df_grouped.sort_values(by=['업체', '발주번호', '품번'])

    # 7. 컬럼 순서 정리
    desired_order = ['업체', '발주번호', '품번', '품명', '납품단가', '납품수량', '납품금액(세전)', '부가세', '납품금액(세후)']
    final_cols = [c for c in desired_order if c in df_grouped.columns]
    df_final = df_grouped[final_cols]

    # 8. 추가 관리 컬럼 생성
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
                ws[f"{col_balance}{r}"] = f"={col_total}{r}-{col_prepay}{r}"
                
                # 천단위 콤마 서식
                cols_to_format = ['납품단가', '납품금액(세전)', '부가세', '납품금액(세후)', '선금 금액', '잔여금액']
                for col_name in cols_to_format:
                    if col_name in header_map:
                         ws[f"{get_column_letter(header_map[col_name])}{r}"].number_format = '#,##0'
    except Exception:
        pass

    final_output = BytesIO()
    wb.save(final_output)
    final_output.seek(0)
    return final_output

# --- 메인 화면 UI ---
st.title("📊 납품대금 집계 프로그램")
st.markdown("ERP 파일을 업로드하고 **[변환 및 집계 실행]** 버튼을 누르면 자동 집계된 결과를 보여줍니다.")

# 1. 파일 업로드
uploaded_file = st.file_uploader("파일 업로드 (xlsx, csv)", type=['xlsx', 'csv', 'xls'])

# 세션 상태 초기화
if 'processed_data' not in st.session_state:
    st.session_state.processed_data = None

# 2. 실행 버튼
if uploaded_file:
    # 버튼을 눌러야만 실행됨
    if st.button("🚀 변환 및 집계 실행", type="primary"):
        with st.spinner("데이터 분석 중..."):
            df_result, error_msg = load_and_aggregate_data(uploaded_file)
            
            if df_result is not None:
                st.session_state.processed_data = df_result
                st.success("완료되었습니다!")
            else:
                st.error(f"오류 발생: {error_msg}")

# 3. 결과 표시 및 다운로드
if st.session_state.processed_data is not None:
    st.divider()
    st.subheader("📋 결과 미리보기")
    
    # [수정됨] Pandas Style을 사용하여 천 단위 콤마 포맷팅 적용
    format_dict = {
        "납품단가": "{:,.0f}",
        "납품수량": "{:,.0f}",
        "납품금액(세전)": "{:,.0f}",
        "부가세": "{:,.0f}",
        "납품금액(세후)": "{:,.0f}",
        "선금 금액": "{:,.0f}",
        "잔여금액": "{:,.0f}",
    }
    
    # 데이터프레임에 실제 존재하는 컬럼만 포맷 적용
    valid_format = {k: v for k, v in format_dict.items() if k in st.session_state.processed_data.columns}
    
    st.dataframe(
        st.session_state.processed_data.style.format(valid_format), 
        use_container_width=True
    )
    
    # 다운로드 버튼
    excel_data = create_excel_with_formula(st.session_state.processed_data)
    
    st.download_button(
        label="📥 엑셀 파일 다운로드",
        data=excel_data,
        file_name="납품대금_집계표.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
