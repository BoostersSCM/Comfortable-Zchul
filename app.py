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
        # 사용자가 "헤더는 1행"이라고 하셨으므로 header=0 사용
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

    # 컬럼명 공백 제거 (오류 방지)
    df.columns = [str(col).strip() for col in df.columns]

    # 2. 필요한 컬럼 매핑
    column_mapping = {
        '거래처': '업체', '발주번호': '발주번호', '품번': '품번', '품명': '품명',
        '단가': '납품단가', '납품수량': '납품수량', '금액': '납품금액(세전)',
        '부가세': '부가세', '금액계': '납품금액(세후)'
    }

    # 파일에 필수 컬럼이 있는지 확인
    valid_cols = [col for col in column_mapping.keys() if col in df.columns]
    if not valid_cols:
        return None, f"파일 첫 줄에 필요한 제목(거래처, 발주번호 등)이 없습니다. (감지된 제목: {list(df.columns)})"

    # 데이터 추출 및 컬럼명 변경
    df_extracted = df[valid_cols].copy()
    df_extracted.rename(columns=column_mapping, inplace=True)

    # 3. 숫자 데이터 변환 (집계를 위해 필수, 콤마 제거)
    numeric_cols = ['납품수량', '납품금액(세전)', '부가세', '납품금액(세후)']
    for col in numeric_cols:
        if col in df_extracted.columns:
            df_extracted[col] = pd.to_numeric(
                df_extracted[col].astype(str).str.replace(',', ''), errors='coerce'
            ).fillna(0)

    # 4. 집계 (GroupBy) - 업체, 발주번호, 품번, 품명 기준
    group_keys = ['업체', '발주번호', '품번', '품명']
    # 실제 데이터에 존재하는 키만 사용
    real_keys = [k for k in group_keys if k in df_extracted.columns]
    
    if not real_keys:
        return None, "그룹화할 기준 컬럼(업체, 발주번호 등)이 없습니다."

    # 합계 계산
    df_grouped = df_extracted.groupby(real_keys, as_index=False)[['납품수량', '납품금액(세전)', '부가세', '납품금액(세후)']].sum()

    # 5. 단가 재계산 (합계 금액 / 합계 수량)
    if '납품금액(세전)' in df_grouped.columns and '납품수량' in df_grouped.columns:
        df_grouped['납품단가'] = df_grouped.apply(
            lambda x: x['납품금액(세전)'] / x['납품수량'] if x['납품수량'] != 0 else 0, axis=1
        )

    # 6. [요청사항] 거래처(업체) 순으로 정렬
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
    """엑셀 파일 생성 및 수식/서식 적용"""
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
                # 엑셀 수식: 잔여금액 = 세후금액 - 선금
                ws[f"{col_balance}{r}"] = f"={col_total}{r}-{col_prepay}{r}"
                
                # 천단위 콤마 서식 적용
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
st.markdown("ERP 파일을 업로드하고 **[변환하기]** 버튼을 누르면, 업체별/발주번호별로 집계된 결과를 확인할 수 있습니다.")

# 1. 파일 업로드
uploaded_file = st.file_uploader("파일 업로드 (xlsx, csv)", type=['xlsx', 'csv', 'xls'])

# 세션 상태 초기화 (버튼 클릭 후에도 데이터 유지용)
if 'processed_data' not in st.session_state:
    st.session_state.processed_data = None

# 2. 변환하기 버튼
if uploaded_file:
    # 파일을 새로 올렸을 때 기존 데이터 초기화 (선택사항)
    # st.session_state.processed_data = None 

    if st.button("🚀 변환 및 집계 실행", type="primary"):
        with st.spinner("데이터 분석 및 집계 중..."):
            df_result, error_msg = load_and_aggregate_data(uploaded_file)
            
            if df_result is not None:
                st.session_state.processed_data = df_result
                st.success("작업이 완료되었습니다! 아래 결과를 확인하세요.")
            else:
                st.error(f"오류 발생: {error_msg}")

# 3. 결과 표시 및 다운로드 (데이터가 있을 때만 표시)
if st.session_state.processed_data is not None:
    st.divider()
    st.subheader("📋 결과 미리보기 (업체순 정렬됨)")
    
    # 데이터프레임 표시 (숫자 포맷팅)
    st.dataframe(
        st.session_state.processed_data, 
        use_container_width=True,
        column_config={
            "납품단가": st.column_config.NumberColumn(format="%d"),
            "납품수량": st.column_config.NumberColumn(format="%d"),
            "납품금액(세전)": st.column_config.NumberColumn(format="%d"),
            "납품금액(세후)": st.column_config.NumberColumn(format="%d"),
            "선금 금액": st.column_config.NumberColumn(format="%d"),
            "잔여금액": st.column_config.NumberColumn(format="%d"),
        }
    )
    
    # 엑셀 생성
    excel_data = create_excel_with_formula(st.session_state.processed_data)
    
    # 다운로드 버튼
    st.download_button(
        label="📥 엑셀 파일 다운로드",
        data=excel_data,
        file_name="납품대금_집계표_업체별.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
