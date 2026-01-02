import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# 페이지 기본 설정
st.set_page_config(page_title="납품대금 정리 프로그램 (Final)", layout="wide")

def find_header_row_index(df):
    """
    데이터프레임에서 실제 헤더(컬럼명)가 위치한 행 번호를 찾습니다.
    '발주번호'와 '품명'이라는 단어가 동시에 있는 행을 헤더로 간주합니다.
    """
    # 엑셀 파일 상위 30줄을 검사해서 진짜 헤더 위치를 찾음
    scan_limit = min(30, len(df))
    for i in range(scan_limit):
        row_values = [str(val).strip() for val in df.iloc[i].values]
        has_order_no = any('발주번호' in val for val in row_values)
        has_item_name = any('품명' in val for val in row_values)
        if has_order_no and has_item_name:
            return i
    return None

def get_cleaned_dataframe(uploaded_file):
    """업로드된 파일을 읽어 정제된 DataFrame을 반환합니다."""
    try:
        # 1. 파일 읽기 (헤더 없이 전체 읽기)
        if uploaded_file.name.endswith('.csv'):
            try:
                df_raw = pd.read_csv(uploaded_file, header=None)
            except UnicodeDecodeError:
                uploaded_file.seek(0)
                df_raw = pd.read_csv(uploaded_file, header=None, encoding='cp949')
        else:
            df_raw = pd.read_excel(uploaded_file, header=None, engine='openpyxl')
            
    except Exception as e:
        st.error(f"❌ 파일을 읽을 수 없습니다: {e}")
        return None

    # 2. 진짜 헤더 위치 찾기
    header_idx = find_header_row_index(df_raw)
    
    if header_idx is None:
        st.error("❌ 데이터 양식을 인식할 수 없습니다. 파일 안에 '발주번호'와 '품명' 열이 있는지 확인해주세요.")
        # 디버깅용: 파일 앞부분을 보여줌
        st.caption("▼ 파일의 앞부분 데이터입니다 (헤더를 못 찾는 원인 확인용):")
        st.dataframe(df_raw.head(5))
        return None

    # 3. 데이터 재구성 (헤더 적용)
    df = df_raw.iloc[header_idx + 1:].copy()
    df.columns = df_raw.iloc[header_idx].values
    
    # 컬럼명 앞뒤 공백 제거 (필수)
    df.columns = [str(col).strip() for col in df.columns]

    # 4. 컬럼 매핑
    column_mapping = {
        '거래처': '업체', '발주번호': '발주번호', '품번': '품번', '품명': '품명',
        '단가': '납품단가', '납품수량': '납품수량', '금액': '납품금액(세전)',
        '부가세': '부가세', '금액계': '납품금액(세후)'
    }

    valid_cols = [col for col in column_mapping.keys() if col in df.columns]
    
    if not valid_cols:
        st.error(f"❌ 필요한 컬럼이 발견되지 않았습니다. 감지된 컬럼: {list(df.columns)}")
        return None

    df_result = df[valid_cols].copy()
    df_result.rename(columns=column_mapping, inplace=True)
    
    # 5. 숫자 변환 (콤마 제거)
    numeric_cols = ['납품단가', '납품수량', '납품금액(세전)', '부가세', '납품금액(세후)']
    for col in numeric_cols:
        if col in df_result.columns:
            df_result[col] = pd.to_numeric(
                df_result[col].astype(str).str.replace(',', ''), errors='coerce'
            ).fillna(0)

    # 6. 추가 관리 컬럼
    df_result['선금 지급일'] = ''
    df_result['선금 금액'] = 0
    df_result['잔여금액'] = 0 
    
    return df_result

def create_excel_file(df):
    """DataFrame을 엑셀 파일(BytesIO)로 변환하고 수식을 적용합니다."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    
    output.seek(0)
    wb = load_workbook(output)
    ws = wb.active
    
    header_map = {str(cell.value).strip(): cell.col_idx for cell in ws[1]}
    
    try:
        col_total = get_column_letter(header_map['납품금액(세후)'])
        col_prepay = get_column_letter(header_map['선금 금액'])
        col_balance = get_column_letter(header_map['잔여금액'])
        
        row_count = ws.max_row
        for r in range(2, row_count + 1):
            # 수식 적용: 잔여금액 = 세후금액 - 선금
            ws[f"{col_balance}{r}"] = f"={col_total}{r}-{col_prepay}{r}"
            
            # 서식 적용
            ws[f"{col_total}{r}"].number_format = '#,##0'
            ws[f"{col_prepay}{r}"].number_format = '#,##0'
            ws[f"{col_balance}{r}"].number_format = '#,##0'
            
            for col_name in ['납품단가', '납품금액(세전)', '부가세']:
                if col_name in header_map:
                    ws[f"{get_column_letter(header_map[col_name])}{r}"].number_format = '#,##0'
                    
    except KeyError:
        pass

    final_output = BytesIO()
    wb.save(final_output)
    final_output.seek(0)
    return final_output

# --- UI 실행 ---
st.title("📊 납품대금 자동 정리기")
st.info("💡 엑셀 상단에 제목이나 결재란이 있어도 자동으로 건너뛰고 변환합니다.")

uploaded_file = st.file_uploader("파일 업로드 (xlsx, csv)", type=['xlsx', 'csv', 'xls'])

if uploaded_file:
    with st.spinner("파일을 정밀 분석 중입니다..."):
        df_clean = get_cleaned_dataframe(uploaded_file)
        
    if df_clean is not None:
        st.success("✅ 분석 완료! 아래 미리보기를 확인하세요.")
        
        # 1. 미리보기 표 출력
        st.markdown("### 📋 변환 결과 미리보기")
        st.dataframe(
            df_clean, 
            use_container_width=True,
            column_config={
                "납품단가": st.column_config.NumberColumn(format="%d"),
                "납품수량": st.column_config.NumberColumn(format="%d"),
                "납품금액(세전)": st.column_config.NumberColumn(format="%d"),
                "부가세": st.column_config.NumberColumn(format="%d"),
                "납품금액(세후)": st.column_config.NumberColumn(format="%d"),
                "선금 금액": st.column_config.NumberColumn(format="%d"),
                "잔여금액": st.column_config.NumberColumn(format="%d"),
            }
        )
        st.caption("※ '잔여금액'은 이곳에서는 0으로 보이지만, 엑셀 파일에는 자동 계산 수식이 적용되어 있습니다.")

        # 2. 다운로드 버튼
        excel_data = create_excel_file(df_clean)
        
        st.download_button(
            label="📥 엑셀 파일 다운로드",
            data=excel_data,
            file_name="납품대금_정리_완료.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
