import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# 페이지 기본 설정
st.set_page_config(page_title="납품대금 정리 프로그램", layout="wide")

def find_header_index(df):
    """
    데이터프레임을 순회하며 실제 헤더(컬럼명)가 있는 행의 인덱스를 찾습니다.
    핵심 키워드들이 포함된 행을 헤더로 간주합니다.
    """
    # 이 단어들이 포함된 행을 찾으면 헤더로 인식함
    required_keywords = ['발주번호', '품명', '거래처', '금액', '단가', '수량']
    
    # 상위 20줄까지만 탐색 (속도 최적화)
    search_limit = min(20, len(df))
    
    for i in range(search_limit):
        row_values = df.iloc[i].astype(str).values
        # 행의 값 중 키워드가 2개 이상 포함되어 있으면 해당 행을 헤더로 본다
        match_count = sum(1 for keyword in required_keywords if any(keyword in val for val in row_values))
        
        if match_count >= 2:
            return i
            
    return None

def process_excel(uploaded_file):
    """업로드된 파일을 처리하여 엑셀 바이너리 데이터를 반환하는 함수"""
    
    # 1. 일단 헤더 없이 전체를 읽어옵니다.
    try:
        if uploaded_file.name.endswith('.csv'):
            # CSV는 인코딩 문제가 있을 수 있어 utf-8과 cp949 둘 다 시도
            try:
                df_raw = pd.read_csv(uploaded_file, header=None)
            except UnicodeDecodeError:
                uploaded_file.seek(0)
                df_raw = pd.read_csv(uploaded_file, header=None, encoding='cp949')
        else:
            df_raw = pd.read_excel(uploaded_file, header=None)
            
    except Exception as e:
        st.error(f"파일을 읽는 도중 오류가 발생했습니다: {e}")
        return None

    # 2. 진짜 헤더 위치 찾기 (스마트 스캔)
    header_idx = find_header_index(df_raw)
    
    if header_idx is None:
        st.error("데이터에서 '발주번호', '품명' 같은 핵심 컬럼을 찾을 수 없습니다. ERP 파일 양식을 확인해주세요.")
        return None
        
    # 3. 찾은 위치를 헤더로 설정하여 데이터 재구성
    # 헤더 행을 컬럼으로 설정하고, 그 이후의 데이터만 사용
    df = df_raw.iloc[header_idx+1:].copy()
    df.columns = df_raw.iloc[header_idx].values
    
    # 컬럼명 앞뒤 공백 제거 (매우 중요)
    df.columns = [str(col).strip() for col in df.columns]

    # 4. 필요한 컬럼 매핑
    column_mapping = {
        '거래처': '업체',
        '발주번호': '발주번호',
        '품번': '품번',
        '품명': '품명',
        '단가': '납품단가',
        '납품수량': '납품수량',
        '금액': '납품금액(세전)',
        '부가세': '부가세',
        '금액계': '납품금액(세후)'
    }

    # 파일에 존재하는 컬럼만 매핑
    valid_columns = [col for col in column_mapping.keys() if col in df.columns]
    
    if not valid_columns:
        st.error(f"필요한 컬럼이 하나도 없습니다. 감지된 컬럼명: {list(df.columns)}")
        return None

    # 데이터 추출 및 컬럼명 변경
    df_selected = df[valid_columns].copy()
    df_selected.rename(columns=column_mapping, inplace=True)
    
    # 데이터가 비어있는지 확인
    if df_selected.empty:
        st.warning("추출된 데이터가 없습니다.")
        return None

    # 5. 숫자 데이터 정제 (콤마 제거 및 숫자 변환)
    numeric_cols = ['납품단가', '납품수량', '납품금액(세전)', '부가세', '납품금액(세후)']
    for col in numeric_cols:
        if col in df_selected.columns:
            # 문자열로 된 숫자가 있을 경우 콤마 제거 후 숫자로 변환
            df_selected[col] = pd.to_numeric(df_selected[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)

    # 6. 추가 관리 컬럼 생성
    df_selected['선금 지급일'] = ''
    df_selected['선금 금액'] = 0
    df_selected['잔여금액'] = 0

    # 7. 엑셀 파일 생성
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_selected.to_excel(writer, index=False, sheet_name='Sheet1')
    
    # 8. 엑셀 수식 및 서식 적용
    output.seek(0)
    wb = load_workbook(output)
    ws = wb.active

    row_count = ws.max_row
    header = {cell.value: cell.col_idx for cell in ws[1]}
    
    try:
        # 컬럼 위치 찾기 (없으면 안전하게 패스하도록 처리)
        col_total = get_column_letter(header.get('납품금액(세후)')) if '납품금액(세후)' in header else None
        col_prepay = get_column_letter(header.get('선금 금액')) if '선금 금액' in header else None
        col_balance = get_column_letter(header.get('잔여금액')) if '잔여금액' in header else None
        
        # 수식 적용이 가능한 경우에만 실행
        if col_total and col_prepay and col_balance:
            for row in range(2, row_count + 1):
                # 수식: 잔여금액 = 납품금액(세후) - 선금 금액
                ws[f"{col_balance}{row}"] = f"={col_total}{row}-{col_prepay}{row}"
                
                # 서식 적용
                ws[f"{col_total}{row}"].number_format = '#,##0'
                ws[f"{col_prepay}{row}"].number_format = '#,##0'
                ws[f"{col_balance}{row}"].number_format = '#,##0'
                
                # 기타 숫자 컬럼 서식
                for key in ['납품단가', '납품금액(세전)', '부가세']:
                    if key in header:
                        col_letter = get_column_letter(header[key])
                        ws[f"{col_letter}{row}"].number_format = '#,##0'

    except Exception as e:
        st.warning(f"엑셀 수식 적용 중 일부 오류가 있었으나 파일은 생성되었습니다: {e}")

    final_output = BytesIO()
    wb.save(final_output)
    final_output.seek(0)
    
    return final_output

# --- UI 구성 ---
st.title("📊 납품대금 자동 정리기 (ver 2.0)")
st.markdown("""
**사용 방법:**
1. ERP에서 다운받은 엑셀 파일을 그대로 업로드하세요. (상단에 결재란이 있어도 자동으로 처리합니다)
2. 변환된 파일을 다운로드 받으세요.
3. **'선금 금액'** 칸에 숫자를 입력하면 **잔여금액**이 자동으로 계산됩니다.
""")

uploaded_file = st.file_uploader("엑셀(.xlsx) 또는 CSV 파일을 업로드하세요", type=['xlsx', 'csv', 'xls'])

if uploaded_file is not None:
    with st.spinner('파일 구조를 분석하고 변환 중입니다...'):
        processed_data = process_excel(uploaded_file)
        
    if processed_data:
        st.success('✅ 변환이 완료되었습니다!')
        st.download_button(
            label="📥 결과 엑셀 파일 다운로드",
            data=processed_data,
            file_name="납품대금_관리대장_result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
