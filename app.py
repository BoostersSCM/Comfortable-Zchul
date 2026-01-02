import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# 페이지 기본 설정
st.set_page_config(page_title="납품대금 정리 프로그램", layout="wide")

def process_excel(uploaded_file):
    """업로드된 파일을 처리하여 엑셀 바이너리 데이터를 반환하는 함수"""
    
    # 1. 파일 읽기 (CSV 또는 Excel 구분)
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")
        return None

    # 2. 필요한 컬럼 매핑
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

    # 컬럼 존재 여부 확인 및 추출
    available_cols = [col for col in column_mapping.keys() if col in df.columns]
    if not available_cols:
        st.error("ERP 파일 형식이 맞지 않습니다. (필수 컬럼 없음)")
        return None
        
    df_selected = df[available_cols].copy()
    df_selected.rename(columns=column_mapping, inplace=True)

    # 3. 추가 관리 컬럼 생성
    df_selected['선금 지급일'] = ''
    df_selected['선금 금액'] = 0
    df_selected['잔여금액'] = 0

    # 4. 메모리 상에서 엑셀 파일 생성 (BytesIO 사용)
    output = BytesIO()
    
    # Pandas로 1차 저장
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_selected.to_excel(writer, index=False, sheet_name='Sheet1')
    
    # 5. Openpyxl로 다시 불러와서 수식 및 서식 적용
    output.seek(0) # 커서를 처음으로 이동
    wb = load_workbook(output)
    ws = wb.active

    row_count = ws.max_row
    
    # 헤더 위치 찾기
    header = {cell.value: cell.col_idx for cell in ws[1]}
    
    # 필요한 열의 알파벳 위치 찾기 (없을 경우 대비해 try-except)
    try:
        col_total = get_column_letter(header['납품금액(세후)'])
        col_prepay = get_column_letter(header['선금 금액'])
        col_balance = get_column_letter(header['잔여금액'])
        
        # 수식 및 서식 적용 Loop
        for row in range(2, row_count + 1):
            # 수식: 잔여금액 = 납품금액(세후) - 선금 금액
            ws[f"{col_balance}{row}"] = f"={col_total}{row}-{col_prepay}{row}"
            
            # 천단위 콤마 서식
            ws[f"{col_total}{row}"].number_format = '#,##0'
            ws[f"{col_prepay}{row}"].number_format = '#,##0'
            ws[f"{col_balance}{row}"].number_format = '#,##0'
            ws[f"{get_column_letter(header['납품단가'])}{row}"].number_format = '#,##0'
            ws[f"{get_column_letter(header['납품금액(세전)'])}{row}"].number_format = '#,##0'
            ws[f"{get_column_letter(header['부가세'])}{row}"].number_format = '#,##0'

    except KeyError as e:
        st.warning(f"일부 컬럼을 찾을 수 없어 수식 적용에 실패했습니다: {e}")

    # 최종 결과를 바이너리로 저장
    final_output = BytesIO()
    wb.save(final_output)
    final_output.seek(0)
    
    return final_output

# --- Streamlit UI 구성 ---
st.title("📊 ERP 납품대금 자동 정리기")
st.markdown("""
ERP에서 다운로드 받은 엑셀/CSV 파일을 업로드하면,  
**업체/발주/금액** 등을 정리하고 **선금 관리 수식**을 자동으로 넣어줍니다.
""")

uploaded_file = st.file_uploader("여기에 파일을 드래그하거나 클릭하여 업로드하세요", type=['xlsx', 'csv', 'xls'])

if uploaded_file is not None:
    with st.spinner('파일을 변환하고 있습니다...'):
        processed_data = process_excel(uploaded_file)
        
    if processed_data:
        st.success('변환이 완료되었습니다!')
        
        # 다운로드 버튼
        st.download_button(
            label="📥 변환된 엑셀 파일 다운로드",
            data=processed_data,
            file_name="납품대금_관리대장_result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.info("다운로드 받은 파일을 열어서 '선금 금액'만 입력하면 잔여금액이 자동 계산됩니다.")
