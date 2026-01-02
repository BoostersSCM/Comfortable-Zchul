import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# 페이지 기본 설정
st.set_page_config(page_title="납품대금 정리 프로그램 (Final)", layout="wide")

def find_header_row_index(df):
    """
    데이터프레임에서 가장 '헤더(컬럼명)'다운 행을 찾습니다.
    특정 키워드가 가장 많이 포함된 행을 헤더로 선정합니다.
    """
    # 헤더에 등장할 것으로 예상되는 단어들
    keywords = ['발주번호', '품명', '품번', '거래처', '단가', '수량', '금액', '부가세', '공급가', '업체']
    
    best_idx = -1
    max_matches = 0
    
    # 상위 20줄 검사
    scan_limit = min(20, len(df))
    
    for i in range(scan_limit):
        # 해당 행의 모든 값을 문자열로 합침 (공백 제거)
        row_values = [str(val).strip() for val in df.iloc[i].values]
        row_str = " ".join(row_values)
        
        # 키워드가 몇 개나 포함되어 있는지 카운트
        matches = 0
        for k in keywords:
            if k in row_str:
                matches += 1
        
        # 가장 많이 매칭된 행을 기억
        if matches > max_matches:
            max_matches = matches
            best_idx = i
            
    # 매칭된 키워드가 2개 이상이면 그 행을 헤더로 인정
    if max_matches >= 2:
        return best_idx
        
    return None

def get_cleaned_dataframe(uploaded_file):
    try:
        if uploaded_file.name.endswith('.csv'):
            try:
                df_raw = pd.read_csv(uploaded_file, header=None)
            except UnicodeDecodeError:
                uploaded_file.seek(0)
                df_raw = pd.read_csv(uploaded_file, header=None, encoding='cp949')
        else:
            df_raw = pd.read_excel(uploaded_file, header=None, engine='openpyxl')
            
    except Exception as e:
        st.error(f"❌ 파일을 읽는 중 오류가 발생했습니다: {e}")
        return None

    # 1. 헤더 위치 찾기 (개선된 로직)
    header_idx = find_header_row_index(df_raw)
    
    if header_idx is None:
        st.error("❌ 표의 머리글(Header)을 찾을 수 없습니다.")
        st.warning("엑셀 파일 안에 '발주번호', '품명', '금액' 같은 단어가 포함되어 있는지 확인해주세요.")
        st.caption("▼ 업로드된 파일의 앞부분 데이터:")
        st.dataframe(df_raw.head(5))
        return None

    # 2. 데이터 재구성
    df = df_raw.iloc[header_idx + 1:].copy()
    df.columns = df_raw.iloc[header_idx].values
    
    # 컬럼명 정리 (문자열 변환 및 공백 제거)
    df.columns = [str(col).strip() for col in df.columns]

    # 3. 컬럼 매핑
    column_mapping = {
        '거래처': '업체', '발주번호': '발주번호', '품번': '품번', '품명': '품명',
        '단가': '납품단가', '납품수량': '납품수량', '금액': '납품금액(세전)',
        '부가세': '부가세', '금액계': '납품금액(세후)'
    }

    # 파일에 존재하는 컬럼만 선택
    valid_cols = [col for col in column_mapping.keys() if col in df.columns]
    
    # 만약 '발주번호' 같은 핵심 컬럼이 없더라도, 있는 것만이라도 추출하도록 유연하게 처리
    if not valid_cols:
        st.error(f"❌ 매칭되는 컬럼이 하나도 없습니다. (감지된 컬럼: {list(df.columns)})")
        return None

    df_result = df[valid_cols].copy()
    df_result.rename(columns=column_mapping, inplace=True)
    
    # 누락된 필수 컬럼이 있다면 빈 값으로라도 생성 (에러 방지)
    expected_cols = list(column_mapping.values())
    for col in expected_cols:
        if col not in df_result.columns:
            df_result[col] = 0 if '금액' in col or '수량' in col else ''

    # 순서 재배치 (원하는 순서대로)
    final_order = ['업체', '발주번호', '품번', '품명', '납품단가', '납품수량', '납품금액(세전)', '부가세', '납품금액(세후)']
    # 실제 존재하는 컬럼만 선택해서 순서 맞춤
    final_order = [c for c in final_order if c in df_result.columns]
    df_result = df_result[final_order]

    # 4. 숫자 데이터 변환
    numeric_cols = ['납품단가', '납품수량', '납품금액(세전)', '부가세', '납품금액(세후)']
    for col in numeric_cols:
        if col in df_result.columns:
            df_result[col] = pd.to_numeric(
                df_result[col].astype(str).str.replace(',', ''), errors='coerce'
            ).fillna(0)

    # 5. 추가 관리 컬럼
    df_result['선금 지급일'] = ''
    df_result['선금 금액'] = 0
    df_result['잔여금액'] = 0 
    
    return df_result

def create_excel_file(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    
    output.seek(0)
    wb = load_workbook(output)
    ws = wb.active
    
    header_map = {str(cell.value).strip(): cell.col_idx for cell in ws[1]}
    
    try:
        col_total = get_column_letter(header_map.get('납품금액(세후)', 1))
        col_prepay = get_column_letter(header_map.get('선금 금액', 1))
        col_balance = get_column_letter(header_map.get('잔여금액', 1))
        
        # 필요한 컬럼이 다 있을 때만 수식 적용
        if '납품금액(세후)' in header_map and '선금 금액' in header_map and '잔여금액' in header_map:
            row_count = ws.max_row
            for r in range(2, row_count + 1):
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

# --- UI 실행 ---
st.title("📊 납품대금 자동 정리기")
st.info("💡 파일의 1행, 2행 어디에 헤더가 있든 자동으로 찾아냅니다.")

uploaded_file = st.file_uploader("파일 업로드 (xlsx, csv)", type=['xlsx', 'csv', 'xls'])

if uploaded_file:
    with st.spinner("파일 분석 중..."):
        df_clean = get_cleaned_dataframe(uploaded_file)
        
    if df_clean is not None:
        st.success("✅ 분석 완료! 아래 미리보기를 확인하세요.")
        
        # 1. 미리보기
        st.markdown("### 📋 변환 결과 미리보기")
        st.dataframe(
            df_clean, 
            use_container_width=True
        )
        
        # 2. 다운로드
        excel_data = create_excel_file(df_clean)
        
        st.download_button(
            label="📥 엑셀 파일 다운로드",
            data=excel_data,
            file_name="납품대금_정리_완료.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
