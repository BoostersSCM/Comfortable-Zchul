import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import requests
import urllib.parse

# 페이지 설정
st.set_page_config(page_title="Boosters 납품대금 집계 시스템", layout="wide")

# --- 1. 인증(Auth) 관련 설정 및 함수 ---

# Secrets 가져오기 (예외처리 포함)
try:
    CLIENT_ID = st.secrets["google_auth"]["client_id"]
    CLIENT_SECRET = st.secrets["google_auth"]["client_secret"]
    REDIRECT_URI = st.secrets["google_auth"]["redirect_uri"]
except Exception:
    st.error("⚠️ Secrets 설정이 되어있지 않습니다. Streamlit Cloud의 Settings > Secrets를 확인해주세요.")
    st.stop()
    
query_params = st.query_params

if "error" in query_params:
    st.error(f"OAuth error: {query_params.get('error')}")
    st.write(query_params)
    st.stop()
    
def get_login_url():
    """구글 로그인 URL 생성"""
    base_url = "https://accounts.google.com/o/oauth2/v2/auth"
    params = {
        "response_type": "code",
        "client_id": CLIENT_ID,
        "redirect_uri": REDIRECT_URI,
        "scope": "openid email profile",
        "access_type": "offline",
        "prompt": "consent"
    }
    return f"{base_url}?{urllib.parse.urlencode(params)}"

def get_token_from_code(code):
    """인증 코드로 액세스 토큰 교환"""
    token_url = "https://oauth2.googleapis.com/token"
    data = {
        "code": code,
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "redirect_uri": REDIRECT_URI,
        "grant_type": "authorization_code"
    }
    response = requests.post(token_url, data=data)
    return response.json()

def get_user_info(access_token):
    """액세스 토큰으로 유저 정보(이메일 등) 조회"""
    user_info_url = "https://www.googleapis.com/oauth2/v1/userinfo"
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get(user_info_url, headers=headers)
    return response.json()

# --- 2. 데이터 처리(ERP) 관련 함수 ---

def load_and_aggregate_data(uploaded_file):
    try:
        # 헤더 1행(index 0) 고정 읽기
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
        return None, f"파일 읽기 실패: {e}"

    df.columns = [str(col).strip() for col in df.columns]

    column_mapping = {
        '거래처': '업체', '발주번호': '발주번호', '품번': '품번', '품명': '품명',
        '단가': '납품단가', '납품수량': '납품수량', '금액': '납품금액(세전)',
        '부가세': '부가세', '금액계': '납품금액(세후)'
    }

    valid_cols = [col for col in column_mapping.keys() if col in df.columns]
    if not valid_cols:
        return None, f"필수 컬럼 없음. 감지된 제목: {list(df.columns)}"

    df_extracted = df[valid_cols].copy()
    df_extracted.rename(columns=column_mapping, inplace=True)

    # 숫자 변환
    numeric_cols = ['납품수량', '납품금액(세전)', '부가세', '납품금액(세후)']
    for col in numeric_cols:
        if col in df_extracted.columns:
            df_extracted[col] = pd.to_numeric(
                df_extracted[col].astype(str).str.replace(',', ''), errors='coerce'
            ).fillna(0)

    # 집계 (GroupBy)
    group_keys = ['업체', '발주번호', '품번', '품명']
    real_keys = [k for k in group_keys if k in df_extracted.columns]
    
    if not real_keys:
        return None, "그룹화 기준 컬럼 부족"

    df_grouped = df_extracted.groupby(real_keys, as_index=False)[['납품수량', '납품금액(세전)', '부가세', '납품금액(세후)']].sum()

    # 단가 재계산
    if '납품금액(세전)' in df_grouped.columns and '납품수량' in df_grouped.columns:
        df_grouped['납품단가'] = df_grouped.apply(
            lambda x: x['납품금액(세전)'] / x['납품수량'] if x['납품수량'] != 0 else 0, axis=1
        )

    # 정렬 (업체명 순)
    if '업체' in df_grouped.columns:
        df_grouped = df_grouped.sort_values(by=['업체', '발주번호', '품번'])

    # 컬럼 순서 및 추가
    desired_order = ['업체', '발주번호', '품번', '품명', '납품단가', '납품수량', '납품금액(세전)', '부가세', '납품금액(세후)']
    final_cols = [c for c in desired_order if c in df_grouped.columns]
    df_final = df_grouped[final_cols]

    df_final['선금 지급일'] = ''
    df_final['선금 금액'] = 0
    df_final['잔여금액'] = 0 
    
    return df_final, None

def create_excel_with_formula(df):
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

# --- 3. 메인 애플리케이션 화면 (로그인 성공 시 보임) ---

def main_app():
    # 사이드바: 로그인 정보 및 로그아웃
    with st.sidebar:
        st.success(f"접속자: {st.session_state.user_email}")
        if st.button("로그아웃"):
            st.session_state.clear()
            st.rerun()

    # 메인 컨텐츠
    st.title("📊 납품대금 집계 프로그램")
    st.markdown("""
    ERP 파일을 업로드하고 **[변환 및 집계 실행]**을 누르면  
    **업체별/발주번호별**로 자동 집계하여 정리해줍니다.
    """)

    uploaded_file = st.file_uploader("파일 업로드 (xlsx, csv)", type=['xlsx', 'csv', 'xls'])

    if 'processed_data' not in st.session_state:
        st.session_state.processed_data = None

    if uploaded_file:
        if st.button("🚀 변환 및 집계 실행", type="primary"):
            with st.spinner("데이터 분석 중..."):
                df_result, error_msg = load_and_aggregate_data(uploaded_file)
                if df_result is not None:
                    st.session_state.processed_data = df_result
                    st.success("집계 완료!")
                else:
                    st.error(f"오류: {error_msg}")

    if st.session_state.processed_data is not None:
        st.divider()
        st.subheader("📋 결과 미리보기")
        
        # 1,000 단위 콤마 포맷팅
        format_dict = {
            "납품단가": "{:,.0f}",
            "납품수량": "{:,.0f}",
            "납품금액(세전)": "{:,.0f}",
            "부가세": "{:,.0f}",
            "납품금액(세후)": "{:,.0f}",
            "선금 금액": "{:,.0f}",
            "잔여금액": "{:,.0f}",
        }
        valid_format = {k: v for k, v in format_dict.items() if k in st.session_state.processed_data.columns}
        
        st.dataframe(
            st.session_state.processed_data.style.format(valid_format), 
            use_container_width=True
        )
        
        excel_data = create_excel_with_formula(st.session_state.processed_data)
        st.download_button(
            label="📥 엑셀 파일 다운로드",
            data=excel_data,
            file_name="납품대금_집계표.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# --- 4. 실행 흐름 제어 (로그인 체크) ---

# 세션 상태 초기화
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False
if 'user_email' not in st.session_state:
    st.session_state['user_email'] = ''

# 로그인 상태가 아니면 로그인 로직 수행
if not st.session_state['logged_in']:
    # URL에 인증 코드(code)가 있는지 확인
    query_params = st.query_params
    
    if "code" in query_params:
        code = query_params["code"]
        try:
            token_res = get_token_from_code(code)
            if "access_token" in token_res:
                user_info = get_user_info(token_res["access_token"])
                email = user_info.get("email", "")
                
                # 도메인 체크 (@boosters.kr)
                if email.endswith("@boosters.kr"):
                    st.session_state['logged_in'] = True
                    st.session_state['user_email'] = email
                    st.query_params.clear() # URL 파라미터 정리
                    st.rerun()
                else:
                    st.error(f"접속 권한이 없습니다. ({email}) @boosters.kr 계정만 가능합니다.")
            else:
                st.error("로그인 실패: 토큰 오류")
        except Exception as e:
            st.error(f"로그인 처리 중 오류: {e}")
    
    # 로그인 화면 표시
    else:
        st.title("🔒 Boosters Internal Tool")
        st.write("관계자 외 접근을 금지합니다.")
        
        login_url = get_login_url()
        st.markdown(f'''
            <a href="{login_url}" target="_self">
                <button style="
                    background-color: #4285F4; color: white; padding: 12px 24px; 
                    border: none; border-radius: 4px; cursor: pointer; 
                    font-size: 16px; font-weight: bold;">
                    G Suite 계정으로 로그인 (Boosters)
                </button>
            </a>
        ''', unsafe_allow_html=True)

# 로그인 상태면 메인 앱 실행
else:
    main_app()
