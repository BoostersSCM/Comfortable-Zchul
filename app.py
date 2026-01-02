import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import requests
import urllib.parse

# --- 설정 ---
st.set_page_config(page_title="Boosters 납품대금 관리", layout="wide")

# Secrets에서 정보 가져오기
try:
    CLIENT_ID = st.secrets["google_auth"]["client_id"]
    CLIENT_SECRET = st.secrets["google_auth"]["client_secret"]
    REDIRECT_URI = st.secrets["google_auth"]["redirect_uri"]
except FileNotFoundError:
    st.error("Secrets 파일이 없습니다. .streamlit/secrets.toml을 확인해주세요.")
    st.stop()

# --- 인증 관련 함수 ---
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
    """인증 코드로 토큰 교환"""
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
    """토큰으로 유저 정보 조회"""
    user_info_url = "https://www.googleapis.com/oauth2/v1/userinfo"
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get(user_info_url, headers=headers)
    return response.json()

# --- 메인 앱 로직 (엑셀 변환) ---
def process_excel(uploaded_file):
    # (기존 코드와 동일합니다)
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")
        return None

    column_mapping = {
        '거래처': '업체', '발주번호': '발주번호', '품번': '품번', '품명': '품명',
        '단가': '납품단가', '납품수량': '납품수량', '금액': '납품금액(세전)',
        '부가세': '부가세', '금액계': '납품금액(세후)'
    }

    available_cols = [col for col in column_mapping.keys() if col in df.columns]
    if not available_cols:
        st.error("ERP 파일 형식이 맞지 않습니다. (필수 컬럼 없음)")
        return None
        
    df_selected = df[available_cols].copy()
    df_selected.rename(columns=column_mapping, inplace=True)
    df_selected['선금 지급일'] = ''
    df_selected['선금 금액'] = 0
    df_selected['잔여금액'] = 0

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_selected.to_excel(writer, index=False, sheet_name='Sheet1')
    
    output.seek(0)
    wb = load_workbook(output)
    ws = wb.active
    row_count = ws.max_row
    header = {cell.value: cell.col_idx for cell in ws[1]}
    
    try:
        col_total = get_column_letter(header['납품금액(세후)'])
        col_prepay = get_column_letter(header['선금 금액'])
        col_balance = get_column_letter(header['잔여금액'])
        
        for row in range(2, row_count + 1):
            ws[f"{col_balance}{row}"] = f"={col_total}{row}-{col_prepay}{row}"
            ws[f"{col_total}{row}"].number_format = '#,##0'
            ws[f"{col_prepay}{row}"].number_format = '#,##0'
            ws[f"{col_balance}{row}"].number_format = '#,##0'
            ws[f"{get_column_letter(header['납품단가'])}{row}"].number_format = '#,##0'
            ws[f"{get_column_letter(header['납품금액(세전)'])}{row}"].number_format = '#,##0'
            ws[f"{get_column_letter(header['부가세'])}{row}"].number_format = '#,##0'
    except KeyError:
        pass

    final_output = BytesIO()
    wb.save(final_output)
    final_output.seek(0)
    return final_output

def main_app():
    st.title("📊 Boosters 납품대금 자동 정리기")
    
    # 로그인 정보 표시
    user_email = st.session_state.get('user_email', '')
    st.sidebar.success(f"로그인됨: {user_email}")
    if st.sidebar.button("로그아웃"):
        st.session_state.clear()
        st.rerun()

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
            st.download_button(
                label="📥 변환된 엑셀 파일 다운로드",
                data=processed_data,
                file_name="납품대금_관리대장_result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.info("다운로드 받은 파일을 열어서 '선금 금액'만 입력하면 잔여금액이 자동 계산됩니다.")

# --- 실행 흐름 제어 (로그인 체크) ---

# 1. 이미 로그인 된 상태인지 확인
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False

# 2. URL에 인증 코드(code)가 있는지 확인 (구글 로그인 후 리다이렉트 되었을 때)
if not st.session_state['logged_in']:
    query_params = st.query_params
    if "code" in query_params:
        code = query_params["code"]
        try:
            token_response = get_token_from_code(code)
            if "access_token" in token_response:
                user_info = get_user_info(token_response["access_token"])
                email = user_info.get("email", "")
                
                # [중요] 이메일 도메인 체크
                if email.endswith("@boosters.kr"):
                    st.session_state['logged_in'] = True
                    st.session_state['user_email'] = email
                    # URL 정리 (code 파라미터 제거)
                    st.query_params.clear()
                    st.rerun()
                else:
                    st.error(f"접속 권한이 없습니다. ({email}) \n @boosters.kr 계정만 사용할 수 있습니다.")
            else:
                st.error("로그인 실패: 토큰을 받아오지 못했습니다.")
        except Exception as e:
            st.error(f"로그인 처리 중 오류 발생: {e}")

# 3. 화면 표시 분기
if st.session_state['logged_in']:
    main_app()
else:
    st.title("🔒 Boosters 내부 시스템")
    st.warning("관계자 외 접근을 금지합니다.")
    
    login_url = get_login_url()
    st.markdown(f'''
        <a href="{login_url}" target="_self">
            <button style="
                background-color: #4285F4; 
                color: white; 
                padding: 10px 20px; 
                border: none; 
                border-radius: 5px; 
                cursor: pointer; 
                font-size: 16px; 
                font-weight: bold;">
                Google 계정으로 로그인 (Boosters)
            </button>
        </a>
    ''', unsafe_allow_html=True)
