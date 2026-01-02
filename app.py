import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import requests
import urllib.parse

# =========================================================
# 0. 페이지 설정
# =========================================================
st.set_page_config(
    page_title="Boosters 납품대금 집계 시스템",
    layout="wide"
)

# =========================================================
# 1. Google OAuth 설정
# =========================================================
try:
    CLIENT_ID = st.secrets["google_auth"]["client_id"]
    CLIENT_SECRET = st.secrets["google_auth"]["client_secret"]
    REDIRECT_URI = st.secrets["google_auth"]["redirect_uri"]
except Exception:
    st.error("⚠️ Streamlit Secrets에 google_auth 설정이 없습니다.")
    st.stop()


def get_login_url():
    base_url = "https://accounts.google.com/o/oauth2/v2/auth"
    params = {
        "response_type": "code",
        "client_id": CLIENT_ID,
        "redirect_uri": REDIRECT_URI,
        "scope": "openid email profile",
        "access_type": "offline",
        "prompt": "consent",
        "hd": "boosters.kr",  # 도메인 힌트
    }
    return f"{base_url}?{urllib.parse.urlencode(params)}"


def get_token_from_code(code):
    token_url = "https://oauth2.googleapis.com/token"
    data = {
        "code": code,
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "redirect_uri": REDIRECT_URI,
        "grant_type": "authorization_code",
    }
    return requests.post(token_url, data=data).json()


def get_user_info(access_token):
    userinfo_url = "https://openidconnect.googleapis.com/v1/userinfo"
    headers = {"Authorization": f"Bearer {access_token}"}
    return requests.get(userinfo_url, headers=headers).json()


# =========================================================
# 2. ERP 데이터 처리
# =========================================================
def load_and_aggregate_data(uploaded_file):
    try:
        if uploaded_file.name.endswith(".csv"):
            try:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file)
            except:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, encoding="cp949")
        else:
            uploaded_file.seek(0)
            df = pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception as e:
        return None, f"파일 읽기 실패: {e}"

    df.columns = [str(c).strip() for c in df.columns]

    column_mapping = {
        "거래처": "업체",
        "발주번호": "발주번호",
        "품번": "품번",
        "품명": "품명",
        "단가": "납품단가",
        "납품수량": "납품수량",
        "금액": "납품금액(세전)",
        "부가세": "부가세",
        "금액계": "납품금액(세후)",
    }

    valid_cols = [c for c in column_mapping if c in df.columns]
    if not valid_cols:
        return None, f"필수 컬럼 없음: {list(df.columns)}"

    df = df[valid_cols].rename(columns=column_mapping)

    for col in ["납품수량", "납품금액(세전)", "부가세", "납품금액(세후)"]:
        if col in df.columns:
            df[col] = (
                df[col].astype(str)
                .str.replace(",", "")
                .astype(float)
                .fillna(0)
            )

    group_keys = ["업체", "발주번호", "품번", "품명"]
    df = (
        df.groupby(group_keys, as_index=False)[
            ["납품수량", "납품금액(세전)", "부가세", "납품금액(세후)"]
        ]
        .sum()
        .sort_values(["업체", "발주번호", "품번"])
    )

    df["납품단가"] = df.apply(
        lambda x: x["납품금액(세전)"] / x["납품수량"]
        if x["납품수량"] != 0
        else 0,
        axis=1,
    )

    df["선금 지급일"] = ""
    df["선금 금액"] = 0
    df["잔여금액"] = 0

    return df, None


def create_excel_with_formula(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)

    output.seek(0)
    wb = load_workbook(output)
    ws = wb.active

    header = {cell.value: cell.col_idx for cell in ws[1]}
    if {"납품금액(세후)", "선금 금액", "잔여금액"}.issubset(header):
        col_total = get_column_letter(header["납품금액(세후)"])
        col_prepay = get_column_letter(header["선금 금액"])
        col_balance = get_column_letter(header["잔여금액"])

        for r in range(2, ws.max_row + 1):
            ws[f"{col_balance}{r}"] = f"={col_total}{r}-{col_prepay}{r}"

    final = BytesIO()
    wb.save(final)
    final.seek(0)
    return final


# =========================================================
# 3. 메인 앱
# =========================================================
def main_app():
    with st.sidebar:
        st.success(f"접속자: {st.session_state.user_email}")
        if st.button("로그아웃"):
            st.session_state.clear()
            st.rerun()

    st.title("📊 납품대금 집계 프로그램")

    uploaded_file = st.file_uploader("ERP 파일 업로드", type=["xlsx", "xls", "csv"])

    if uploaded_file and st.button("🚀 변환 및 집계 실행", type="primary"):
        with st.spinner("처리 중..."):
            df, err = load_and_aggregate_data(uploaded_file)
            if err:
                st.error(err)
            else:
                st.session_state.df = df
                st.success("완료!")

    if "df" in st.session_state:
        st.dataframe(
            st.session_state.df.style.format("{:,.0f}"),
            use_container_width=True,
        )
        excel = create_excel_with_formula(st.session_state.df)
        st.download_button(
            "📥 엑셀 다운로드",
            excel,
            "납품대금_집계표.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# =========================================================
# 4. 로그인 흐름 + 🔥 OAuth 에러 디버그
# =========================================================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "user_email" not in st.session_state:
    st.session_state.user_email = ""

query_params = st.query_params

# 🔥 OAuth 에러 즉시 표시
if "error" in query_params:
    st.error("Google OAuth 에러 발생")
    st.write(query_params)
    st.stop()

if not st.session_state.logged_in:
    if "code" in query_params:
        code = query_params["code"]

        token_res = get_token_from_code(code)
        st.write("🔍 Token response", token_res)

        if "access_token" not in token_res:
            st.error("토큰 발급 실패")
            st.stop()

        user_info = get_user_info(token_res["access_token"])
        st.write("🔍 User info", user_info)

        email = user_info.get("email", "")
        if email.endswith("@boosters.kr"):
            st.session_state.logged_in = True
            st.session_state.user_email = email
            st.query_params.clear()
            st.rerun()
        else:
            st.error(f"접근 권한 없음: {email}")

    else:
        st.title("🔒 Boosters Internal Tool")
        st.markdown(
            f"""
            <a href="{get_login_url()}">
                <button style="
                    background:#4285F4;
                    color:white;
                    padding:12px 24px;
                    border:none;
                    border-radius:6px;
                    font-size:16px;
                    cursor:pointer;">
                    Google 계정으로 로그인
                </button>
            </a>
            """,
            unsafe_allow_html=True,
        )
else:
    main_app()
