# app.py
import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# =========================================================
# 0. 페이지 설정
# =========================================================
st.set_page_config(
    page_title="Boosters 납품대금 집계 시스템",
    layout="wide"
)

# =========================================================
# 1. 파일 읽기 (헤더 행 선택)
# =========================================================
def read_file_with_header(uploaded_file, header_row_excel_1based: int, header_row_csv_1based: int = 1):
    name = uploaded_file.name.lower()

    if name.endswith(".csv"):
        header_idx = max(header_row_csv_1based - 1, 0)
        try:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, header=header_idx)
        except Exception:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, header=header_idx, encoding="cp949")
        return df

    header_idx = max(header_row_excel_1based - 1, 0)
    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file, header=header_idx, engine="openpyxl")
    return df

# =========================================================
# 2. ERP 데이터 집계
# =========================================================
def load_and_aggregate_data(uploaded_file, header_row_excel_1based: int):
    try:
        df = read_file_with_header(uploaded_file, header_row_excel_1based)
    except Exception as e:
        return None, f"파일 읽기 실패: {e}"

    # 컬럼 정리
    df.columns = [str(col).strip() for col in df.columns]
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]

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

    # 필요한 컬럼 체크
    required_cols = ["발주번호", "품번", "납품수량", "납품금액(세전)", "부가세", "납품금액(세후)"]
    # 원본에서 실제로 존재하는 컬럼명(한글)을 먼저 매핑 가능한지 확인
    if not ("발주번호" in df.columns and "품번" in df.columns):
        return None, f"필수 컬럼 없음. 감지된 컬럼: {list(df.columns)}"

    valid_cols = [c for c in column_mapping if c in df.columns]
    df = df[valid_cols].rename(columns=column_mapping)

    # 숫자 변환
    numeric_cols = ["납품수량", "납품금액(세전)", "부가세", "납품금액(세후)"]
    for c in numeric_cols:
        if c in df.columns:
            df[c] = (
                df[c].astype(str)
                .str.replace(",", "")
                .pipe(pd.to_numeric, errors="coerce")
                .fillna(0)
            )

    # 문자열 정리(품번/발주번호/품명 공백 차이로 그룹 쪼개지는 것 방지)
    for c in ["발주번호", "품번", "품명", "업체"]:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()

    # ✅ 집계 키: 발주번호 + 품번 (요청사항 반영)
    group_keys = ["발주번호", "품번"]

    agg_dict = {
        "납품수량": "sum",
        "납품금액(세전)": "sum",
        "부가세": "sum",
        "납품금액(세후)": "sum",
    }

    df_grouped = df.groupby(group_keys, as_index=False).agg(agg_dict)

    # (옵션) 대표 정보(업체/품명)는 첫 값으로 붙임
    if "업체" in df.columns:
        vendor_first = df.groupby(group_keys, as_index=False)["업체"].first()
        df_grouped = df_grouped.merge(vendor_first, on=group_keys, how="left")

    if "품명" in df.columns:
        name_first = df.groupby(group_keys, as_index=False)["품명"].first()
        df_grouped = df_grouped.merge(name_first, on=group_keys, how="left")

    # 단가 재계산 (세전/수량 기준)
    df_grouped["납품단가"] = df_grouped.apply(
        lambda x: x["납품금액(세전)"] / x["납품수량"] if x["납품수량"] else 0,
        axis=1
    )

    # 컬럼 순서 정리
    final_cols = ["발주번호", "품번"]
    if "업체" in df_grouped.columns:
        final_cols = ["업체"] + final_cols
    if "품명" in df_grouped.columns:
        final_cols = final_cols + ["품명"]

    final_cols += ["납품단가", "납품수량", "납품금액(세전)", "부가세", "납품금액(세후)"]

    df_final = df_grouped[final_cols].copy()

    df_final["선금 지급일"] = ""
    df_final["선금 금액"] = 0
    df_final["잔여금액"] = 0

    # 보기 좋게 정렬
    df_final = df_final.sort_values(by=["발주번호", "품번"])

    return df_final, None

# =========================================================
# 3. 엑셀 생성 (잔여금액 수식 포함)
# =========================================================
def create_excel_with_formula(df: pd.DataFrame) -> BytesIO:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")

    output.seek(0)
    wb = load_workbook(output)
    ws = wb.active

    header_map = {str(cell.value): cell.col_idx for cell in ws[1]}

    if {"납품금액(세후)", "선금 금액", "잔여금액"}.issubset(header_map):
        col_total = get_column_letter(header_map["납품금액(세후)"])
        col_prepay = get_column_letter(header_map["선금 금액"])
        col_balance = get_column_letter(header_map["잔여금액"])

        for r in range(2, ws.max_row + 1):
            ws[f"{col_balance}{r}"] = f"={col_total}{r}-{col_prepay}{r}"

    final_output = BytesIO()
    wb.save(final_output)
    final_output.seek(0)
    return final_output

# =========================================================
# 4. 화면 표시용 DF (콤마 포맷)
# =========================================================
def make_display_df(df: pd.DataFrame) -> pd.DataFrame:
    df_disp = df.copy()
    num_cols = [
        "납품단가", "납품수량",
        "납품금액(세전)", "부가세", "납품금액(세후)",
        "선금 금액", "잔여금액"
    ]
    for c in num_cols:
        if c in df_disp.columns:
            df_disp[c] = (
                pd.to_numeric(df_disp[c], errors="coerce")
                .fillna(0)
                .astype(int)
                .map(lambda x: f"{x:,}")
            )
    return df_disp

# =========================================================
# 5. 메인 UI
# =========================================================
st.title("📊 납품대금 집계 프로그램")
st.markdown("""
ERP 파일을 업로드하고  
**엑셀 헤더 행을 지정한 뒤 [변환 및 집계 실행]**을 누르세요.
""")

uploaded_file = st.file_uploader(
    "파일 업로드 (xlsx, xls, csv)",
    type=["xlsx", "xls", "csv"]
)

header_row_excel = st.number_input(
    "엑셀 헤더 행 (1부터)",
    min_value=1,
    value=2,
    step=1
)

if "processed_data" not in st.session_state:
    st.session_state.processed_data = None

if uploaded_file:
    with st.expander("🔎 헤더 미리보기", expanded=True):
        try:
            preview = read_file_with_header(uploaded_file, header_row_excel)
            st.write("감지된 컬럼:", list(preview.columns))
            st.dataframe(preview.head(5), use_container_width=True)
        except Exception as e:
            st.error(f"미리보기 실패: {e}")

    if st.button("🚀 변환 및 집계 실행", type="primary"):
        with st.spinner("데이터 처리 중..."):
            df_result, error_msg = load_and_aggregate_data(uploaded_file, header_row_excel)
            if df_result is not None:
                st.session_state.processed_data = df_result
                st.success("집계 완료!")
            else:
                st.error(error_msg)

if st.session_state.processed_data is not None:
    st.divider()
    st.subheader("📋 결과 미리보기")
    st.dataframe(
        make_display_df(st.session_state.processed_data),
        use_container_width=True
    )

    excel_data = create_excel_with_formula(st.session_state.processed_data)
    st.download_button(
        "📥 엑셀 파일 다운로드",
        excel_data,
        "납품대금_집계표.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
