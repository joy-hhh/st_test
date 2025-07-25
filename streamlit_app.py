# streamlit_app.py
import streamlit as st
import pandas as pd
from io import BytesIO

# ==========================
# 전처리 + 합산 함수
# ==========================
def clean_footnote_rowwise(df):
    df = df.copy()
    df.columns = df.columns.str.strip()
    df["계정코드"] = df["계정코드"].astype(str).str.strip().replace("nan", "")
    df["구분"] = df["구분"].astype(str).str.strip()

    numeric_cols = [col for col in df.columns if col not in ["계정코드", "구분"]]

    for col in numeric_cols:
        df[col] = (
            df[col].astype(str)
            .str.replace(",", "")
            .str.replace("(", "-")
            .str.replace(")", "")
            .astype(float)
        )

    return df, numeric_cols

def sum_footnotes_preserve_rows(footnote_dfs: list[pd.DataFrame]) -> pd.DataFrame:
    if not footnote_dfs:
        return pd.DataFrame()

    cleaned_dfs = []
    numeric_cols = None
    for df in footnote_dfs:
        cleaned, cols = clean_footnote_rowwise(df)
        numeric_cols = cols
        cleaned_dfs.append(cleaned)

    base = cleaned_dfs[0][["계정코드", "구분"]].copy()

    for col in numeric_cols:
        base[col] = sum(df[col] for df in cleaned_dfs)

    # 계정코드 보완
    for df in cleaned_dfs:
        base["계정코드"] = base["계정코드"].mask(base["계정코드"].isin(["", "nan", "None"]), df["계정코드"])

    return base

# ==========================
# FS 비교 함수
# ==========================
def check_fs_match(footnote_sum: pd.DataFrame, bspl_df: pd.DataFrame) -> pd.DataFrame:
    df = footnote_sum.copy()
    bspl_df = bspl_df.copy()

    df["계정코드"] = df["계정코드"].astype(str).str.strip()
    bspl_df["계정코드"] = bspl_df["계정코드"].astype(str).str.strip()

    fs_map = bspl_df.set_index("계정코드")["금액"].to_dict()
    book_value_col = df.columns[-1]

    def compare(row):
        code = row["계정코드"]
        if pd.isna(code) or code == "":
            return ""
        fs_val = fs_map.get(code, 0)
        return "일치" if abs(row[book_value_col] - fs_val) < 1 else "불일치"

    df["FS비교"] = df.apply(compare, axis=1)
    return df

# ==========================
# 엑셀 다운로드 변환 함수
# ==========================
def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="FS_비교결과")
    output.seek(0)
    return output

# ==========================
# Streamlit 앱
# ==========================
st.title("📑 주석 합산 + FS 비교 앱")

uploaded_footnotes = st.file_uploader("주석 파일들 (모회사+자회사 전체)", type="xlsx", accept_multiple_files=True)
uploaded_bspl = st.file_uploader("재무제표 파일 (bspl)", type="xlsx")

if uploaded_footnotes and uploaded_bspl:
    # 주석 데이터 읽기
    footnote_dfs = [pd.read_excel(file, dtype={"계정코드": str}) for file in uploaded_footnotes]
    bspl_df = pd.read_excel(uploaded_bspl, dtype={"계정코드": str})

    # 합산 + 비교
    sum_df = sum_footnotes_preserve_rows(footnote_dfs)
    result_df = check_fs_match(sum_df, bspl_df)

    # 미리보기
    st.subheader("✅ FS 비교 결과")
    st.dataframe(result_df)

    # 다운로드
    excel_bytes = convert_df_to_excel(result_df)
    st.download_button(
        label="📥 엑셀로 다운로드",
        data=excel_bytes,
        file_name="FS_비교결과.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("주석 파일과 재무제표 파일을 모두 업로드해주세요.")
