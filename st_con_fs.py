import streamlit as st
import pandas as pd

st.set_page_config(page_title="연결재무제표 도구", layout="wide")
st.title("📊 연결재무제표 합산 도구 (연결조정 포함)")

# -------------------------------
# 📂 파일 업로드
# -------------------------------
st.sidebar.header("1. 파일 업로드")
coa_file = st.sidebar.file_uploader("🧾 CoA 파일", type=["xlsx"], key="coa")
bspl_file = st.sidebar.file_uploader("🏢 모회사 재무제표", type=["xlsx"], key="bspl")
bspl_s_files = st.sidebar.file_uploader("🏬 자회사 재무제표들", type=["xlsx"], accept_multiple_files=True, key="bspl_s")
adjust_file = st.sidebar.file_uploader("🔧 연결조정 파일 (선택)", type=["xlsx"], key="adjust")


# -------------------------------
# 📑 엑셀 로드 및 유효성 검사
# -------------------------------
def load_clean_excel(file, name, require_amount=False):
    df = pd.read_excel(file, dtype={"계정코드": str}).rename(columns=lambda x: x.strip())
    cols = df.columns.tolist()

    if "계정코드" not in cols:
        st.warning(f"⚠️ [{name}] '계정코드' 열이 없습니다.")
    if require_amount and "금액" not in cols:
        st.warning(f"⚠️ [{name}] '금액' 열이 없습니다.")
    return df


def warn_duplicate(df, name):
    dup = df["계정코드"].value_counts()
    dups = dup[dup > 1]
    if not dups.empty:
        st.warning(f"⚠️ [{name}] 중복 계정코드 {len(dups)}개:\n" + "\n".join(f"- {k}: {v}회" for k, v in dups.items()))
    else:
        st.success(f"✅ [{name}] 중복 계정코드 없음")


def check_invalid_codes(coa_df, df, name):
    valid = set(coa_df["계정코드"])
    input_ = set(df["계정코드"])
    invalid = input_ - valid
    if invalid:
        st.error(f"🚨 [{name}] CoA에 없는 계정코드 {len(invalid)}개:\n" + "\n".join(f"- {x}" for x in sorted(invalid)))
    else:
        st.success(f"✅ [{name}] 모든 계정코드가 CoA에 존재")


# -------------------------------
# ✅ 모든 파일 준비 시 처리
# -------------------------------
if coa_file and bspl_file and bspl_s_files:

    coa_df = load_clean_excel(coa_file, "CoA", require_amount=False)
    bspl_df = load_clean_excel(bspl_file, "모회사", require_amount=True)
    bspl_s_dfs = [load_clean_excel(f, f"자회사 {i+1}", require_amount=True) for i, f in enumerate(bspl_s_files)]
    adjust_df = load_clean_excel(adjust_file, "연결조정", require_amount=True) if adjust_file else pd.DataFrame(columns=["계정코드", "금액"])

    # 유효성 검사
    warn_duplicate(bspl_df, "모회사")
    check_invalid_codes(coa_df, bspl_df, "모회사")

    for i, df in enumerate(bspl_s_dfs):
        warn_duplicate(df, f"자회사 {i+1}")
        check_invalid_codes(coa_df, df, f"자회사 {i+1}")

    if not adjust_df.empty:
        warn_duplicate(adjust_df, "연결조정")
        check_invalid_codes(coa_df, adjust_df, "연결조정")

    # -------------------------------
    # 🔗 병합
    # -------------------------------
    merged = coa_df.copy()
    merged = merged.merge(bspl_df[["계정코드", "금액"]], on="계정코드", how="left").rename(columns={"금액": "모회사"})

    for i, df in enumerate(bspl_s_dfs):
        merged = merged.merge(df[["계정코드", "금액"]], on="계정코드", how="left").rename(columns={"금액": f"자회사{i+1}"})

    if not adjust_df.empty:
        merged = merged.merge(adjust_df, on="계정코드", how="left").rename(columns={"금액": "연결조정"})
    else:
        merged["연결조정"] = 0

    # -------------------------------
    # 🧮 계산
    # -------------------------------
    all_cols = ["모회사"] + [f"자회사{i+1}" for i in range(len(bspl_s_dfs))]
    merged[all_cols + ["연결조정"]] = merged[all_cols + ["연결조정"]].fillna(0)

    merged["단순합산"] = merged[all_cols].sum(axis=1)
    merged["최종금액"] = merged["단순합산"] + merged["연결조정"]

    # -------------------------------
    # 📊 결과 출력
    # -------------------------------
    st.subheader("🔍 연결조정 포함 결과 미리보기")
    st.dataframe(merged)

    # 요약
    st.sidebar.markdown("### ✅ 금액 요약")
    st.sidebar.write(f"📌 모회사 총액: {bspl_df['금액'].sum():,.0f}")
    for i, df in enumerate(bspl_s_dfs):
        st.sidebar.write(f"📌 자회사{i+1} 총액: {df['금액'].sum():,.0f}")
    if not adjust_df.empty:
        st.sidebar.write(f"📌 연결조정 총액: {adjust_df['금액'].sum():,.0f}")
    st.sidebar.write(f"📌 최종합산 총액: {merged['최종금액'].sum():,.0f}")

    # -------------------------------
    # 📥 다운로드
    # -------------------------------
    @st.cache_data
    def convert_df(df):
        return df.to_excel(index=False, engine="openpyxl")

    st.download_button("📥 Excel 다운로드", convert_df(merged), file_name="연결재무제표_합산결과.xlsx")

else:
    st.info("📂 CoA, 모회사, 자회사 파일을 업로드하면 자동으로 병합됩니다. 연결조정은 선택입니다.")
