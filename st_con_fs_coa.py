import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="연결 재무제표 자동 집계", layout="wide")
st.title("📊 연결 재무제표 집계 자동화")

# ------------------------------
# 파일 업로드
# ------------------------------
st.sidebar.header("📁 파일 업로드")
coa_file = st.sidebar.file_uploader("🔢 CoA 파일 (계정코드 + 부호 + 계층구조 포함)", type="xlsx")
bspl_file = st.sidebar.file_uploader("🏢 모회사 재무제표", type="xlsx")
bspl_s_files = st.sidebar.file_uploader("🏬 자회사 재무제표들", type="xlsx", accept_multiple_files=True)
adjust_file = st.sidebar.file_uploader("🛠 연결조정 파일 (선택)", type="xlsx")

# ------------------------------
# 유틸 함수
# ------------------------------
def load_excel(file, name):
    df = pd.read_excel(file, dtype={"계정코드": str})
    if "계정코드" not in df.columns or "금액" not in df.columns:
        st.error(f"[{name}] 파일에는 반드시 '계정코드'와 '금액' 열이 있어야 합니다.")
    return df

def insert_group_totals_below(df, group_col, label_col_name, value_col="조정금액"):
    grouped = []
    for key, group in df.groupby(group_col, sort=False):
        subtotal = group[value_col].sum()
        last_row = group.iloc[-1]
        grouped.append(group)
        summary_row = {
            group.columns[0]: f"{key}_합계",
            label_col_name: str(last_row[label_col_name]) + " 합계",
            value_col: subtotal,
        }
        for col in df.columns:
            if col not in summary_row:
                summary_row[col] = None
        summary_df = pd.DataFrame([summary_row])
        grouped.append(summary_df)
    return pd.concat(grouped, ignore_index=True)

# ------------------------------
# 처리 시작
# ------------------------------
if coa_file and bspl_file:
    coa_df = pd.read_excel(coa_file, dtype={"계정코드": str})
    bspl_df = load_excel(bspl_file, "모회사")
    bspl_s_dfs = [load_excel(f, f"자회사{i+1}") for i, f in enumerate(bspl_s_files)] if bspl_s_files else []
    adjust_df = load_excel(adjust_file, "연결조정") if adjust_file else pd.DataFrame(columns=["계정코드", "금액"])

    # 계정코드 기준 병합
    merged = coa_df.copy()
    merged = coa_df.iloc[:, [1, 2]]
    merged = merged.merge(bspl_df[["계정코드", "금액"]], on="계정코드", how="left").rename(columns={"금액": "모회사"})

    for i, df in enumerate(bspl_s_dfs):
        merged = merged.merge(df[["계정코드", "금액"]], on="계정코드", how="left").rename(columns={"금액": f"자회사{i+1}"})

    merged = merged.merge(adjust_df, on="계정코드", how="left").rename(columns={"금액": "연결조정"})

    # 결측값 처리 및 단순합산
    value_cols = ["모회사"] + [f"자회사{i+1}" for i in range(len(bspl_s_dfs))] + ["연결조정"]
    merged[value_cols] = merged[value_cols].fillna(0)
    merged["단순합산"] = merged[["모회사"] + [f"자회사{i+1}" for i in range(len(bspl_s_dfs))]].sum(axis=1)
    merged["최종금액"] = merged["단순합산"] + merged["연결조정"]

    # 부호 적용한 조정금액
    merged["부호값"] = merged["부호"].map({"+": 1, "-": -1})
    merged["조정금액"] = merged["최종금액"] * merged["부호값"].fillna(0)

    # 집계 삽입 (모든 레벨 순서대로)
    # 순서 기준으로 부호, 계정코드, L, 계정코드, L... 추출
    level_cols = [col for col in merged.columns if col != "부호"]
    level_pairs = [(level_cols[i], level_cols[i+1]) for i in range(0, len(level_cols)-1, 2)]
    level_map = level_pairs
    # 이미 level_map 위에서 재정의됨

    final = merged.copy()
    for code_col, label_col in level_map:
        final = insert_group_totals_below(final, group_col=code_col, label_col_name=label_col)

    # 출력
    st.subheader("📋 연결재무제표 결과")
    col_subset = [col for col in ["계정코드", "L1", "L2", "L3", "L4", "L5", "L6"] if col in final.columns]
    col_subset += ["모회사"] + [f"자회사{i+1}" for i in range(len(bspl_s_dfs))] + ["연결조정", "단순합산", "최종금액", "조정금액"]
    st.dataframe(final[col_subset])

    # 다운로드
    @st.cache_data
    def convert_df(df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        return output.getvalue()

    st.download_button(
        "📥 결과 다운로드 (Excel)",
        convert_df(final),
        file_name="연결재무제표_집계결과.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("좌측에서 CoA와 모회사 파일을 업로드하세요. 자회사와 연결조정은 선택입니다.")

