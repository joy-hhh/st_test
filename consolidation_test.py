import pandas as pd

# 파일 로드
coa_df = pd.read_excel("CoA.xlsx").rename(columns=lambda x: x.strip())
bspl = pd.read_excel("bspl.xlsx").rename(columns=lambda x: x.strip())      # 모회사
bspl_s = pd.read_excel("bspl_s.xlsx").rename(columns=lambda x: x.strip())  # 자회사


# 중복 계정코드 확인
def warn_duplicate_codes_in_statement(df, name="재무제표"):
    """재무제표 내 중복된 계정코드를 경고"""
    dup_counts = df["계정코드"].value_counts()
    duplicates = dup_counts[dup_counts > 1]

    if not duplicates.empty:
        print(f"⚠️ [{name}] 내 중복된 계정코드가 {len(duplicates)}개 있습니다:")
        for code, count in duplicates.items():
            print(f" - 계정코드 {code}: {count}회 등장")
    else:
        print(f"✅ [{name}] 내에는 중복 계정코드가 없습니다.")


# 중복 검사 실행
warn_duplicate_codes_in_statement(bspl, name="모회사")
warn_duplicate_codes_in_statement(bspl_s, name="자회사")


# CoA 기준 외 계정코드 확인
def check_account_codes_against_coa(coa_df, df, name="재무제표"):
    """CoA 기준으로 계정코드 유효성 검사"""
    valid_codes = set(coa_df["계정코드"])
    input_codes = set(df["계정코드"])
    invalid_codes = input_codes - valid_codes

    if invalid_codes:
        print(f"🚨 [{name}] CoA에 없는 계정코드가 {len(invalid_codes)}개 있습니다:")
        for code in sorted(invalid_codes):
            print(f" - {code}")
    else:
        print(f"✅ [{name}] 모든 계정코드가 CoA에 존재합니다.")


# CoA 기준 외 계정코드 검사 실행
check_account_codes_against_coa(coa_df, bspl, name="모회사")
check_account_codes_against_coa(coa_df, bspl_s, name="자회사")
    
    
    
# 병합: CoA를 기준으로 각 재무제표 금액을 붙이기
merged = coa_df.copy()
merged = merged.merge(bspl[["계정코드", "금액"]], on="계정코드", how="left").rename(columns={"금액": "모회사"})
merged = merged.merge(bspl_s[["계정코드", "금액"]], on="계정코드", how="left").rename(columns={"금액": "자회사"})

# merge 확인
bspl_amount = bspl.loc[:,'금액'].sum()
bspl_s_amount = bspl_s.loc[:,'금액'].sum()
merged_amount_p = merged.loc[:,'모회사'].sum()
merged_amount_s = merged.loc[:,'자회사'].sum()

print(f'모회사 업로드 금액 Check :{merged_amount_p == bspl_amount}')
print(f'자회사 업로드 금액 Check :{merged_amount_s == bspl_s_amount}')




# 결측값 0으로 채움
merged[["모회사", "자회사"]] = merged[["모회사", "자회사"]].fillna(0)

# 단순합산 열 추가
merged["단순합산"] = merged["모회사"] + merged["자회사"]

# 결과 보기
print(merged.head())



## footnote 대사 확인
import pandas as pd


# 주석과 재무제표 로드
footnote_df = pd.read_excel("footnote.xlsx", dtype={"계정코드": str})
footnote_s_df = pd.read_excel("footnote_s.xlsx", dtype={"계정코드": str})


bspl_df = pd.read_excel("bspl.xlsx", dtype={"계정코드": str})

# 열 정리
footnote_df["계정코드"] = footnote_df["계정코드"].str.strip()
bspl_df["계정코드"] = bspl_df["계정코드"].str.strip()

# 재무제표 금액 매핑 딕셔너리
bspl_map = bspl_df.set_index("계정코드")["금액"].to_dict()

# 가장 오른쪽 열 이름 가져오기
book_value_col = footnote_df.columns[-1]

# 장부금액 숫자 변환
footnote_df[book_value_col] = (
    footnote_df[book_value_col]
    .astype(str)
    .str.replace(",", "")
    .str.replace("(", "-")
    .str.replace(")", "")
    .astype(float)
)


# 비교 결과 생성 함수
def compare(row):
    code = row["계정코드"]
    if pd.isna(code) or code == "":
        return ""
    fs_value = bspl_map.get(code, 0)
    val = row[book_value_col]
    return "일치" if abs(val - fs_value) < 1 else "불일치"

footnote_df["FS비교"] = footnote_df.apply(compare, axis=1)



## 두개의 인덱스 있는 경우의 합산
# 공통 전처리 함수
def clean_footnote(df):
    df = df.copy()
    df.columns = df.columns.str.strip()
    df["계정코드"] = df["계정코드"].astype(str).str.strip()
    df["구분"] = df["구분"].astype(str).str.strip()

    # 숫자열 정리
    numeric_cols = [col for col in df.columns if col not in ["계정코드", "구분"]]
    for col in numeric_cols:
        df[col] = (
            df[col].astype(str)
            .str.replace(",", "")
            .str.replace("(", "-")
            .str.replace(")", "")
            .astype(float)
        )

    return df.set_index(["계정코드", "구분"])

# 각각 정리
fn = clean_footnote(footnote_df)
fn_s = clean_footnote(footnote_s_df)

# 같은 위치끼리 더하기
footnote_sum = fn.add(fn_s, fill_value=0).reset_index()


# 동적으로 리스트 생성 📌 예: Streamlit에서 여러 파일 업로드되는 상황
uploaded_files = st.file_uploader("자회사 주석 파일들", accept_multiple_files=True, type="xlsx")
uploaded_footnote_dfs = []

for file in uploaded_files:
    df = pd.read_excel(file)
    uploaded_footnote_dfs.append(df)
    
    


import pandas as pd

def clean_footnote(df):
    """
    주석 DataFrame 전처리: 계정코드/구분 정리 + 숫자형 변환 + 복합인덱스 설정
    """
    df = df.copy()
    df.columns = df.columns.str.strip()
    df["계정코드"] = df["계정코드"].astype(str).str.strip()
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

    return df.set_index(["계정코드", "구분"])

# 자회사 수에 관계없이 유연하게 대응 가능한 sum_footnotes() 유틸 함수
def sum_footnotes(footnote_dfs: list[pd.DataFrame]) -> pd.DataFrame:
    """
    복수의 주석 DataFrame을 계정코드+구분 기준으로 합산하는 함수.
    합계 행 포함. 항목 정렬은 유지.
    """
    if not footnote_dfs:
        return pd.DataFrame()

    result = clean_footnote(footnote_dfs[0])
    for df in footnote_dfs[1:]:
        result = result.add(clean_footnote(df), fill_value=0)

    return result.reset_index()


# 모회사 + 자회사 주석 파일 읽기
footnote_df = pd.read_excel("footnote.xlsx")
footnote_s1 = pd.read_excel("footnote_s1.xlsx")
footnote_s2 = pd.read_excel("footnote_s2.xlsx")

# 리스트로 묶어서 합산
combined = sum_footnotes([footnote_df, footnote_s1, footnote_s2])
print(combined)


# 또는 Streamlit에서
uploaded_files = st.file_uploader("자회사 주석들", accept_multiple_files=True)
footnote_dfs = [pd.read_excel(f) for f in uploaded_files]

# 모회사 포함해서 합산
total_df = sum_footnotes([parent_df] + footnote_dfs)



# 합산된 주석표를 기준으로 FS 일치 여부를 판단하는 check_fs_match() 함수
def check_fs_match(footnote_sum: pd.DataFrame, bspl_df: pd.DataFrame) -> pd.DataFrame:
    """
    합산된 주석표와 재무제표를 비교해 FS 일치 여부를 추가하는 함수.
    - 계정코드 있는 행만 비교
    - 비교 대상 금액 열은 가장 오른쪽 열로 자동 인식
    - 일치하면 "일치", 불일치하면 "불일치", 비교 불가면 공란
    """
    df = footnote_sum.copy()
    bspl_df = bspl_df.copy()

    # 계정코드 정리
    df["계정코드"] = df["계정코드"].astype(str).str.strip()
    bspl_df["계정코드"] = bspl_df["계정코드"].astype(str).str.strip()

    # 금액 맵 생성
    bspl_map = bspl_df.set_index("계정코드")["금액"].to_dict()

    # 비교할 장부금액 열 (가장 오른쪽 열)
    book_value_col = df.columns[-1]

    # 비교 함수 정의
    def fs_compare(row):
        code = row["계정코드"]
        if pd.isna(code) or code == "":
            return ""
        fs_value = bspl_map.get(code, 0)
        return "일치" if abs(row[book_value_col] - fs_value) < 1 else "불일치"

    # 비교 결과 컬럼 추가
    df["FS비교"] = df.apply(fs_compare, axis=1)
    return df

# 사용예시
# 모든 주석 합산
footnote_sum = sum_footnotes([footnote_df, footnote_s1, footnote_s2])

# 재무제표 불러오기
bspl_df = pd.read_excel("bspl.xlsx", dtype={"계정코드": str})

# FS 비교
result = check_fs_match(footnote_sum, bspl_df)

# 결과 출력
print(result[["계정코드", "구분", "FS비교"]])



