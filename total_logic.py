import pandas as pd
from pathlib import Path

# 1. 파일 불러오기
bspl_df = pd.read_excel("bspl.xlsx", dtype={"계정코드": str})

bspl_s_files = ["bspl_s1.xlsx", "bspl_s2.xlsx"]
# 자회사별 DataFrame 로드
bspl_s_dfs = [pd.read_excel(f, dtype={"계정코드": str}).rename(columns=lambda x: x.strip()) for f in bspl_s_files]


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
warn_duplicate_codes_in_statement(bspl_df, name="모회사")
warn_duplicate_codes_in_statement(bspl_s_dfs, name="자회사")



# 2. CoA 전체 계층 정보도 불러오기
coa_all = pd.read_excel("CoA_Level.xlsx", dtype=str)


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
    
    



# 3. FS 항목과 CoA 병합
merged = coa_all.merge(bspl_df[["계정코드", "금액"]], how="left", on="계정코드").rename(columns={"금액": "지배회사"})
for i, df in enumerate(bspl_s_dfs):
    merged = merged.merge(df[["계정코드", "금액"]], how="left", on="계정코드").rename(columns={"금액": f"자회사{i+1}"})


# 금액열 자동 감지
amount_cols = merged.select_dtypes(include='number').columns.tolist()


# 연결조정 불러오기
con_adj = pd.read_excel("con_adj.xlsx", dtype={"계정코드": str})
con_adj["금액"].sum() == 0
con_adj_grouped = con_adj.groupby(["계정코드"], as_index=False)["금액"].sum()

def adj_sign(df: pd.DataFrame, coa_all: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    
    df = df.merge(coa_all[["계정코드", "FS_Element"]], how='left', on="계정코드")    
    # FS_Element에 따른 부호 매핑
    adj_sign_map = {
        "A":  1,   # 자산 +
        "L": -1,   # 부채 -
        "E": -1,   # 자본 -
        "R": -1,   # 수익 -
        "X":  1    # 비용 +
    }

    # sign 열 추가
    df["sign"] = df["FS_Element"].map(adj_sign_map).fillna(1)

    # 숫자형 열 찾기
    amount_cols = df.select_dtypes(include="number").columns.tolist()
    amount_cols = [c for c in amount_cols if c != "sign"]  # sign 자체는 제외

    # 각 숫자형 열에 sign 곱하기
    for col in amount_cols:
        df[col] = df[col] * df["sign"]
    
    
    return df




sined_con_adj = adj_sign(con_adj_grouped, coa_all)

con_adj_grouped["금액"].sum()
sined_con_adj["금액"].sum()

# 단순합계 및 연결조정 열 생성 
merged["단순합계"] = merged[amount_cols].sum(axis=1)

merged = merged.merge(
    con_adj_grouped[["계정코드", "금액"]].rename(columns={"금액": "연결조정"}),
    how="left",
    on="계정코드"
)

# merged = merged.iloc[:, :-1]   # 마지막 열 제거

merged = merged.merge(
    sined_con_adj[["계정코드", "금액"]],
    how="left",
    on="계정코드"
)

merged["연결금액"] = merged[["단순합계","금액"]].sum(axis=1)
merged = merged.drop(columns="금액")




con_amtcols = merged.select_dtypes(include='number').columns.tolist()


# 재무상태표: 자산(A), 부채(L), 자본(E)
df_bs = merged[merged["FS_Element"].isin(["A", "L", "E"])].copy()

# 손익계산서: 수익(R), 비용(E)
df_pl = merged[merged["FS_Element"].isin(["R", "X"])].copy()



# 계정코드 열
code_idx = [idx for idx in range(3, len(coa_all.columns)-1, 2)]
code_cols = coa_all.columns[code_idx].tolist()

# 계정명은 코드 열 바로 다음 열이므로 +1
name_idx = [i+1 for i in code_idx]
name_cols = coa_all.columns[name_idx].tolist()

# iloc[0, :]으로 뽑은 한 행에서 값이 비어있는 열을 찾으려면 
row = coa_all.iloc[0, :]  # 첫 번째 행 선택
missing_cols = row[row.isna() | (row.astype(str).str.strip() == "")].index.tolist()

# bs 해당 열만 리스트로 정리
code_cols_bs = [c for c in code_cols if c not in missing_cols]
name_cols_bs = [c for c in name_cols if c not in missing_cols]

code_name =  [[n, c] for n, c in zip(code_cols_bs, name_cols_bs)]
    
code_name_dict = {}
for i in code_name:
    code_name_unique = df_bs[i].drop_duplicates().dropna()
    code_name_dict.update(dict(zip(code_name_unique.iloc[:, 0], code_name_unique.iloc[:, 1])))


name_code_dict = {v: k for k, v in code_name_dict.items()}



# 재무상태표 그룹 합계를 먼저 쓰게 하는 재귀함수
def bs_recursive(df, name_cols_bs, amount_cols ,label_col='계정명'):
    df = df.copy()
    
    # 재귀 종료 조건
    if not name_cols_bs:
        return [row.to_dict() for _, row in df.iterrows()]

    current_col = name_cols_bs[0]
    result = []

    for key, group in df.groupby(current_col, sort=False):
        # 현재 그룹 합계 행 생성 (그룹 값들 그대로 유지)
            
        subtotal_raw = group[amount_cols].sum(skipna=True)
        
        sum_row = {col: "" for col in df.columns}
        sum_row.update(group.iloc[0].to_dict())
        
        sum_row[current_col] = f"{key}"        
        sum_row["계정명"] = f"{key}"
        
         # ✅ 현재 그룹의 코드 열을 찾아서 계정코드에 입력
        sum_row["계정코드"] = name_code_dict.get(key)

        for col in amount_cols:
                sum_row[col] = subtotal_raw[col]
        
                
        result.append(sum_row)

        # 하위 그룹/데이터 추가
        result.extend(bs_recursive(group, name_cols_bs[1:], amount_cols))

    return result



financial_position = pd.DataFrame(bs_recursive(df_bs, name_cols_bs, con_amtcols))
financial_position.to_excel("financial_position.xlsx")



# 손익계산서 중첩 더하기 빼기로 순이익 계산 그룹합
def signed_income_statement(df, code_cols, name_cols, amount_cols):
    df = df.copy()

    df["sign"] = df["FS_Element"].map({"R": 1, "X": -1})
    for col in amount_cols:
        df[col + "_signed"] = df[col] * df["sign"]

    signed_cols = [c for c in df.columns if c.endswith("_signed")]

    for col in name_cols + code_cols:
        if col in df.columns:
            df[col] = df[col].fillna("")

    added_keys = set()

    def find_first_code(group):
        """하위 데이터에서 가장 먼저 등장하는 코드 반환"""
        for c in code_cols[::-1]:  # L5 -> L1 순으로 탐색
            vals = group[c].dropna().unique()
            vals = [v for v in vals if str(v).strip() != ""]
            if len(vals) > 0:
                return vals[0]
        return ""

    def recursive(data, cols, parents=()):
        # 재귀 종료 조건
        if not cols:
            return data.to_dict("records")

        current_col = cols[0]
        next_cols = cols[1:]
        result = []

        for key, group in data.groupby(current_col, sort=False):
            children = recursive(group, next_cols, parents + (key,))
            result.extend(children)

            if all(k.strip() == "" for k in parents + (key,)):
                continue

            full_path = tuple(k for k in parents + (key,) if k.strip() != "")
            if full_path in added_keys:
                continue
            added_keys.add(full_path)

            subtotal_signed = group[signed_cols].sum(skipna=True)

            sum_row = {col: "" for col in data.columns}
            sum_row.update(group.iloc[0].to_dict())

            last_name = full_path[-1]
            sum_row[current_col] = f"{last_name}"
            sum_row["계정명"] = f"{last_name}"

            # code가 비어 있으면 하위 코드 중 첫 번째 값 사용
            code_val = group.iloc[0][f"L{len(parents)+1}_code"]
            if str(code_val).strip() == "":
                code_val = find_first_code(group)
            sum_row["계정코드"] = code_val

            for col in signed_cols:
                sum_row[col] = subtotal_signed[col]

            result.append(sum_row)

        return result

    # ✅ 최종 결과 생성
    out_df = pd.DataFrame(recursive(df, name_cols))

    # ✅ amount_cols = (_signed 열) × sign
    for col in amount_cols:
        signed_col = col + "_signed"
        out_df[col] = out_df[signed_col] * out_df["sign"]

    # ✅ sign, _signed 열 제거
    out_df = out_df.drop(columns=["sign"] + signed_cols, errors="ignore")

    return out_df
    


# ✅ 실행 예시
income_statement = signed_income_statement(df_pl, code_cols, name_cols, con_amtcols)

# 엑셀로 저장
income_statement.to_excel("income_statement.xlsx", index=False)


level_cols = code_cols + name_cols
con_wp = pd.concat([financial_position, income_statement]).drop(level_cols, axis=1)

con_wp.to_excel("con_wp.xlsx")




## footnote 대사 확인
# 주석과 재무제표 로드
footnote_df = pd.read_excel("footnote.xlsx", dtype={"계정코드": str})
footnote_s1_df = pd.read_excel("footnote_s1.xlsx", dtype={"계정코드": str})
footnote_s2_df = pd.read_excel("footnote_s2.xlsx", dtype={"계정코드": str})


bspl_df = pd.read_excel("bspl.xlsx", dtype={"계정코드": str})




def read_footnote_files(files):
    footnote_dict = {}

    for file in files:
        xls = pd.ExcelFile(file)
        target_sheets = [s for s in xls.sheet_names if "주석" in s]

        for sheet in target_sheets:
            df = pd.read_excel(file, sheet_name=sheet)
            df["소스파일"] = Path(file).stem  # 파일명 추가

            if sheet not in footnote_dict:
                footnote_dict[sheet] = []
            footnote_dict[sheet].append(df)

    return footnote_dict


def combine_by_position(df_list):
    combined = df_list[0].copy()

    for df in df_list[1:]:
        for col in df.columns[2:-1]:  # 코드/행이름 제외, 소스파일 제외
            if pd.api.types.is_numeric_dtype(df[col]):
                combined[col] = combined[col].fillna(0) + df[col].fillna(0)
            else:
                # 문자형 열이 포함 → concat 후 파일명 유지
                return pd.concat(df_list, ignore_index=True)

    # 숫자형만 있어서 합산이 끝난 경우 → 소스파일 열을 combined로 변경
    combined["소스파일"] = "combined"
    return combined


def combine_all_files(files):
    footnote_dict = read_footnote_files(files)
    combined_result = {}

    for sheet_name, df_list in footnote_dict.items():
        combined_result[sheet_name] = combine_by_position(df_list)

    return combined_result




# ✅ 실행

files = list(Path.cwd().glob("*footnote*.xlsx"))
combined_result = combine_all_files(files)



with pd.ExcelWriter("combined_footnote.xlsx", engine="openpyxl") as writer:
    for sheet_name, df in combined_result.items():
        df.to_excel(writer, sheet_name=sheet_name[:31], index=False)





