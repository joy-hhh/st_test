import pandas as pd

data = {
    '왼쪽': ['A', 'B', 'C', 'D', 'E', 'F'],
    '오른쪽': [1, 1, 2, 2, 3, 3],
    '값': [10, 20, 30, 40, 50, 60]
}
df = pd.DataFrame(data)

# 열 이름 자동 인식
cols = df.columns.tolist()
left_col, right_col, value_col = cols[0], cols[1], cols[2]

result_parts = []

for group, group_df in df.groupby(right_col, sort=False):
    sum_row = pd.DataFrame([{
        left_col: f"{group}_합계",
        right_col: group,
        value_col: group_df[value_col].sum()
    }])
    result_parts.append(pd.concat([group_df, sum_row], ignore_index=True))

result_df = pd.concat(result_parts, ignore_index=True)
print(result_df)





def prepare_all_group_sums(df, group_cols, value_cols):
    group_sums = {}
    for col in group_cols:
        sums = df.groupby(col, sort=False)[value_cols].sum(numeric_only=True)
        group_sums[col] = sums.to_dict()
    return group_sums


def insert_group_sums_with_precalc(df, group_cols, value_cols, label_col):
    group_sums = prepare_all_group_sums(
        df[~df[label_col].astype(str).str.endswith("_합계", na=False)],
        group_cols,
        value_cols
    )

    result_df = df.copy()

    for group_col in group_cols:
        base_df = result_df[~result_df[label_col].astype(str).str.endswith("_합계", na=False)]
        last_names = base_df.groupby(group_col, sort=False)[label_col].last().tolist()

        new_rows = []
        for _, row in result_df.iterrows():
            new_rows.append(row.to_dict())

            if row[label_col] in last_names:
                group_val = row[group_col]
                sum_vals = group_sums[group_col].get(group_val, {})

                sum_row = {col: pd.NA for col in result_df.columns}
                sum_row[group_col] = pd.NA      # ✅ 합계 행에서는 그룹 값 제거
                sum_row[label_col] = f"{group_val}_합계"

                for col in value_cols:
                    sum_row[col] = sum_vals.get(col, pd.NA)

                new_rows.append(sum_row)

        result_df = pd.DataFrame(new_rows).reset_index(drop=True)

    return result_df


# ✅ 테스트
data = {
    '왼쪽': ['A', 'B', 'C', 'D', 'E', 'F'],
    '오른쪽': [1, 1, 2, 2, 3, 3],
    '큰그룹': ['가', '가', '가', '가', '나', '나'],
    '더큰그룹': ['총', '총', '총', '총', '총', '총'],
    '값': [10, 20, 30, 40, 50, 60]
}

df = pd.DataFrame(data)
group_cols = ["오른쪽", "큰그룹", "더큰그룹"]
value_cols = ["값"]

df_result = insert_group_sums_with_precalc(df, group_cols, value_cols, "왼쪽")
print(df_result)



# 오른쪽 그룹별로 DataFrame을 쪼개서 담는 코드
# 오른쪽 값으로 그룹별 DataFrame을 딕셔너리로 저장
dfs_by_group = {key: gdf for key, gdf in df.groupby("오른쪽", sort=False)}

for k, v in dfs_by_group.items():
    print(f"=== 그룹 {k} ===")
    print(v)
    
dfs_by_group[1]  # 오른쪽이 1인 DataFrame
dfs_by_group[2]  # 오른쪽이 2인 DataFrame
dfs_by_group[3]  # 오른쪽이 3인 DataFrame


# 2️⃣ 그룹별 합계 DataFrame 생성
group_sums_df = df.groupby("오른쪽", sort=False, as_index=False)["값"].sum()


# 3️⃣ 합계 DataFrame을 dict로 변환 (그룹 값 기준)
sum_dfs_by_group = {row["왼쪽"]: pd.DataFrame([row]) for _, row in group_sums_df.iterrows()}

for k in dfs_by_group:
    print(f"\n=== 그룹 {k} ===")
    print(dfs_by_group[k])
    print(sum_dfs_by_group[k])
    
    
    
    
# 합계를 먼저 쓰게 하는 재귀함수
def build_nested_recursive(df, group_cols, label_col, value_col):
    if not group_cols:
        return [row.to_dict() for _, row in df.iterrows()]

    current_col = group_cols[0]
    result = []

    for group_val, gdf in df.groupby(current_col, sort=False):
        # 현재 그룹 합계 행 생성 (그룹 값들 그대로 유지)
        sum_row = gdf.iloc[0].to_dict()      # 첫 행을 복사해서 그룹 값 유지
        sum_row[label_col] = f"{group_val}_합계"
        sum_row[value_col] = gdf[value_col].sum()
        result.append(sum_row)

        # 하위 그룹/데이터 추가
        result.extend(build_nested_recursive(gdf, group_cols[1:], label_col, value_col))

    return result


# ✅ 테스트
data = {
    '왼쪽': ['A', 'B', 'C', 'D', 'E', 'F'],
    '오른쪽': [1, 1, 2, 2, 3, 3],
    '큰그룹': ['가', '가', '가', '가', '나', '나'],
    '더큰그룹': ['총', '총', '총', '총', '총', '총'],
    '값': [10, 20, 30, 40, 50, 60]
}

df = pd.DataFrame(data)

df_result = pd.DataFrame(build_nested_recursive(df, ["더큰그룹", "큰그룹", "오른쪽"], "왼쪽", "값"))
print(df_result)






# ✅ 손익계산서 DataFrame 불러오기
bspl_df = pd.read_excel("bspl.xlsx", dtype={"계정코드": str})

# coa_df = pd.read_excel("CoA_Level.xlsx", dtype=str).iloc[:, [1, 2]]

bspl_s_files = ["bspl_s1.xlsx", "bspl_s2.xlsx"]
# 자회사별 DataFrame 로드
bspl_s_dfs = [pd.read_excel(f, dtype={"계정코드": str}).rename(columns=lambda x: x.strip()) for f in bspl_s_files]
# 금액 열 int 및 NaN 0으로 변환
for i, df in enumerate(bspl_s_dfs):
    df["금액"] = pd.to_numeric(df["금액"], errors="coerce").fillna(0)
    bspl_s_dfs[i] = df


coa_all = pd.read_excel("CoA_Level.xlsx", dtype=str)  # 업로드한 파일을 사용
merged = coa_all.merge(bspl_df, how="left", on="계정코드").rename(columns={"금액": "지배회사"})
for i, df in enumerate(bspl_s_dfs):
    merged = merged.merge(df[["계정코드", "금액"]], on="계정코드", how="left").rename(columns={"금액": f"자회사{i+1}"})



# merged.to_excel("merged_df.xlsx")



# 재무상태표: 자산(A), 부채(L), 자본(E)
df_bs = merged[merged["FS_Element"].isin(["A", "L", "E"])].copy()

# 손익계산서: 수익(R), 비용(E)
df_pl = merged[merged["FS_Element"].isin(["R", "X"])].copy()
# df_pl.to_excel("df_pl.xlsx")




def build_signed_income_statement(df):
    df = df.copy()
    
    amount_cols = df.select_dtypes(include='number').columns.tolist()

    df["sign"] = df["FS_Element"].map({"R": 1, "X": -1})
    for col in amount_cols:
        df[col + "_signed"] = df[col] * df["sign"]

    signed_cols = [c for c in df.columns if c.endswith("_signed")]
    level_cols = ["L1_name", "L2_name", "L3_name", "L4_name", "L5_name"]
    code_cols = ["L1_code", "L2_code", "L3_code", "L4_code", "L5_code"]

    for col in level_cols + code_cols:
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

            subtotal_raw = group[amount_cols].sum(skipna=True)
            subtotal_signed = group[signed_cols].sum(skipna=True)

            sum_row = {col: "" for col in data.columns}
            sum_row.update(group.iloc[0].to_dict())

            last_name = full_path[-1]
            sum_row[current_col] = f"{last_name}"
            sum_row["계정명_y"] = f"{last_name}"

            # code가 비어 있으면 하위 코드 중 첫 번째 값 사용
            code_val = group.iloc[0][f"L{len(parents)+1}_code"]
            if str(code_val).strip() == "":
                code_val = find_first_code(group)
            sum_row["계정코드"] = code_val

            for col in amount_cols:
                sum_row[col] = subtotal_raw[col]
            for col in signed_cols:
                sum_row[col] = subtotal_signed[col]

            result.append(sum_row)

        return result

    return pd.DataFrame(recursive(df, level_cols))



# ✅ 실행 예시

final_income_statement = build_signed_income_statement(df_pl)


# 엑셀로 저장
final_income_statement.to_excel("income_statement.xlsx", index=False)



