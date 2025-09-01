import pandas as pd

coa = pd.read_excel("CoA_Level.xlsx", dtype=str)
e_element = coa[coa["FS_Element"]=="E"].dropna(axis=1).copy()

# 찾고 싶은 계정코드를 변수에 저장
account_code_to_find = "37500"

# .loc을 사용하여 인덱스에서 계정코드를 찾고, 원하는 열을 선택
result_by_position = e_element[e_element['계정코드'] == account_code_to_find].iloc[:, -2:]

# .values[0]는 첫 번째 행을 배열로 반환합니다.
# 이 배열의 각 요소를 l3_code와 l3_name 변수에 순서대로 할당합니다.
l3_code, l3_name = result_by_position.values[0]

print(f"그룹 코드: {l3_code}")
print(f"그룹 이름: {l3_name}")


parent_name = "모회사"
subs_names = ["자회사A", "자회사B"]
    
xls_parent = pd.ExcelFile("모회사_FS.xlsx")
parent_ce_df = pd.read_excel(xls_parent, sheet_name="CE", header=None) if "CE" in xls_parent.sheet_names else pd.DataFrame()

subs_ce_dfs = []
for f in ["자회사A_FS.xlsx", "자회사B_FS.xlsx"]:
    xls_sub = pd.ExcelFile(f)
    subs_ce_dfs.append(pd.read_excel(xls_sub, sheet_name="CE", header=None) if "CE" in xls_sub.sheet_names else pd.DataFram)



# 2. 입력된 CE 시트 파싱 (위치 기반으로 수정)
all_parsed_dfs = []
all_input_dfs = [(parent_name, parent_ce_df)] + list(zip(subs_names, subs_ce_dfs))

for name, df in all_input_dfs:
    if df.empty:
        continue
    code_row_index = df[df.iloc[:, 2] == '계정코드'].index[0]
    codes = df.iloc[code_row_index, 3:].astype(str).str.strip().tolist()
    data_start_row = code_row_index + 1
    
    data_df = df.iloc[data_start_row:].copy()
    
    num_desc_cols = 3
    num_data_cols = len(codes)
    data_df = data_df.iloc[:, :(num_desc_cols + num_data_cols)]
    # 위치를 기준으로 컬럼 이름을 명시적으로 지정
    desc_names = ['_col_company', '구분', '계정코드'] 
    data_df.columns = desc_names + codes
    data_df['회사명'] = name
    all_parsed_dfs.append(data_df)

coa_df = pd.read_excel("CoA_Level.xlsx", dtype=str)
e_element_df = coa_df[coa_df['FS_Element'] == 'E'].dropna(axis=1).copy()
level_code_col = e_element_df.columns[-2]
level_name_col = e_element_df.columns[-1]
equity_groups = e_element_df[[level_code_col, level_name_col]].dropna().drop_duplicates().sort_values(by=level_code_col)
l3_codes_map = pd.Series(equity_groups[level_name_col].values, index=equity_groups[level_code_col]).to_dict()


combined_ce_df = pd.concat(all_parsed_dfs, ignore_index=True)
rename_dict = {code: name for code, name in l3_codes_map.items() if code in combined_ce_df.columns}
combined_ce_df.rename(columns=rename_dict, inplace=True)
        
beginning_simple_sum = combined_ce_df[combined_ce_df['계정코드'] == 'Beginning'][sce_cols].sum()