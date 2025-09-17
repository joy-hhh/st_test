import pandas as pd

coa_df = pd.read_excel("CoA_Level.xlsx", dtype=str)

parent_name = "모회사"
parent_ce_df = pd.read_excel("모회사_FS.xlsx", sheet_name="CE", header=None)
subs_names = ["자회사A", "자회사B"]
subs_ce_dfs = []
for f in ["자회사A_FS.xlsx", "자회사B_FS.xlsx"]:
    subs_ce_dfs.append(pd.read_excel(f, sheet_name="CE", header=None))


# def generate_sce_df(coa_df, parent_ce_df, subs_ce_dfs, parent_name, subs_names, adjustment_file, merged_bspl_df):

# 1. CoA 기반 동적 컬럼 정의
e_element_df = coa_df[coa_df['FS_Element'] == 'E'].dropna(axis=1).copy()

level_code_col = e_element_df.columns[-2]
level_name_col = e_element_df.columns[-1]

equity_groups = e_element_df[[level_code_col, level_name_col]].dropna().drop_duplicates().sort_values(by=level_code_col)
l3_codes_map = pd.Series(equity_groups[level_name_col].values, index=equity_groups[level_code_col]).to_dict()
nci_equity_row = coa_df[coa_df["FS_Element"] == "CE"]
nci_code = nci_equity_row.iloc[0]["계정코드"]
nci_name = nci_equity_row.iloc[0]["계정명"]
l3_codes_map[nci_code] = nci_name
sce_cols = list(l3_codes_map.values())
col_to_l3_map = {v: k for k, v in l3_codes_map.items()}

# 2. 입력된 CE 시트 파싱 (위치 기반으로 수정)
all_parsed_dfs = []
all_input_dfs = [(parent_name, parent_ce_df)] + list(zip(subs_names, subs_ce_dfs))
for name, df in all_input_dfs:
    if df.empty:
        continue
    try:
        code_row_index = df[df.iloc[:, 2] == '계정코드'].index[0]
        codes = df.iloc[code_row_index, 3:].astype(str).str.strip().tolist()
        data_start_row = code_row_index + 1
        
        data_df = df.iloc[data_start_row:].copy()
        
        num_desc_cols = 3
        num_data_cols = len(codes)
        data_df = data_df.iloc[:, :(num_desc_cols + num_data_cols)]
        # 위치를 기준으로 컬럼 이름을 명시적으로 지정
        desc_names = ['회사명', '구분', '계정코드']
        data_df.columns = desc_names + codes 
        all_parsed_dfs.append(data_df)
    except (IndexError, KeyError) as e:
        continue

combined_ce_df = pd.concat(all_parsed_dfs, ignore_index=True)
 

for col in sce_cols:
    if col in combined_ce_df.columns:
        combined_ce_df[col] = pd.to_numeric(combined_ce_df[col], errors='coerce').fillna(0)
    else:
        combined_ce_df[col] = 0

# 3. 기초자본 계산
beginning_simple_sum = combined_ce_df[combined_ce_df['계정코드'] == 'Beginning'][sce_cols].sum()
adj_xls = "CAJE_generated.xlsx"
full_adj_df = pd.read_excel(adj_xls, sheet_name="CAJE_BSPL", dtype={'계정코드': str})

beginning_adjustments = pd.Series(dtype='float64')
full_adj_df = full_adj_df.dropna(subset=['계정코드'])
full_adj_df['계정코드'] = full_adj_df['계정코드'].astype(str).str.strip().str.split('.').str[0]

if 'FS_Element' in full_adj_df.columns:
    full_adj_df = full_adj_df.drop(columns=['FS_Element'])

full_adj_df = full_adj_df.merge(coa_df[['계정코드', 'FS_Element', 'L3_code']], on='계정코드', how='left')
# L3_code가 없는 자본/비지배지분 항목은 계정코드를 L3_code로 사용
is_equity_like = full_adj_df['FS_Element'].isin(['E', 'CE'])
is_l3_missing = full_adj_df['L3_code'].isna()
full_adj_df.loc[is_equity_like & is_l3_missing, 'L3_code'] = full_adj_df.loc[is_equity_like & is_l3_missing, '계정코드']
full_adj_df['금액'] = pd.to_numeric(full_adj_df['금액'], errors='coerce').fillna(0)
beg_adj_df = full_adj_df[full_adj_df['당기전기'] != '당기'].copy()
beg_equity_adjs = beg_adj_df[beg_adj_df['FS_Element'].isin(['E', 'CE'])].copy()
if not beg_equity_adjs.empty:
    beg_equity_adjs.loc[:, '금액'] *= -1
    beginning_adjustments = beg_equity_adjs.groupby('L3_code')['금액'].sum()
beginning_row = pd.Series(0, index=sce_cols, name='기초')
beginning_row.update(beginning_simple_sum)
for code, amount in beginning_adjustments.items():
    if code in l3_codes_map:
        beginning_row[l3_codes_map[code]] += amount

# 4. 당기 변동분 계산
current_changes_df = combined_ce_df[~combined_ce_df['계정코드'].isin(['Beginning', 'Ending'])].copy()

r_adj_sum = merged_bspl_df.loc[merged_bspl_df["FS_Element"] == "R", "연결조정"].sum()
x_adj_sum = merged_bspl_df.loc[merged_bspl_df["FS_Element"] == "X", "연결조정"].sum()
pl_adj_sum = -r_adj_sum - x_adj_sum
nci_pl_adj = merged_bspl_df.loc[merged_bspl_df["FS_Element"] == "CR", "연결조정"].sum()

ni_adj_row = pd.DataFrame([{'구분': '당기순손익(연결조정)', '이익잉여금': pl_adj_sum - nci_pl_adj, '비지배지분': nci_pl_adj}]).fillna(0)

curr_adj_df = full_adj_df[full_adj_df['당기전기'] == '당기'].copy()
curr_equity_adjs = curr_adj_df[curr_adj_df['FS_Element'].isin(['E', 'CE'])].copy()
curr_equity_adjs.loc[:, '금액'] *= -1

pl_related_codes = merged_bspl_df[merged_bspl_df['FS_Element'].isin(['R', 'X', 'CR'])]['계정코드'].unique()
direct_equity_adjs = curr_equity_adjs[~curr_equity_adjs['계정코드'].isin(pl_related_codes)]
direct_adj_by_code = direct_equity_adjs.groupby('L3_code')['금액'].sum()
direct_adj_row_data = {'구분': '기타자본(연결조정)'}
for code, amount in direct_adj_by_code.items():
    if code in l3_codes_map:
        direct_adj_row_data[l3_codes_map[code]] = amount
direct_adj_row = pd.DataFrame([direct_adj_row_data]).fillna(0)

# 5. 최종 조립
beg_sce = pd.DataFrame([beginning_row])
final_sce = pd.concat([beg_sce, current_changes_df.groupby('구분')[sce_cols].sum()], ignore_index=False)
final_sce = pd.concat([final_sce, ni_adj_row.set_index('구분')], ignore_index=False)
if not direct_adj_row.empty and direct_adj_row.drop(columns=['구분']).iloc[0].abs().sum() > 1:
     final_sce = pd.concat([final_sce, direct_adj_row.set_index('구분')], ignore_index=False)
final_sce = final_sce.loc[(final_sce[sce_cols].abs().sum(axis=1)) > 1].fillna(0)
final_sce.loc['기말', sce_cols] = final_sce[sce_cols].sum()
# 6. 검증 행 추가
l3_map = dict(zip(coa_df['계정코드'], coa_df['L3_code']))
if 'L3_code' not in merged_bspl_df.columns:
     merged_bspl_df['L3_code'] = merged_bspl_df['계정코드'].map(l3_map)
# For CE elements (NCI), if L3_code is null, use the account code itself.
is_ce = merged_bspl_df['FS_Element'] == 'CE'
is_l3_missing = merged_bspl_df['L3_code'].isna()
merged_bspl_df.loc[is_ce & is_l3_missing, 'L3_code'] = merged_bspl_df.loc[is_ce & is_l3_missing, '계정코드']
l3_totals = merged_bspl_df.groupby('L3_code')['연결금액'].sum()

verification_row = pd.Series(index=sce_cols, name="검증(연결BS)")
for col, code in col_to_l3_map.items():
    verification_row[col] = l3_totals.get(code, 0)
final_sce.loc['검증(연결BS)'] = verification_row
final_sce = final_sce.reset_index().rename(columns={'index': '구분'})

# '구분'에 중복이 있을 수 있으므로, 첫 번째 '계정코드'를 사용하도록 중복을 제거하여 map을 생성
temp_map_df = combined_ce_df[['구분', '계정코드']].dropna(subset=['구분']).drop_duplicates(subset=['구분'])
row_to_code_map = pd.Series(temp_map_df.계정코드.values, index=temp_map_df.구분).to_dict()
row_to_code_map.update({'기초': 'Beginning', '기말': 'Ending', '검증(연결BS)': 'Verification', '당기순손익(연결조정)': 'CE11_NI', '기타자본(연결조정)': 'CE12_CAJE'})
final_sce.insert(1, '조정코드', final_sce['구분'].map(row_to_code_map).fillna('CE9999'))

return final_sce