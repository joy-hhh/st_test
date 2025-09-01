# 연결자본변동표 합산 로직

import pandas as pd

try:
    # 엑셀 파일을 DataFrame으로 읽기 (파일 경로는 실제 파일 위치에 맞게 수정)
    df_raw = pd.read_excel('자회사A_FS.xlsx', sheet_name="CE" , header=0)

    # --- 1. 지분율 정보 추출 (수정된 로직) ---
    # 첫 행, 두 번째 열에서 원본 값 가져오기
    raw_value = df_raw.iloc[0, 1]

    # 원본 값을 문자열로 변환하여 '%' 포함 여부 확인
    if '%' in str(raw_value):
        # '%'가 포함된 문자열이면, '%'를 제거하고 100으로 나눔 (예: '80%')
        ownership_percentage = float(str(raw_value).replace('%', '').strip()) / 100
    else:
        # '%'가 없으면, 소수점 값으로 간주하고 그대로 숫자로 변환 (예: 0.6 또는 '0.6')
        ownership_percentage = float(raw_value)

    print("✅ 지분율 정보")
    print(f"  - 원본 값: '{raw_value}'")
    print(f"  - 최종 변환된 값: {ownership_percentage}\n")


    # --- 2. 자본 계정 코드 추출 ---
    account_codes = df_raw.iloc[0, 3:].astype(str).tolist()
    
    print("✅ 자본 계정 코드")
    print(f"  - 추출된 코드 리스트: {account_codes}")

    # --- 3. 연결 조정 코드 추출 ---
    caje_codes = df_raw.iloc[1, 3:].astype(str).tolist()
    
    print("✅ 비지배지분 조정 코드")
    print(f"  - 추출된 코드 리스트: {caje_codes}")



except FileNotFoundError:
    print("오류: 파일을 찾을 수 없습니다.")
except Exception as e:
    print(f"오류가 발생했습니다: {e}")
    



excel_files = ['모회사_FS.xlsx','자회사A_FS.xlsx','자회사B_FS.xlsx']
df_list = []
for file in excel_files:
        try:
            df = pd.read_excel(file, sheet_name="CE", header=1)
            df_list.append(df)
        except Exception as e:
            print("파일 처리 중 오류 발생: {e}")



# 비지배지분 배부
'''
비지배지분 배부 로직 (엑셀 파일에서 지분율을 추출합니다.
비지배지분율을 계산합니다 (1 - 지분율).
네 번째 열부터 비지배지분 열(가장 오른쪽 열) 전까지를 계산 대상으로 지정합니다.
각 행을 순회하며 계산 대상 열의 각 값에 대해 비지배지분 해당액을 계산합니다.
원본 값에서 비지배지분 해당액을 차감합니다.
해당 행에서 발생한 비지배지분 해당액의 총합을 가장 오른쪽 '비지배지분' 열에 더해줍니다.
연결조정에서 생성한 비지배지분과 이익잉여금 대체 행을 자본변동표에 추가합니다.
'''
# sheet_name="CE"

filename = '자회사B_FS.xlsx'
try:
    # --- 1. 데이터 준비 및 지분율 추출 ---
    df_raw = pd.read_excel(filename, sheet_name="CE", header=0)
    raw_ownership_value = df_raw.iloc[0, 1]
    if '%' in str(raw_ownership_value):
        ownership_percentage = float(str(raw_ownership_value).replace('%', '').strip()) / 100
    else:
        ownership_percentage = float(raw_ownership_value)
    nci_percentage = 1 - ownership_percentage
    print(f"✅ 지분율: {ownership_percentage:.2%}, 비지배지분율: {nci_percentage:.2%}")

    df = pd.read_excel(filename, sheet_name="CE", header=[1, 2])
    df.columns = ['회사명', '구분', '조정코드'] + [f'계정_{col[1]}' for col in df.columns[3:]]
    
    # 'Beginning', 'Ending' 행 제거
    df = df[~df['조정코드'].isin(['Beginning', 'Ending'])].copy()
    print("✅ 데이터 로딩 및 'Beginning', 'Ending' 행 제거 완료")

    # --- 2. 계산 영역 설정 ---
    nci_col_name = df.columns[-1]
    calculation_cols = df.columns[3:-1]
    print(f" - 비지배지분 열: '{nci_col_name}'")
    print(f" - 계산 대상 열: {list(calculation_cols)}")

    # --- 3. 데이터 타입 검사, 경고 및 변환 ---
    # (경고 기능 포함)
    print("\n--- 3. 데이터 타입 검사 및 변환 시작 ---")
    conversion_warnings = []
    columns_to_check = list(calculation_cols) + [nci_col_name]
    for col in columns_to_check:
        original_na_mask = df[col].isna()
        numeric_series = pd.to_numeric(df[col], errors='coerce')
        failed_mask = numeric_series.isna() & ~original_na_mask
        if failed_mask.any():
            for index in df.index[failed_mask]:
                original_value = df.loc[index, col]
                excel_row_num = index + 3
                warning_msg = (f"  - [경고] 열 '{col}', 엑셀 {excel_row_num}행의 값 "
                               f"'{original_value}'는 숫자가 아니므로 계산 시 0으로 처리됩니다.")
                conversion_warnings.append(warning_msg)
        df[col] = numeric_series.fillna(0)
    if conversion_warnings:
        print("\n⚠️  주의: 일부 데이터가 숫자가 아니므로 0으로 자동 변환되었습니다.")
        for msg in conversion_warnings:
            print(msg)
    else:
        print("✅ 모든 계산 열이 유효한 숫자 타입임을 확인했습니다.")

    
    
    # --- 4. 핵심 재배부 로직 및 기록 ---
    # 4-1. 각 항목별 비지배지분 해당액을 계산 
    row_sums = df[calculation_cols].sum(axis=1)
    total_nci_per_row = row_sums * nci_percentage
    safe_row_sums = row_sums.replace(0, 1)
    weights = df[calculation_cols].div(safe_row_sums, axis=0)
    nci_distribution = weights.mul(total_nci_per_row, axis=0)

    # 4-2. (기록용) '조정코드'와 '항목별 비지배지분액'을 합쳐 새로운 데이터프레임 생성
    # '조정코드' 열을 인덱스로 설정
    nci_log = nci_distribution.copy()
    nci_log['조정코드'] = df['조정코드']
    
    # 4-3. 비지배지분 계산 및 반영
    df[nci_col_name] += total_nci_per_row
    df[calculation_cols] -= nci_distribution
    print("\n✅ 비지배지분 재배부 계산을 완료했습니다.")

    # --- 5. 최종 데이터 정제 및 결과 저장 ---
    # 5-1. 모든 숫자 열이 0인 행 제거
    numeric_cols = list(calculation_cols) + [nci_col_name]
    all_zero_condition = (df[numeric_cols].round(2) == 0).all(axis=1) # 반올림하여 비교
    final_df = df[~all_zero_condition]
    print(f"✅ 모든 숫자 열의 값이 0인 행 {all_zero_condition.sum()}개를 제거했습니다.")

    # 5-2. 비지배지분 배부 내역 요약
    # '조정코드'를 기준으로 각 계정별 비지배지분액을 합산
    nci_summary = nci_log.groupby('조정코드').sum().reset_index()
    # 합계가 0인 행은 요약에서 제외
    nci_summary = nci_summary.loc[(nci_summary.iloc[:, 1:] != 0).any(axis=1)]
    print("✅ 조정코드별 비지배지분 배부 내역을 요약했습니다.")

    # 5-3. 결과를 여러 시트에 나누어 하나의 엑셀 파일로 저장
    output_filename = '자본변동표_계산결과_상세.xlsx'
    with pd.ExcelWriter(output_filename) as writer:
        final_df.to_excel(writer, sheet_name='자본변동표_계산결과', index=False)
        nci_summary.to_excel(writer, sheet_name='비지배지분_배부내역', index=False)
    
    print(f"\n🎉 모든 작업 완료! 결과가 '{output_filename}' 파일로 저장되었습니다.")
    print("   - '자본변동표_계산결과': 최종 자본변동표")
    print("   - '비지배지분_배부내역': 조정코드별 비지배지분 배부 상세 내역")
    
    

except FileNotFoundError:
    print(f"오류: '{filename}' 파일을 찾을 수 없습니다.")
except Exception as e:
    print(f"처리 중 오류가 발생했습니다: {e}")
    
    
