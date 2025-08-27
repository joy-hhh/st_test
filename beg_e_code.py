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

    