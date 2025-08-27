import pandas as pd

coa_df = pd.read_excel("CoA_Level.xlsx")
fs_map = dict(zip(coa_df["계정코드"], coa_df["FS_Element"]))

[key for key, value in fs_map.items() if value == "CE"][0]

aje_code = pd.read_excel("CoA_Level.xlsx", "AJE", dtype=str)
aje_code[aje_code["FS_Element"] == "L"]["계정코드"][0]
aje_code[aje_code["FS_Element"] == "L"]["계정명"][0]

info = pd.read_excel("조정분개_입력템플릿_BeforeTaxNci (1).xlsx", sheet_name="Info")
info_df = info.set_index(info.columns[0])


def parse_percent(s):
    """
    다양한 형태의 퍼센트 값을 소수점 형태로 변환합니다.
    - '60%': 0.6
    - 60: 0.6 (1보다 크므로 퍼센트로 간주)
    - 0.6: 0.6 (1보다 작거나 같으므로 소수점으로 간주)
    """
    # 1. 입력값이 문자열일 경우
    if isinstance(s, str):
        try:
            # 문자열은 항상 '%'가 있거나 퍼센트 숫자로 간주하고 100으로 나눔
            return float(s.strip().strip('%')) / 100
        except (ValueError, TypeError):
            # "hello" 같이 변환 불가능한 문자열은 0.0 처리
            return 0.
    # 2. 입력값이 숫자(int, float)일 경우
    elif isinstance(s, (int, float)):
        # 숫자의 절댓값이 1보다 크면 (e.g., 60, -50) 퍼센트로 간주하고 100으로 나눔
        if abs(s) > 1:
            return float(s) / 100
        # 숫자의 절댓값이 1보다 작거나 같으면 (e.g., 0.6, -0.5, 1) 이미 변환된 소수점으로 간주하고 그대로 반환
        else:
            return float(s)

    # 3. 그 외 타입은 0.0 반환
    else:
        return 0.0



info_df["당기세율_num"] = info_df["당기세율"].apply(parse_percent)
tax_rates = info_df["당기세율_num"].to_dict()
tax_rate = tax_rates.get("자회사A", 0.0)

