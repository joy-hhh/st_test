import pandas as pd
from openpyxl import Workbook


# 엑셀 템플릿 다운로드
# 조정유형 정의
adjustment_types = [
    ("CAJE01_채권채무제거", "Intercompany Elimination"),
    ("CAJE02_미실현이익제거", "Unrealized Profit Elimination"),
    ("CAJE03_투자자본상계", "Investment-Equity Elimination"),
    ("CAJE04_배당조정", "Dividend Adjustment"),
    ("CAJE05_감가상각조정", "Depreciation Adjustment"),
    ("CAJE06_상각비조정", "Amortization Adjustment"),
    ("CAJE07_손익조정", "Profit & Loss Adjustment"),
    ("CAJE08_회계정책조정", "Accounting Policy Adjustment"),
    ("CAJE09_지분법조정", "Equity Method Adjustment"),
    ("CAJE10_공정가치조정", "Fair Value Adjustment"),
    ("CAJE99_기타조정", "Other Adjustment"),
]

# 공통 입력 양식
columns = ["법인1", "계정1", "금액1", "법인2", "계정2", "금액2", "설명"]

# Excel 파일 생성
with pd.ExcelWriter("조정분개_입력템플릿.xlsx", engine="openpyxl") as writer:
    for sheet_name, _ in adjustment_types:
        df = pd.DataFrame(columns=columns)
        df.to_excel(writer, sheet_name=sheet_name, index=False)


# 입력한 조정분개_입력템플릿 업로드 - 조정분개 산출 기능
def get_fs_sign(fs_element):
    """FS_Element에 따라 금액 부호 결정"""
    if fs_element in ["A", "X"]:
        return -1
    elif fs_element in ["L", "E", "R"]:
        return 1
    else:
        return 1  # 기본값은 양수 처리



coa_df = pd.read_excel('CoA_Level.xlsx', dtype=str)


def build_caje_from_excel(adjustment_file, coa_file, output_file):
    # CoA 파일 불러오기
    coa_df = pd.read_excel(coa_file, dtype=str)
    fs_map = dict(zip(coa_df["계정코드"], coa_df["FS_Element"]))

    xls = pd.ExcelFile(adjustment_file)
    all_entries = []

    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet, dtype={"계정1": str,"계정2": str }).fillna("")

        try:
            caje_type = sheet.split("_")[1]
        except IndexError:
            caje_type = sheet

        for _, row in df.iterrows():
            설명 = row.get("설명", "")

            # 법인1
            code1 = str(row["계정1"]).strip()
            if row["법인1"] and code1 and row["금액1"]:
                fs1 = fs_map.get(code1, "")
                try:
                    금액1 = float(str(row["금액1"]).replace(",", ""))
                    sign1 = get_fs_sign(fs1)
                    all_entries.append({
                        "조정유형": caje_type,
                        "법인": row["법인1"],
                        "계정": code1,
                        "금액": 금액1 * sign1,
                        "설명": 설명,
                        "FS": fs1
                    })
                except ValueError:
                    pass

            # 법인2
            code2 = str(row["계정2"]).strip()
            if row["법인2"] and code2 and row["금액2"]:
                fs2 = fs_map.get(code2, "")
                try:
                    금액2 = float(str(row["금액2"]).replace(",", ""))
                    sign2 = get_fs_sign(fs2)
                    all_entries.append({
                        "조정유형": caje_type,
                        "법인": row["법인2"],
                        "계정": code2,
                        "금액": 금액2 * sign2,
                        "설명": 설명,
                        "FS": fs2
                    })
                except ValueError:
                    pass

    # 결과 저장
    df_result = pd.DataFrame(all_entries)
    df_result.to_excel(output_file, index=False)
    print(f"✅ 저장 완료: {output_file}")



build_caje_from_excel('조정분개_입력템플릿.xlsx', 'CoA_Level.xlsx', 'CAJE.xlsx')