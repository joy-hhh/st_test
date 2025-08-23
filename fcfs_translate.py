import pandas as pd
import os

# =========================
# 설정 (필요시 바꿔서 사용)
# =========================
AMOUNT_COL_CANDIDATES = ("외화금액", "FC_Amount", "Amount")
EQUITY_CARRY_COL = "이월금액"             # E/RE 전기말 환산잔액(이월금액)
NAME_COL_CANDIDATES = ("계정명", "Account", "Name")
RE_NEW_NAME = "이월이익잉여금(환산)"       # RE 행이 없을 때 생성 시 이름
EPS_BS = 1e-6                             # 차대검증 허용 오차


# ============
# 유틸 함수들
# ============
def _first_numeric_in_row(row):
    s = pd.to_numeric(row, errors="coerce").dropna()
    return None if s.empty else float(s.iloc[0])

def _find_col(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None


# =======================
# 1) 파일 읽기 + 환율 추출
# =======================
def read_rates_and_table(xlsx_path):
    """
    전제:
    - 첫 번째 행: 헤더
    - 그 다음 2행: [기말환율, 평균환율]
    - 4행부터 실제 데이터
    """
    all_df = pd.read_excel(xlsx_path, header=0)

    closing_rate = _first_numeric_in_row(all_df.iloc[0])  # 기말환율
    average_rate = _first_numeric_in_row(all_df.iloc[1])  # 평균환율
    if closing_rate is None or average_rate is None:
        raise ValueError("기말/평균환율을 2~3행(데이터 첫 2행)에서 찾지 못했습니다.")

    # 환율 2줄 제거 → 본문 데이터만 남김
    df = all_df.drop(index=[0, 1]).reset_index(drop=True)

    if "FS_Element" not in df.columns:
        raise ValueError("파일에 FS_Element 컬럼이 없습니다. (A/L/E/RE/R/X/PI)")

    return closing_rate, average_rate, df


# ====================================
# 2) 사전 체크 (외화 기준 A-L-E, NI_FC)
# ====================================
def precheck_foreign_currency(df, eps=EPS_BS):
    """
    업로드 원시 '외화 금액'으로 먼저 검증:
      - BS: A - L - (E+RE) == 0 ?
      - IS: NI_FC = ΣR - ΣX  (X는 외화 원시금액이 양(+))이라고 가정하고 빼줌
    """
    df = df.copy()
    amount_col = _find_col(df, AMOUNT_COL_CANDIDATES)
    if amount_col is None:
        raise ValueError(f"외화금액 컬럼을 찾지 못했습니다. 후보={AMOUNT_COL_CANDIDATES}")

    df[amount_col] = pd.to_numeric(df[amount_col], errors="coerce").fillna(0.0)

    is_A  = df["FS_Element"].eq("A")
    is_L  = df["FS_Element"].eq("L")
    is_E  = df["FS_Element"].eq("E")
    is_RE = df["FS_Element"].eq("RE")
    is_R  = df["FS_Element"].eq("R")
    is_X  = df["FS_Element"].eq("X")

    a_fc  = df.loc[is_A,  amount_col].sum()
    l_fc  = df.loc[is_L,  amount_col].sum()
    e_fc  = df.loc[is_E | is_RE, amount_col].sum()   # 자본은 E+RE
    ni_fc = df.loc[is_R, amount_col].sum() - df.loc[is_X, amount_col].sum()

    bs_gap_fc = a_fc - l_fc - e_fc

    print("[PRECHECK] (외화) A-L-(E+RE) =", bs_gap_fc, "->", "OK" if abs(bs_gap_fc) < eps else "NG")
    print("[PRECHECK] (외화) NI_FC =", ni_fc)

    return {
        "A_FC": a_fc,
        "L_FC": l_fc,
        "E_plus_RE_FC": e_fc,
        "NI_FC": ni_fc,
        "BS_GAP_FC": bs_gap_fc,
        "BS_OK_FC": abs(bs_gap_fc) < eps,
    }


# ======================================================
# 3) 환산 + 당기순이익을 RE에 합산 + PI Plug-in + 사후검증
# ======================================================
def translate_fcfs(
    df,
    closing_rate,
    average_rate,
    eps=EPS_BS
):
    """
    규칙:
      - A/L: 기말환율로 환산
      - E/RE: '이월금액'(전기말 환산금액) 그대로 사용(환산 안 함)
      - R/X: 평균환율로 환산 (단, NI 계산 시 X는 차감)
      - NI_KRW = Σ(R_환산) - Σ(X_환산)
      - NI_KRW를 RE(이월이익잉여금)의 '환산금액'에 **합산**
      - PI = A - L - (E + RE)  (RE에는 이미 NI 포함)
      - 사후검증: A - L - (E + RE + PI) == 0 ?
    """
    df = df.copy()

    amount_col = _find_col(df, AMOUNT_COL_CANDIDATES)
    if amount_col is None:
        raise ValueError(f"외화금액 컬럼을 찾지 못했습니다. 후보={AMOUNT_COL_CANDIDATES}")

    name_col = _find_col(df, NAME_COL_CANDIDATES)

    # 숫자화
    df[amount_col] = pd.to_numeric(df[amount_col], errors="coerce").fillna(0.0)
    if EQUITY_CARRY_COL in df.columns:
        df[EQUITY_CARRY_COL] = pd.to_numeric(df[EQUITY_CARRY_COL], errors="coerce").fillna(0.0)

    out_col = "금액"
    if out_col not in df.columns:
        df[out_col] = 0.0

    is_A  = df["FS_Element"].eq("A")
    is_L  = df["FS_Element"].eq("L")
    is_E  = df["FS_Element"].eq("E")
    is_RE = df["FS_Element"].eq("RE")
    is_R  = df["FS_Element"].eq("R")
    is_X  = df["FS_Element"].eq("X")
    is_PI = df["FS_Element"].eq("PI")

    # 1) A/L: 기말환율 환산
    df.loc[is_A | is_L, out_col] = df.loc[is_A | is_L, amount_col] * closing_rate

    # 2) E/RE: 이월금액(전기말 환산액) 그대로 사용
    df.loc[is_E | is_RE, out_col] = df[EQUITY_CARRY_COL] if EQUITY_CARRY_COL in df.columns else 0.0

    # 3) R/X: 평균환율 (X는 화면표시를 위해 양수로)
    df.loc[is_R, out_col] = df.loc[is_R, amount_col] * average_rate
    df.loc[is_X, out_col] = df.loc[is_X, amount_col] * average_rate

    # 4) NI 계산(R-X) 후 RE에 합산
    ni_krw = df.loc[is_R, out_col].sum() - df.loc[is_X, out_col].sum()
    if is_RE.any():
        re_idxs = df.index[is_RE]
        df.loc[re_idxs[0], out_col] = df.loc[re_idxs[0], out_col] + ni_krw
        # 나머지 RE 행은 그대로 둔다
    else:
        # RE 행이 없으면 생성
        new_row = {col: None for col in df.columns}
        new_row["FS_Element"] = "RE"
        new_row[out_col] = ni_krw
        new_row[amount_col] = 0.0
        if EQUITY_CARRY_COL in df.columns:
            new_row[EQUITY_CARRY_COL] = 0.0
        if name_col is not None:
            new_row[name_col] = RE_NEW_NAME
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        is_RE = df["FS_Element"].eq("RE")  # 마스크 갱신

    # 5) PI Plug-in: A - L - (E + RE)
    assets_sum = df.loc[df["FS_Element"].eq("A"), out_col].sum()
    liabs_sum  = df.loc[df["FS_Element"].eq("L"), out_col].sum()
    equity_sum = df.loc[df["FS_Element"].isin(["E", "RE"]), out_col].sum()

    diff = assets_sum - liabs_sum - equity_sum  # 이 값을 PI에 꽂음

    if is_PI.any():
        pi_idxs = df.index[is_PI]
        # 여러 개면 첫 번째에만 설정, 나머지는 0
        df.loc[pi_idxs, out_col] = 0.0
        df.loc[pi_idxs[0], out_col] = diff
    else:
        new_row = {col: None for col in df.columns}
        new_row["FS_Element"] = "PI"
        new_row[out_col] = diff
        new_row[amount_col] = 0.0
        if EQUITY_CARRY_COL in df.columns:
            new_row[EQUITY_CARRY_COL] = 0.0
        if name_col is not None:
            new_row[name_col] = "해외사업환산손익(PI)"
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

    # 6) 사후 검증 (원화): A - L - (E + RE + PI) == 0 ?
    pi_krw = diff
    e_total_with_pi = df.loc[df["FS_Element"].isin(["E", "RE", "PI"]), out_col].sum()
    bs_gap_after = assets_sum - liabs_sum - e_total_with_pi

    print(f"[POSTCHECK] (환산 후) A-L-(E+RE+PI) = {bs_gap_after:.4f}  -> {'OK' if abs(bs_gap_after)<eps else 'NG'}")
    print(f"[POSTCHECK] A={assets_sum:.2f}, L={liabs_sum:.2f}, (E+RE+PI)={e_total_with_pi:.2f}, "
          f"NI(from R&X)={ni_krw:.2f}, PI={pi_krw:.2f}")

    totals = {
        "A(KRW)": assets_sum,
        "L(KRW)": liabs_sum,
        "E_plus_RE_plus_PI(KRW)": e_total_with_pi,
        "NI(KRW from R&X)": ni_krw,
        "PI(KRW)": pi_krw,
        "A-L-(E+RE+PI) (after)": bs_gap_after
    }

    # 주요 숫자 컬럼(`외화금액`, `이월금액`, `금액`)이 모두 0인 행을 제거
    cols_to_check = [amount_col, out_col]
    if EQUITY_CARRY_COL in df.columns:
        cols_to_check.append(EQUITY_CARRY_COL)

    is_zero_row = (df[cols_to_check].fillna(0) == 0).all(axis=1)
    df = df[~is_zero_row].reset_index(drop=True)

    return df, totals


# ==========================
# 4) 엔드투엔드 실행 + 저장
# ==========================
def process_fcfs_file(
    xlsx_path="FCFS(Foreign Currency Financial Statements).xlsx",
    output_path=None
):
    closing_rate, average_rate, df = read_rates_and_table(xlsx_path)

    # 1) 사전(외화) 체크
    pre = precheck_foreign_currency(df, eps=EPS_BS)

    # 2) 환산 + RE에 NI 합산 + PI Plug-in + 사후검증
    out_df, totals = translate_fcfs(df, closing_rate, average_rate, eps=EPS_BS)

    # 3) 저장 경로
    if output_path is None:
        stem = os.path.splitext(os.path.basename(xlsx_path))[0]
        dir_name = os.path.dirname(xlsx_path)
        output_path = os.path.join(dir_name, f"{stem}_translated.xlsx")

    # 4) 저장
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        # 데이터 쓰기
        out_df.to_excel(writer, index=False, sheet_name="translated")
        
        rates_summary = pd.DataFrame({
            "항목": ["기말환율", "평균환율"],
            "값": [closing_rate, average_rate]
        })
        
        summary = pd.concat([
            rates_summary,
            pd.DataFrame({"항목": list(pre.keys()), "값": list(pre.values())}),
            pd.DataFrame({"항목": list(totals.keys()), "값": list(totals.values())})
        ], ignore_index=True)
        summary.to_excel(writer, index=False, sheet_name="summary")

        # 서식 설정
        workbook = writer.book
        header_format = workbook.add_format({
            'bold': True,
            'fg_color': '#D9D9D9',  # Light grey
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })

        # 'translated' 시트 서식 적용
        worksheet_translated = writer.sheets['translated']
        for col_num, value in enumerate(out_df.columns.values):
            worksheet_translated.write(0, col_num, value, header_format)
        worksheet_translated.set_column(0, len(out_df.columns) - 1, 17)

        # 'summary' 시트 서식 적용
        worksheet_summary = writer.sheets['summary']
        for col_num, value in enumerate(summary.columns.values):
            worksheet_summary.write(0, col_num, value, header_format)
        worksheet_summary.set_column(0, len(summary.columns) - 1, 17)

    return output_path


# ======
# 예시
# ======
result_path = process_fcfs_file()
print("저장:", result_path)




