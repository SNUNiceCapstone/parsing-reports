import pandas as pd
import json

# 엑셀 파일 경로
file_path = "(전달용)NICE코드표_IFRS_제조_입력대상 재무제표.xlsx"

# 시트에서 필요한 열만 읽기
df = pd.read_excel(file_path, sheet_name="(00)제조", dtype=str)

# 공백 제거
df.columns = df.columns.str.strip()
df["보고서코드"] = df["보고서코드"].str.strip()
df["개별사용여부"] = df["개별사용여부"].str.strip()
df["개별검증산식"] = df["개별검증산식"].fillna("").astype(str).str.strip()
df["항목명"] = df["항목명"].str.strip()

# 11번 보고서코드 중 개별사용여부가 'O'인 항목만 추출
df = df[(df["보고서코드"] == "11") & (df["개별사용여부"] == "O")].copy()

# 대괄호 포함 항목코드 수집: 항목명에 '[' 또는 ']' 포함
bracket_codes = set(df[df["항목명"].str.contains(r"\[|\]", regex=True)]["항목코드"])

# 개별검증산식 반영 매핑
name_map = df.set_index("항목코드")["항목명"].to_dict()
formula_map = df.set_index("항목코드")["개별검증산식"].to_dict()

# 대괄호 포함 항목의 하위 항목까지 제거
def collect_all_descendants(code, visited=None):
    if visited is None:
        visited = set()
    visited.add(code)
    formula = formula_map.get(code, "")
    children = [token.strip(" +-") for token in formula.replace("-", "+").split("+") if token.strip(" +-") in name_map]
    for child in children:
        if child not in visited:
            visited.add(child)
            collect_all_descendants(child, visited)
    return visited

# 제거 대상 전체 집합
remove_codes = set()
for bc in bracket_codes:
    remove_codes |= collect_all_descendants(bc)

# 제거 대상 제외
df = df[~df["항목코드"].isin(remove_codes)].copy()

# 계층 구조 만들기
def build_hierarchy_path(code, visited=None):
    if visited is None:
        visited = set()
    if code in visited or code not in name_map:
        return None
    visited.add(code)
    for parent_code, formula in formula_map.items():
        tokens = [token.strip(" +-") for token in formula.replace("-", "+").split("+")]
        if code in tokens:
            parent_path = build_hierarchy_path(parent_code, visited.copy())
            if parent_path:
                return f"{parent_path}|{name_map[code]}"
    return name_map[code]

result = {}
for code in df["항목코드"]:
    path = build_hierarchy_path(code)
    if path:
        result[code] = {"label": path}

# JSON 저장
with open("parsed_11.json", "w", encoding="utf-8") as f:
    json.dump(result, f, ensure_ascii=False, indent=4)