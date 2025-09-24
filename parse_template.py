# parse_template.py
import pandas as pd
import math

INPUT = "타임테이블_템플릿.xlsx"

def _to_list(cell):
    if pd.isna(cell) or str(cell).strip() == "":
        return []
    # 쉼표로 자르고 공백 제거
    return [p.strip() for p in str(cell).split(",") if p.strip()]

def _to_int_or_none(cell):
    if pd.isna(cell) or str(cell).strip() == "":
        return None
    try:
        v = int(float(cell))  # 2.0 같은 것도 2로 처리
        return v
    except:
        return None

# 1) 읽기
stages = pd.read_excel(INPUT, sheet_name="무대")
options = pd.read_excel(INPUT, sheet_name="옵션")

# 2) 컬럼 존재/타입 점검(필수 컬럼 체크)
required_cols = ["이름", "길이(초)", "참가자", "고정순서"]
missing = [c for c in required_cols if c not in stages.columns]
if missing:
    raise ValueError(f"무대 시트에 컬럼 누락: {missing}")

# 3) 정규화(참가자 리스트화, 고정순서 정수화)
normalized = []
for _, row in stages.iterrows():
    name = str(row["이름"]).strip()
    dur = row["길이(초)"]
    performers = _to_list(row["참가자"])
    fixed = _to_int_or_none(row["고정순서"])
    if pd.isna(dur):
        raise ValueError(f"무대 '{name}'의 길이(초)가 비어있음")
    try:
        dur = int(float(dur))
    except:
        raise ValueError(f"무대 '{name}'의 길이(초) 숫자 형식 오류: {dur}")

    normalized.append({
        "name": name,
        "duration": dur,
        "performers": performers,
        "fixed_order": fixed,
    })

# 4) 간단 검증(이름 중복, 참가자 오타 의심 등)
names = [x["name"] for x in normalized]
dup_names = sorted(set([n for n in names if names.count(n) > 1]))

# 참가자 중복 출연(동일 슬롯 금지는 스케줄 단계에서 처리하지만, 여기서는 빈도만 점검)
from collections import Counter
perf_all = [p for x in normalized for p in x["performers"]]
perf_count = Counter(perf_all)
heavy_perf = sorted([p for p,c in perf_count.items() if c >= 3])  # 3회 이상 출연자

# 고정순서 범위 경고(1~무대수)
n = len(normalized)
fixed_out_of_range = [x for x in normalized if (x["fixed_order"] is not None and not (1 <= x["fixed_order"] <= n))]

# 5) 옵션 파싱
opt_map = {}
for _, row in options.iterrows():
    k = str(row["옵션명"]).strip()
    v = row["값"]
    opt_map[k] = v
rest_min_slots = int(opt_map.get("최소휴식슬롯", 1))
num_candidates = int(opt_map.get("후보안개수", 5))
rest_seconds = int(opt_map.get("쉬는시간(초)", 60))

# 6) 요약 출력
print("✅ 무대 수:", n)
print("✅ 총 참가자 수(중복 포함):", len(perf_all))
print("✅ 고정순서 지정된 무대 수:", sum(1 for x in normalized if x["fixed_order"] is not None))
if dup_names:
    print("⚠️ 무대 이름 중복:", dup_names)
if heavy_perf:
    print("⚠️ 3회 이상 출연자:", heavy_perf)
if fixed_out_of_range:
    print("⚠️ 고정순서 범위 벗어남:", [(x['name'], x['fixed_order']) for x in fixed_out_of_range])

print("\n옵션:", {"최소휴식슬롯": rest_min_slots, "후보안개수": num_candidates, "쉬는시간(초)": rest_seconds})

# 7) 파싱된 첫 3개 미리보기
for i, x in enumerate(normalized[:3], start=1):
    print(f"\n[{i}] {x['name']} / {x['duration']}초 / fixed={x['fixed_order']}")
    print("   참가자:", ", ".join(x["performers"]) if x["performers"] else "(없음)")
