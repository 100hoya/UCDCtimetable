# place_fixed.py
import pandas as pd

INPUT = "타임테이블_템플릿.xlsx"

def to_int_or_none(x):
    if pd.isna(x) or str(x).strip() == "":
        return None
    try:
        return int(float(x))
    except:
        return None

# 1) 읽기
stages = pd.read_excel(INPUT, sheet_name="무대")
options = pd.read_excel(INPUT, sheet_name="옵션")

# 2) 필수 컬럼 확인
required = ["이름", "길이(초)", "참가자", "고정순서"]
missing = [c for c in required if c not in stages.columns]
if missing:
    raise ValueError(f"[에러] 무대 시트에 컬럼 누락: {missing}")

# 3) 기본 정규화
rows = []
for _, r in stages.iterrows():
    name = str(r["이름"]).strip()
    dur = r["길이(초)"]
    if pd.isna(dur):
        raise ValueError(f"[에러] '{name}' 길이(초) 비어있음")
    try:
        dur = int(float(dur))
    except:
        raise ValueError(f"[에러] '{name}' 길이(초) 숫자 아님: {dur}")
    fixed = to_int_or_none(r["고정순서"])
    rows.append({"name": name, "duration": dur, "fixed": fixed})

N = len(rows)
print(f"✅ 무대 수: {N}")

# 4) 고정순서 유효성 1차 점검
bad_range = [(x["name"], x["fixed"]) for x in rows
             if x["fixed"] is not None and not (1 <= x["fixed"] <= N)]
if bad_range:
    print("⚠️ 범위 벗어난 고정순서:", bad_range)

# 5) 같은 무대에 고정값이 여러개로 들어간 상황(보통은 없음)이런건 엑셀 구조상 거의 없지만 방어
#   → 현재 구조에서는 한 행=한 무대이므로 스킵

# 6) 같은 슬롯을 여러 무대가 차지하는지(충돌) + 슬롯 배열 만들기
slots = [None] * N  # 1..N 슬롯을 0..N-1 인덱스로
conflicts = []
for x in rows:
    if x["fixed"] is None:
        continue
    if not (1 <= x["fixed"] <= N):
        continue  # 범위 밖은 이미 경고했으니 배치 생략
    idx = x["fixed"] - 1
    if slots[idx] is None:
        slots[idx] = x["name"]
    else:
        conflicts.append((x["name"], x["fixed"], "이미 배치됨:"+slots[idx]))

# 7) 결과 요약
fixed_count = sum(1 for x in rows if x["fixed"] is not None)
print(f"✅ 고정 배치 시도: {fixed_count}개")
print("✅ 현재 슬롯(고정만 반영):")
for i, s in enumerate(slots, start=1):
    print(f"  - {i}번: {s if s else '(비어있음)'}")

# 8) 불가능 리포트
problem = False
if bad_range:
    problem = True
if conflicts:
    problem = True
    print("⚠️ 슬롯 충돌 발생:")
    for c in conflicts:
        print("   ", c)

if not problem:
    print("\n🎉 고정 배치에 문제가 없습니다. 다음 단계(자동 채우기)로 진행 가능합니다.")
else:
    print("\n🛠️ 위 경고를 엑셀에서 먼저 고쳐주세요. (수정 후 다시 실행)")
