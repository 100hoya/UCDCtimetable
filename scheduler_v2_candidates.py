# scheduler_v2_candidates.py
import pandas as pd
import random

INPUT = "타임테이블_템플릿.xlsx"

def to_list(cell):
    if pd.isna(cell) or str(cell).strip() == "":
        return []
    return [p.strip() for p in str(cell).split(",") if p.strip()]

def to_int_or_none(x):
    if pd.isna(x) or str(x).strip() == "":
        return None
    try:
        return int(float(x))
    except:
        return None

# 1) 엑셀 읽기
stages_df = pd.read_excel(INPUT, sheet_name="무대")
opts_df   = pd.read_excel(INPUT, sheet_name="옵션")

# 옵션 파싱
opt_map = {}
for _, r in opts_df.iterrows():
    opt_map[str(r["옵션명"]).strip()] = r["값"]

r_rest = int(opt_map.get("최소휴식슬롯", 1))
num_candidates = int(opt_map.get("후보안개수", 5))
if r_rest < 1: r_rest = 1
if num_candidates < 1: num_candidates = 1

# 2) 정규화
rows = []
for _, r in stages_df.iterrows():
    name = str(r["이름"]).strip()
    dur  = r["길이(초)"]
    if pd.isna(dur):
        raise ValueError(f"[에러] '{name}' 길이(초) 비어있음")
    try:
        dur = int(float(dur))
    except:
        raise ValueError(f"[에러] '{name}' 길이(초) 숫자 아님: {dur}")
    perf = to_list(r["참가자"])
    fixed = to_int_or_none(r["고정순서"])
    rows.append({"name": name, "duration": dur, "performers": perf, "fixed": fixed})

N = len(rows)
name_to_row = {x["name"]: x for x in rows}

# 3) 고정 슬롯 미리 배치
def initial_slots():
    slots = [None]*N
    used = set()
    for x in rows:
        if x["fixed"] is not None and 1 <= x["fixed"] <= N:
            idx = x["fixed"] - 1
            if slots[idx] is not None:
                raise ValueError(f"[에러] 슬롯 {x['fixed']} 충돌: {slots[idx]} vs {x['name']}")
            slots[idx] = x["name"]
            used.add(x["name"])
    return slots, used

# 4) 휴식(r) 제약 검사
def violates_rest(candidate_performers, schedule, pos, r_rest):
    recent = set()
    for back in range(1, r_rest+1):
        j = pos - back
        if j < 0: break
        sname = schedule[j]
        if sname is None: continue
        recent.update(name_to_row[sname]["performers"])
    return any(p in recent for p in candidate_performers)

# 5) 백트래킹(후보 무대 순서 셔플 버전)
def solve_with_seed(seed):
    random.seed(seed)
    slots, used = initial_slots()

    # 시도 순서를 바꿔주기: (고정 없는 무대들만)
    pool = [x for x in rows if x["fixed"] is None]
    random.shuffle(pool)

    def backtrack(i):
        if i == N: 
            return True
        if slots[i] is not None:
            return backtrack(i+1)

        # 탐색 순서: 셔플된 pool에서 아직 안 쓴 것들
        for x in pool:
            nm = x["name"]
            if nm in used: 
                continue
            if violates_rest(x["performers"], slots, i, r_rest):
                continue
            # 배치
            slots[i] = nm
            used.add(nm)
            if backtrack(i+1):
                return True
            # 되돌리기
            used.remove(nm)
            slots[i] = None
        return False

    ok = backtrack(0)
    return ok, slots

# 6) 여러 후보안 생성
results = []
seen = set()  # 중복 스케줄 방지
seed0 = 12345

for k in range(num_candidates * 3):  # 여유 시도(중복이 나와서 못 채우는 경우 대비)
    if len(results) >= num_candidates:
        break
    ok, sched = solve_with_seed(seed0 + k)
    if not ok:
        continue
    key = tuple(sched)  # 중복 확인
    if key in seen:
        continue
    seen.add(key)
    results.append(sched)

print(f"옵션: r={r_rest}, 요청 후보안={num_candidates}, 생성={len(results)}")

# 7) 엑셀에 각 후보안을 개별 시트로 저장
if results:
    with pd.ExcelWriter(INPUT, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
        for idx, sched in enumerate(results, start=1):
            out = []
            for i, s in enumerate(sched, start=1):
                row = name_to_row[s]
                out.append({
                    "슬롯": i,
                    "무대": s,
                    "길이(초)": row["duration"],
                    "참가자": ", ".join(row["performers"])
                })
            df = pd.DataFrame(out)
            df.to_excel(w, sheet_name=f"생성결과_{idx}", index=False)
    print("✅ 엑셀에 시트들 저장 완료:", [f"생성결과_{i+1}" for i in range(len(results))])
else:
    print("❌ 어떤 후보안도 찾지 못했습니다. r 또는 고정/참가자 구성을 조정해 보세요.")
