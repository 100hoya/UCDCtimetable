# scheduler_v3_scoring.py
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

# -------------------- 데이터 읽기 --------------------
stages_df = pd.read_excel(INPUT, sheet_name="무대")
opts_df   = pd.read_excel(INPUT, sheet_name="옵션")

opt_map = {}
for _, r in opts_df.iterrows():
    opt_map[str(r["옵션명"]).strip()] = r["값"]

r_rest = int(opt_map.get("최소휴식슬롯", 1))
num_candidates = int(opt_map.get("후보안개수", 5))

if r_rest < 1: r_rest = 1
if num_candidates < 1: num_candidates = 1

rows = []
for _, r in stages_df.iterrows():
    name = str(r["이름"]).strip()
    dur  = int(float(r["길이(초)"]))
    perf = to_list(r["참가자"])
    fixed = to_int_or_none(r["고정순서"])
    rows.append({"name": name, "duration": dur, "performers": perf, "fixed": fixed})

N = len(rows)
name_to_row = {x["name"]: x for x in rows}

# -------------------- 스케줄링 함수 --------------------
def initial_slots():
    slots = [None]*N
    used = set()
    for x in rows:
        if x["fixed"] is not None and 1 <= x["fixed"] <= N:
            idx = x["fixed"] - 1
            slots[idx] = x["name"]
            used.add(x["name"])
    return slots, used

def violates_rest(candidate_performers, schedule, pos, r_rest):
    recent = set()
    for back in range(1, r_rest+1):
        j = pos - back
        if j < 0: break
        sname = schedule[j]
        if sname is None: continue
        recent.update(name_to_row[sname]["performers"])
    return any(p in recent for p in candidate_performers)

def solve_with_seed(seed):
    random.seed(seed)
    slots, used = initial_slots()
    pool = [x for x in rows if x["fixed"] is None]
    random.shuffle(pool)

    def backtrack(i):
        if i == N:
            return True
        if slots[i] is not None:
            return backtrack(i+1)
        for x in pool:
            nm = x["name"]
            if nm in used: continue
            if violates_rest(x["performers"], slots, i, r_rest): continue
            slots[i] = nm
            used.add(nm)
            if backtrack(i+1): return True
            used.remove(nm)
            slots[i] = None
        return False

    ok = backtrack(0)
    return ok, slots

# -------------------- 채점 함수 --------------------
def score_schedule(slots):
    score = 0
    details = []

    # 파라미터(쉽게 조절 가능)
    NEAR_REPEAT_WINDOW = 2   # 이 창 안에 같은 참가자 재등장 → 감점
    PENALTY_NEAR_REPEAT = 2  # 가까운 재등장 감점 크기(기존보다 낮춤)
    PENALTY_LONG_LONG = 2    # 긴 무대끼리 연속 감점(기존 5 → 2로 완화)
    LONG_THRESHOLD = 200

    BONUS_SPREAD = 1         # 충분히 띄워졌을 때 소보너스
    BONUS_ALT_LS = 1         # 긴/짧은 번갈음 보너스

    # 1) 참가자 분산도: 가까운 재등장은 감점, 충분히 띄우면 소보너스
    last_seen = {}
    for i, s in enumerate(slots):
        perf = name_to_row[s]["performers"]
        for p in perf:
            if p in last_seen:
                dist = i - last_seen[p]
                if dist <= NEAR_REPEAT_WINDOW:
                    score -= PENALTY_NEAR_REPEAT * (NEAR_REPEAT_WINDOW + 1 - dist)
                    details.append(f"- 참가자 {p} 근접 재등장 (슬롯 {last_seen[p]+1}->{i+1})")
                elif dist >= NEAR_REPEAT_WINDOW + 2:
                    score += BONUS_SPREAD
            last_seen[p] = i

    # 2) 무대 길이 균형: 긴 무대가 연속되면 감점, 번갈아 나오면 보너스
    def is_long(name): 
        return name_to_row[name]["duration"] >= LONG_THRESHOLD

    for i in range(1, len(slots)):
        prev_long = is_long(slots[i-1])
        cur_long  = is_long(slots[i])
        if prev_long and cur_long:
            score -= PENALTY_LONG_LONG
            details.append(f"- 긴무대 연속 (슬롯 {i}/{i+1})")
        # 번갈아 나오면 소보너스
        if prev_long != cur_long:
            score += BONUS_ALT_LS

    return score

# -------------------- 여러 후보안 생성 --------------------
results = []
seen = set()
seed0 = 9999

for k in range(num_candidates*3):
    if len(results) >= num_candidates: break
    ok, sched = solve_with_seed(seed0+k)
    if not ok: continue
    key = tuple(sched)
    if key in seen: continue
    seen.add(key)
    sc = score_schedule(sched)
    results.append((sc, sched))

# -------------------- 결과 출력 및 저장 --------------------
results.sort(reverse=True, key=lambda x: x[0])

if results:
    with pd.ExcelWriter(INPUT, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
        for idx, (sc, sched) in enumerate(results, start=1):
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
            df.to_excel(w, sheet_name=f"스코어_{idx}", index=False)
    print("✅ 저장 완료: 스코어 순으로 시트 작성")
    for i, (sc, _) in enumerate(results, start=1):
        print(f"스코어_{i}: 점수 {sc}")
else:
    print("❌ 후보안 생성 실패")
