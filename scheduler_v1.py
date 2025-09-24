# scheduler_v1.py
import pandas as pd

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

r_rest = int(opt_map.get("최소휴식슬롯", 1))  # r=1부터 시작
if r_rest < 1:
    r_rest = 1

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

# 3) 고정 슬롯 우선 배치
slots = [None]*N          # 1..N → 0..N-1
used  = set()             # 이미 배치된 무대 이름
for x in rows:
    if x["fixed"] is not None and 1 <= x["fixed"] <= N:
        idx = x["fixed"]-1
        if slots[idx] is not None:
            raise ValueError(f"[에러] 슬롯 {x['fixed']} 충돌: {slots[idx]} vs {x['name']}")
        slots[idx] = x["name"]
        used.add(x["name"])

# 4) 충돌 검사 함수 (최근 r_rest 슬롯의 출연자와 겹치면 안 됨)
def violates_rest(candidate_performers, schedule, pos, r_rest):
    # 최근 r_rest 슬롯의 모든 출연자 집합
    recent_perf = set()
    for back in range(1, r_rest+1):
        prev_idx = pos - back
        if prev_idx < 0: 
            break
        sname = schedule[prev_idx]
        if sname is None: 
            continue
        recent_perf.update(name_to_row[sname]["performers"])
    # 겹치면 금지
    return any(p in recent_perf for p in candidate_performers)

# 5) 백트래킹
def backtrack(i, schedule, used):
    if i == N:
        return True
    # 이미 고정된 슬롯이면 다음으로
    if schedule[i] is not None:
        return backtrack(i+1, schedule, used)

    # 후보 무대: 아직 안 쓴 것들
    for x in rows:
        nm = x["name"]
        if nm in used: 
            continue
        # 이 무대가 다른 슬롯에 고정되어 있지는 않은지(=이미 배치되었을 것)
        if x["fixed"] is not None:
            continue
        # 휴식 제약(r_rest) 위반 검사
        if violates_rest(x["performers"], schedule, i, r_rest):
            continue

        # 배치해 보고
        schedule[i] = nm
        used.add(nm)
        if backtrack(i+1, schedule, used):
            return True
        # 실패 시 되돌리기
        used.remove(nm)
        schedule[i] = None

    # 어떤 후보도 안 되면 실패
    return False

ok = backtrack(0, slots, set(used))

print(f"옵션: 최소휴식슬롯(r)={r_rest}")
print("=== 스케줄(왼쪽부터 1번 슬롯) ===")
for i, s in enumerate(slots, start=1):
    print(f"{i:>2} : {s}")

if not ok:
    print("\n❌ 스케줄을 찾지 못했습니다.")
    print("   - r 값을 1로 낮추거나(옵션 시트),")
    print("   - 참가자 겹침을 줄이거나,")
    print("   - 고정순서를 조정한 뒤 다시 시도하세요.")
else:
    print("\n🎉 스케줄 생성 성공!")

    # 6) 결과를 엑셀로 저장(새 시트)
    out = []
    for i, s in enumerate(slots, start=1):
        row = name_to_row[s]
        out.append({
            "슬롯": i,
            "무대": s,
            "길이(초)": row["duration"],
            "참가자": ", ".join(row["performers"])
        })
    out_df = pd.DataFrame(out)

    # 기존 파일에 시트 추가/갱신
    with pd.ExcelWriter(INPUT, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
        out_df.to_excel(w, sheet_name="생성결과", index=False)
    print("✅ 엑셀 시트 '생성결과'에 저장 완료")
