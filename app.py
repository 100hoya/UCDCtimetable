# app.py
import streamlit as st
import pandas as pd
import random
from io import BytesIO

st.set_page_config(page_title="타임테이블 자동 생성기", layout="wide")

# ---------- 유틸 ----------
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

# ---------- 스케줄러 핵심 ----------
def parse_excel(file):
    stages_df = pd.read_excel(file, sheet_name="무대")
    opts_df   = pd.read_excel(file, sheet_name="옵션")

    # 옵션
    opt_map = {}
    for _, r in opts_df.iterrows():
        opt_map[str(r["옵션명"]).strip()] = r["값"]
    r_rest = int(opt_map.get("최소휴식슬롯", 1))
    num_candidates = int(opt_map.get("후보안개수", 5))
    if r_rest < 1: r_rest = 1
    if num_candidates < 1: num_candidates = 1

    # 무대 정규화
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

    return rows, r_rest, num_candidates

def initial_slots(rows):
    N = len(rows)
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

def violates_rest(candidate_performers, schedule, pos, r_rest, name_to_row):
    recent = set()
    for back in range(1, r_rest+1):
        j = pos - back
        if j < 0: break
        sname = schedule[j]
        if sname is None: continue
        recent.update(name_to_row[sname]["performers"])
    return any(p in recent for p in candidate_performers)

def solve_with_seed(rows, r_rest, seed):
    random.seed(seed)
    name_to_row = {x["name"]: x for x in rows}
    N = len(rows)

    slots, used = initial_slots(rows)
    pool = [x for x in rows if x["fixed"] is None]
    random.shuffle(pool)

    def backtrack(i):
        if i == N:
            return True
        if slots[i] is not None:
            return backtrack(i+1)
        for x in pool:
            nm = x["name"]
            if nm in used: 
                continue
            if violates_rest(x["performers"], slots, i, r_rest, name_to_row):
                continue
            slots[i] = nm
            used.add(nm)
            if backtrack(i+1): 
                return True
            used.remove(nm)
            slots[i] = None
        return False

    ok = backtrack(0)
    return ok, slots, name_to_row

# 간단 스코어(조절 쉬운 파라미터)
def score_schedule(slots, name_to_row):
    score = 0
    NEAR_REPEAT_WINDOW = 2
    PENALTY_NEAR_REPEAT = 2
    PENALTY_LONG_LONG = 2
    LONG_THRESHOLD = 200
    BONUS_SPREAD = 1
    BONUS_ALT_LS = 1

    last_seen = {}
    for i, s in enumerate(slots):
        perf = name_to_row[s]["performers"]
        for p in perf:
            if p in last_seen:
                dist = i - last_seen[p]
                if dist <= NEAR_REPEAT_WINDOW:
                    score -= PENALTY_NEAR_REPEAT * (NEAR_REPEAT_WINDOW + 1 - dist)
                elif dist >= NEAR_REPEAT_WINDOW + 2:
                    score += BONUS_SPREAD
            last_seen[p] = i

    def is_long(name):
        return name_to_row[name]["duration"] >= LONG_THRESHOLD

    for i in range(1, len(slots)):
        a, b = slots[i-1], slots[i]
        if is_long(a) and is_long(b):
            score -= PENALTY_LONG_LONG
        if is_long(a) != is_long(b):
            score += BONUS_ALT_LS

    return score

def schedule_to_df(slots, name_to_row):
    out = []
    for i, s in enumerate(slots, start=1):
        row = name_to_row[s]
        out.append({
            "슬롯": i,
            "무대": s,
            "길이(초)": row["duration"],
            "참가자": ", ".join(row["performers"])
        })
    return pd.DataFrame(out)

def build_excel_bytes(sheets: dict):
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as w:
        for name, df in sheets.items():
            df.to_excel(w, sheet_name=name, index=False)
    bio.seek(0)
    return bio

# ---------- UI ----------
st.title("🎛️ 타임테이블 자동 생성기 (MVP)")
st.caption("엑셀(무대/옵션 시트) 업로드 → 고정/휴식 제약 반영 → 여러 후보안 생성 → 다운로드")

uploaded = st.file_uploader("엑셀 파일을 업로드하세요 (.xlsx)", type=["xlsx"])

with st.sidebar:
    st.header("옵션 (필요시 덮어쓰기)")
    r_override = st.number_input("최소휴식슬롯 r (빈칸=엑셀값)", min_value=1, step=1)
    n_override = st.number_input("후보안 개수 (빈칸=엑셀값)", min_value=1, step=1, value=5)
    seed0 = st.number_input("랜덤 시드 시작값", value=12345, step=1)
    st.caption("엑셀 '옵션' 시트의 값이 기본이며, 여기서 입력하면 덮어씁니다.")

if uploaded is not None:
    try:
        rows, r_rest_from_file, num_from_file = parse_excel(uploaded)
    except Exception as e:
        st.error(f"엑셀 파싱 오류: {e}")
        st.stop()

    N = len(rows)
    st.success(f"무대 {N}개 읽음")
    st.dataframe(pd.DataFrame(rows)[["name","duration","performers","fixed"]], use_container_width=True)

    # 옵션 적용
    r_rest = int(r_override) if r_override else r_rest_from_file
    num_candidates = int(n_override) if n_override else num_from_file

    generate = st.button("💡 후보안 생성하기")
    if generate:
        results = []
        seen = set()
        tries = num_candidates * 4  # 넉넉히 시도
        for k in range(tries):
            ok, sched, name_to_row = solve_with_seed(rows, r_rest, seed0 + k)
            if not ok: 
                continue
            key = tuple(sched)
            if key in seen:
                continue
            seen.add(key)
            sc = score_schedule(sched, name_to_row)
            results.append((sc, sched))
            if len(results) >= num_candidates:
                break

        if not results:
            st.error("어떤 후보안도 찾지 못했습니다. r을 낮추거나, 고정/참가자 구성을 조정해 보세요.")
        else:
            # 점수순 정렬 & 표시
            results.sort(reverse=True, key=lambda x: x[0])
            tabs = st.tabs([f"스코어 {i+1}: {sc}" for i,(sc,_) in enumerate(results)])
            excel_sheets = {}
            for i, (sc, sched) in enumerate(results):
                df = schedule_to_df(sched, name_to_row)
                excel_sheets[f"결과_{i+1}_점수_{sc}"] = df
                with tabs[i]:
                    st.write(f"**점수:** {sc}  |  **r:** {r_rest}")
                    st.dataframe(df, use_container_width=True)

            # 다운로드
            st.divider()
            st.subheader("📥 엑셀로 받기")
            bio = build_excel_bytes(excel_sheets)
            st.download_button(
                label="결과 엑셀 다운로드 (.xlsx)",
                data=bio,
                file_name="타임테이블_후보안.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
