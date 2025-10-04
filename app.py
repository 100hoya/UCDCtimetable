# app.py - UCDC Timetable Generator v1.0 (Stable)
# ------------------------------------------------
# 기능 요약
# - 입력: 엑셀 업로드(시트: 무대/옵션) 또는 표 직접 입력
# - 조건: 최소 휴식 '무대 수'(0 허용), 최소 휴식 '시간(분)'
# - 후보안: 우선 '휴식 만족'에서 수집 → 부족하면 '완화'로 보충
# - 후보안 최대 9개로 캡(속도/안정성)
# - 시각화: 타임라인(작게), 참가자 히트맵(작게), 휴식 없는 인원 목록
# - UI: 무작위 변수(랜덤시드), 고대비/큰 글자 토글, 템플릿 다운로드, 결과 엑셀 다운로드
# - 브랜딩: logo.png 자동 표기, use_container_width 사용(경고 제거)

from typing import List, Dict, Tuple, Optional
import io
import os
import random

import pandas as pd
import streamlit as st
import altair as alt
from PIL import Image


# ========================= 페이지 & 간단 스타일 =========================
st.set_page_config(page_title="무대 타임테이블 자동 생성기", layout="wide")

# 작고 담백한 스타일 보정(CSS)
st.markdown("""
<style>
.block-container { padding-top: 1.2rem; padding-bottom: 1.2rem; }
.stButton>button { border-radius: 10px; font-weight: 600; }
.dataframe tbody, .dataframe thead { font-size: 0.93rem; }
small, .stCaption { color: #666 !important; }
</style>
""", unsafe_allow_html=True)

# 로고 + 타이틀
logo_path = "logo.png"
if os.path.exists(logo_path):
    try:
        img = Image.open(logo_path)
        col_logo, col_title = st.columns([1, 6])
        with col_logo:
            st.image(img, caption=None, use_container_width=True)
        with col_title:
            st.title("무대 타임테이블 자동 생성기")
    except Exception:
        st.title("무대 타임테이블 자동 생성기")
else:
    st.title("무대 타임테이블 자동 생성기")

st.caption("엑셀로 불러오거나 표로 입력하고, '휴식' 조건에 맞춰 자동으로 후보안을 생성합니다.")


# ========================= 유틸 & 템플릿 =========================
@st.cache_data
def make_template_bytes() -> bytes:
    """예시 템플릿 엑셀 생성 (무대/옵션 시트)"""
    buf = io.BytesIO()
    stage_df = pd.DataFrame({
        "이름": ["오프닝", "댄스A", "보컬B", "댄스B"],
        "길이(초)": [60, 180, 150, 180],
        "참가자": ["모두", "팀X, 김하늘", "팀Y, 김하늘", "팀X"],
        "고정순서": ["", "", "2", ""],  # 보컬B를 2번째로 고정 예시
    })
    option_df = pd.DataFrame({
        "옵션": ["최소휴식무대", "후보안개수", "최소휴식초"],
        "값": [2, 5, 0],
    })
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        stage_df.to_excel(w, sheet_name="무대", index=False)
        option_df.to_excel(w, sheet_name="옵션", index=False)
    buf.seek(0)
    return buf.read()


def normalize_rows_from_df(df: pd.DataFrame) -> List[Dict]:
    """DataFrame → 내부 rows 포맷으로 정규화"""
    rename_map = {}
    if "name" in df.columns: rename_map["name"] = "이름"
    if "duration" in df.columns: rename_map["duration"] = "길이(초)"
    if "performers" in df.columns: rename_map["performers"] = "참가자"
    if "fixed" in df.columns: rename_map["fixed"] = "고정순서"
    if rename_map:
        df = df.rename(columns=rename_map)

    req_cols = {"이름", "길이(초)", "참가자"}
    missing = req_cols - set(df.columns)
    if missing:
        raise ValueError(f"필수 열 누락: {missing}")

    rows: List[Dict] = []
    for _, r in df.iterrows():
        name = str(r["이름"]).strip()
        if not name:
            continue
        try:
            dur = int(float(r["길이(초)"]))
        except Exception:
            raise ValueError(f"길이(초)는 숫자여야 합니다: 무대={name}")

        perf_raw = str(r.get("참가자", "")).strip()
        performers = [p.strip() for p in perf_raw.split(",") if p.strip()] if perf_raw else []

        fx = r.get("고정순서", "")
        fixed: Optional[int] = None
        if pd.notna(fx) and str(fx).strip() != "":
            try:
                fixed = int(float(fx))
                if fixed < 1:
                    raise ValueError
            except Exception:
                raise ValueError(f"고정순서는 1 이상의 정수여야 합니다: 무대={name}")

        rows.append({"name": name, "duration": dur, "performers": performers, "fixed": fixed})

    fixed_positions = [x["fixed"] for x in rows if x["fixed"] is not None]
    if len(fixed_positions) != len(set(fixed_positions)):
        raise ValueError("고정순서 값이 중복됩니다. 서로 다른 무대가 같은 고정번호를 가질 수 없습니다.")
    return rows


def parse_excel(file) -> Tuple[List[Dict], int, int, int]:
    """엑셀에서 무대 rows와 옵션값(r, n, rest_seconds) 읽기"""
    xls = pd.ExcelFile(file)
    stage_df = pd.read_excel(xls, sheet_name="무대")
    try:
        opt_df = pd.read_excel(xls, sheet_name="옵션")
        opt = dict(zip(opt_df["옵션"].astype(str), opt_df["값"]))
        r_from_file = int(opt.get("최소휴식무대", 2))
        n_from_file = int(opt.get("후보안개수", 5))
        rest_seconds_from_file = int(opt.get("최소휴식초", 0))
    except Exception:
        r_from_file, n_from_file, rest_seconds_from_file = 2, 5, 0

    rows = normalize_rows_from_df(stage_df)
    return rows, r_from_file, n_from_file, rest_seconds_from_file


def build_name_to_row(rows: List[Dict]) -> Dict[str, Dict]:
    return {r["name"]: r for r in rows}


def sum_durations_between(slot_names: List[str], name_to_row: Dict[str, Dict], i_prev: int, i_curr: int) -> int:
    """i_prev 이후부터 i_curr까지 누적 시간(초)"""
    if i_curr <= i_prev:
        return 0
    total = 0
    for k in range(i_prev + 1, i_curr + 1):
        total += name_to_row[slot_names[k - 1]]["duration"]
    return total


# ========================= 제약/평가 & 스케줄러 =========================
def check_constraints(
    schedule: List[str],
    rows: List[Dict],
    r_rest: int,
    min_rest_seconds: int,
    enforce_rest: bool,
) -> bool:
    """True=통과. enforce_rest=False이면 휴식제약은 무시."""
    if not enforce_rest:
        return True

    name_to_row = build_name_to_row(rows)
    last_pos: Dict[str, int] = {}
    for i, s in enumerate(schedule):
        for p in name_to_row[s]["performers"]:
            if p in last_pos:
                # 무대 수 기준
                if (i - last_pos[p]) <= r_rest:
                    return False
                # 시간 기준
                if min_rest_seconds > 0:
                    gap_seconds = sum_durations_between(schedule, name_to_row, last_pos[p], i)
                    if gap_seconds < min_rest_seconds:
                        return False
            last_pos[p] = i
    return True


def score_schedule(schedule: List[str], rows: List[Dict]) -> float:
    """간단 점수(낮을수록 좋음): 총 길이 + 근접 재등장 약한 패널티"""
    name_to_row = build_name_to_row(rows)
    total = sum(name_to_row[s]["duration"] for s in schedule)
    last_pos: Dict[str, int] = {}
    penalty = 0.0
    for i, s in enumerate(schedule):
        for p in name_to_row[s]["performers"]:
            if p in last_pos:
                penalty += max(0, 3 - (i - last_pos[p])) * 0.1
            last_pos[p] = i
    return total + penalty


def place_fixed_slots(rows: List[Dict]) -> Tuple[List[Optional[str]], List[str]]:
    """고정순서 배치, 나머지 목록 반환"""
    n = len(rows)
    board: List[Optional[str]] = [None] * n
    remain: List[str] = []
    for r in rows:
        if r["fixed"] is None:
            remain.append(r["name"])
        else:
            pos = r["fixed"] - 1
            if pos < 0 or pos >= n or board[pos] is not None:
                raise ValueError(f"고정 배치 오류: 무대={r['name']}, 위치={r['fixed']}")
            board[pos] = r["name"]
    return board, remain


def fill_board_random(board: List[Optional[str]], remain: List[str], seed: int) -> List[str]:
    """빈 칸에 remain을 랜덤 채우기"""
    rnd = random.Random(seed)
    rem = remain[:]
    rnd.shuffle(rem)
    out = board[:]
    j = 0
    for i in range(len(out)):
        if out[i] is None:
            out[i] = rem[j]
            j += 1
    return out  # type: ignore


def solve_with_seed(
    rows: List[Dict],
    r_rest: int,
    seed: int,
    min_rest_seconds: int,
    enforce_rest: bool,
    max_tries: int,
) -> Tuple[bool, Optional[List[str]]]:
    """주어진 seed부터 max_tries회 시도하여 유효 스케줄 찾기"""
    board, remain = place_fixed_slots(rows)
    if len(rows) == 0:
        return False, None

    for t in range(max_tries):
        sched = fill_board_random(board, remain, seed + t)
        if check_constraints(sched, rows, r_rest, min_rest_seconds, enforce_rest=enforce_rest):
            return True, sched
    return False, None


def make_candidates_one_phase(
    rows: List[Dict],
    r_rest: int,
    num_candidates: int,
    seed0: int,
    min_rest_seconds: int,
    enforce_rest: bool,
    tries_per_candidate: int,
) -> List[List[str]]:
    """한 단계(강제 or 완화)에서 후보안 수집"""
    found: List[List[str]] = []
    seen = set()
    seed = seed0
    hard_cap = num_candidates * tries_per_candidate
    while len(found) < num_candidates and (seed - seed0) < hard_cap:
        ok, sched = solve_with_seed(
            rows, r_rest, seed, min_rest_seconds,
            enforce_rest=enforce_rest, max_tries=tries_per_candidate
        )
        seed += 1
        if not ok or sched is None:
            continue
        key = tuple(sched)
        if key in seen:
            continue
        seen.add(key)
        found.append(sched)
    found.sort(key=lambda s: score_schedule(s, rows))
    return found


def make_candidates_two_phase(
    rows: List[Dict],
    r_rest: int,
    num_candidates: int,
    seed0: int,
    min_rest_seconds: int,
) -> Tuple[List[List[str]], int]:
    """
    1차(강제)에서 최대한 수집 → 부족하면 2차(완화)로 부족분 보충.
    반환: (최종 후보 리스트, 최종 리스트 중 '강제'로 찾은 개수)
    ※ 결과는 최대 9개로 캡(속도/안정화 목적)
    """
    capped_num = min(num_candidates, 9)

    n = max(1, len(rows))
    strict_tries = min(2400, 90 * n)
    relax_tries  = min(1800, 60 * n)

    # 1차: 강제
    strict = make_candidates_one_phase(
        rows, r_rest, capped_num, seed0, min_rest_seconds,
        enforce_rest=True, tries_per_candidate=strict_tries
    )
    strict_count = len(strict)

    if strict_count >= capped_num:
        return strict[:capped_num], capped_num

    # 2차: 완화로 부족분 보충 (시드 영역 분리)
    remaining = capped_num - strict_count
    relaxed = make_candidates_one_phase(
        rows, r_rest, remaining, seed0 + 10_000,
        min_rest_seconds=min_rest_seconds, enforce_rest=False,
        tries_per_candidate=relax_tries
    )

    # 중복 없이 합치기
    seen = set(map(tuple, strict))
    for sch in relaxed:
        t = tuple(sch)
        if t not in seen:
            strict.append(sch)
            seen.add(t)
        if len(strict) == capped_num:
            break

    return strict[:capped_num], min(strict_count, capped_num)


# ========================= 시각화 & 리포트 =========================
def make_timeline_df(schedule: List[str], name_to_row: Dict[str, Dict]) -> pd.DataFrame:
    data = []
    t = 0
    for i, s in enumerate(schedule):
        dur = name_to_row[s]["duration"]
        data.append({"무대순서": i + 1, "무대": s, "시작(초)": t, "끝(초)": t + dur})
        t += dur
    return pd.DataFrame(data)


def show_timeline_chart(timeline_df: pd.DataFrame):
    chart = alt.Chart(timeline_df).mark_bar().encode(
        x=alt.X('시작(초):Q', title='진행 시간(초)'),
        x2='끝(초):Q',
        y=alt.Y('무대순서:O', title='무대 순서', sort='ascending'),
        color=alt.Color('무대:N', legend=None),
        tooltip=['무대순서', '무대',
                 alt.Tooltip('시작(초):Q', format=','), alt.Tooltip('끝(초):Q', format=',')]
    ).properties(height=180)
    st.altair_chart(chart, use_container_width=True)


def make_people_heat_df(
    slots: List[str],
    name_to_row: Dict[str, Dict],
    r_rest: int,
    min_rest_seconds: int
) -> pd.DataFrame:
    rows: List[Dict] = []
    t = 0
    starts = []
    for s in slots:
        starts.append(t)
        t += name_to_row[s]["duration"]

    last_pos: Dict[str, int] = {}
    for i, s in enumerate(slots):
        perfs = name_to_row[s]["performers"]
        for p in perfs:
            viol_slots = False
            viol_time = False
            if p in last_pos:
                if (i - last_pos[p]) <= r_rest:
                    viol_slots = True
                if min_rest_seconds > 0:
                    gap_seconds = sum_durations_between(slots, name_to_row, last_pos[p], i)
                    if gap_seconds < min_rest_seconds:
                        viol_time = True
            rows.append({
                "무대순서": i + 1,
                "참가자": str(p),
                "무대": s,
                "위반(r)": viol_slots,
                "위반(시간)": viol_time,
                "위반여부": "위반" if (viol_slots or viol_time) else "정상",
                "시작(초)": starts[i],
            })
            last_pos[p] = i
    return pd.DataFrame(rows)


def show_people_heatmap_chart(df: pd.DataFrame, high_contrast: bool = False):
    if df.empty:
        st.info("참가자 데이터가 없어 히트맵을 표시할 수 없습니다.")
        return
    n_people = df["참가자"].nunique()
    row_h = 18
    chart_h = max(180, n_people * row_h)
    base_color = '#DDEAFB' if not high_contrast else '#EEEEEE'

    heat = alt.Chart(df).mark_rect().encode(
        x=alt.X('무대순서:O', title='무대 순서'),
        y=alt.Y('참가자:N', title='참가자', sort=alt.SortField(field='참가자', order='ascending')),
        color=alt.value(base_color),
        tooltip=['참가자', '무대', '무대순서',
                 alt.Tooltip('위반(r):N', title='무대수 기준 위반'),
                 alt.Tooltip('위반(시간):N', title='시간 기준 위반')]
    ).properties(height=chart_h)

    viol_stroke = 'red' if not high_contrast else 'black'
    viol_width = 2 if not high_contrast else 3
    viol = alt.Chart(df[df["위반여부"] == "위반"]).mark_rect(
        stroke=viol_stroke, strokeWidth=viol_width, fillOpacity=0
    ).encode(
        x='무대순서:O',
        y=alt.Y('참가자:N', sort=alt.SortField(field='참가자', order='ascending')),
    )
    warn_text = alt.Chart(df[df["위반여부"] == "위반"]).mark_text(text='!', dy=4).encode(
        x='무대순서:O',
        y=alt.Y('참가자:N', sort=alt.SortField(field='참가자', order='ascending')),
    )
    st.altair_chart((heat + viol + warn_text), use_container_width=True)


def list_no_rest_people(df_heat: pd.DataFrame) -> List[str]:
    if df_heat.empty:
        return []
    return df_heat[df_heat["위반여부"] == "위반"]["참가자"].unique().tolist()


@st.cache_data
def make_result_excel(candidates: List[List[str]], rows_df_key: str) -> bytes:
    """가벼운 엑셀(후보안 시트만) 생성"""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for idx, sched in enumerate(candidates, start=1):
            tdf = pd.DataFrame({"순서": list(range(1, len(sched)+1)), "무대": sched})
            tdf.to_excel(w, sheet_name=f"후보안_{idx}", index=False)
    buf.seek(0)
    return buf.read()


# ========================= 입력 UI =========================
mode = st.radio("입력 방식 선택", ["엑셀 업로드", "직접 입력(표)"], horizontal=True)
rows: List[Dict] = []
can_generate = False

if mode == "엑셀 업로드":
    uploaded = st.file_uploader("엑셀(.xlsx) 파일 업로드", type=["xlsx"])
    st.download_button(
        "📥 템플릿 다운로드",
        data=make_template_bytes(),
        file_name="타임테이블_템플릿_예시.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    if uploaded is not None:
        try:
            rows, r_file, n_file, rest_file = parse_excel(uploaded)
            st.info(
                f"엑셀 옵션 감지 → 최소휴식무대:{r_file}, 후보안개수:{n_file}, 쉬는시간(초):{rest_file}\n"
                "※ 실제 적용은 사이드바 설정이 우선입니다."
            )
        except Exception as e:
            st.error(f"엑셀 파싱 오류: {e}")
            st.stop()
        st.success(f"무대 {len(rows)}개 읽음")
        st.dataframe(
            pd.DataFrame(rows)[["name","duration","performers","fixed"]]
              .rename(columns={"name":"무대","duration":"길이(초)","performers":"참가자","fixed":"고정순서"}),
            use_container_width=True
        )
        can_generate = True
else:
    st.info("표의 셀을 수정해 무대를 입력하세요. 행 추가/삭제 가능.")
    init_df = pd.DataFrame({
        "이름": ["오프닝", "댄스A", "보컬B"],
        "길이(초)": [60, 180, 150],
        "참가자": ["모두", "팀X, 김하늘", "팀Y, 김하늘"],
        "고정순서": ["", "", ""],
    })
    edited = st.data_editor(init_df, num_rows="dynamic", use_container_width=True, key="manual_editor")
    try:
        rows = normalize_rows_from_df(edited)
        st.success(f"무대 {len(rows)}개 입력됨")
        can_generate = True
    except Exception as e:
        st.error(f"입력 오류: {e}")
        can_generate = False


# ========================= 사이드바(조건/접근성) =========================
with st.sidebar:
    st.header("조건 설정")

    seed0 = st.number_input("무작위 변수", value=12345, step=1)
    st.caption("💡 값을 바꾸면 다른 후보안이 생성됩니다. 같은 값은 같은 결과가 재현됩니다.")

    high_contrast = st.toggle("고대비 모드", value=False, help="색 대비를 크게 해서 읽기 쉽게 보여줍니다.")
    st.session_state["high_contrast"] = high_contrast
    large_text = st.toggle("큰 글자 모드", value=False, help="앱 전체 글자 크기를 키웁니다.")
    st.divider()

    if large_text:
        st.markdown("""
            <style>
            :root { --app-font-scale: 1.15; }
            html, body, [class*="block-container"] { font-size: calc(1rem * var(--app-font-scale)); }
            h1 { font-size: calc(2rem * var(--app-font-scale)); }
            h2 { font-size: calc(1.6rem * var(--app-font-scale)); }
            h3 { font-size: calc(1.3rem * var(--app-font-scale)); }
            .stDataFrame, .stMetric, .stButton, .stTextInput, .stNumberInput { font-size: calc(1rem * var(--app-font-scale)); }
            </style>
        """, unsafe_allow_html=True)

    # 휴식 조건 (r=0 허용)
    r_rest = st.number_input("최소 휴식 무대 수", min_value=0, value=2, step=1)
    min_rest_minutes = st.number_input("최소 휴식 시간(분)", min_value=0, value=0, step=1)
    min_rest_seconds = int(min_rest_minutes) * 60

    # 후보안 개수(실제 생성은 내부에서 최대 9개로 캡)
    num_candidates = st.number_input("후보안 개수", min_value=1, max_value=20, value=5, step=1)

    st.caption("※ 생성 우선순위: 휴식 조건 '만족' 후보안 → 부족하면 '완화' 후보안으로 보충 (최대 9개)")


# ========================= 후보안 생성 & 표시 =========================
st.divider()
col_btn, col_dl = st.columns([1, 1])
with col_btn:
    gen = st.button("후보안 생성하기", type="primary", disabled=not can_generate)

if gen and not can_generate:
    st.warning("먼저 무대 데이터를 입력하세요.")

candidates: List[List[str]] = []
strict_count: int = 0

if gen and can_generate:
    try:
        candidates, strict_count = make_candidates_two_phase(
            rows, r_rest,
            num_candidates=num_candidates,  # 내부에서 최대 9개로 캡
            seed0=seed0,
            min_rest_seconds=min_rest_seconds
        )
        if not candidates:
            st.error("조건이 과도하여 후보안을 찾지 못했습니다. 조건을 완화해 보세요.")
    except Exception as e:
        st.error(f"후보안 생성 중 오류: {e}")

# --- 결과 표시 ---
if candidates:
    actual = len(candidates)
    if strict_count == actual:
        label = "휴식 조건 ‘만족’ (전부)"
    elif strict_count == 0:
        label = "휴식 조건 ‘완화’ (전부)"
    else:
        label = f"혼합(만족 {strict_count}개 + 완화 {actual - strict_count}개)"

    st.success(f"후보안 {actual}개 생성됨 — {label}")

    # 요청 수보다 적게 나온 경우 안내 (예: 내부 캡 9개)
    if actual < min(num_candidates, 9):
        st.caption(f"요청 {num_candidates}개 중 {actual}개만 생성되었습니다. "
                   "조합이 어려워 자동 탐색 예산 내에서 더 찾지 못했습니다.")

    # 결과 엑셀 다운로드 (경량)
    rows_df_key = pd.DataFrame(rows).to_json(orient="split") if rows else "empty"
    with col_dl:
        st.download_button(
            "결과 엑셀 다운로드",
            data=make_result_excel(candidates, rows_df_key),
            file_name="타임테이블_후보안.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    name_to_row = build_name_to_row(rows)
    tabs = st.tabs([f"후보안 {i+1}" for i in range(len(candidates))])
    for i, (tab, sched) in enumerate(zip(tabs, candidates)):
        with tab:
            st.markdown("#### 순서")
            order_df = pd.DataFrame({"순서": list(range(1, len(sched)+1)), "무대": sched})
            st.dataframe(order_df, use_container_width=True)

            st.markdown("#### 타임라인 (작게)")
            tdf = make_timeline_df(sched, name_to_row)
            show_timeline_chart(tdf)

            st.markdown("#### 참가자 히트맵 (작게)")
            heat_df = make_people_heat_df(sched, name_to_row, r_rest, min_rest_seconds)
            show_people_heatmap_chart(heat_df, high_contrast=st.session_state.get("high_contrast", False))

            st.markdown("#### 휴식 없는 인원")
            bad = list_no_rest_people(heat_df)
            st.write(", ".join(bad) if bad else "없음 ✅")

st.caption("ⓒ TimetableApp — '무작위 변수' 값이 같으면 결과가 재현됩니다.")
