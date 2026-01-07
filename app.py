import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
from datetime import datetime, time, timedelta, date

import holidays  # pip install holidays

st.set_page_config(page_title="회의실 예약률 대시보드", layout="wide")
st.title("🏢 회의실별 월간 예약률 대시보드")
st.caption("Raw 엑셀 업로드 → 월 선택 → 주말/공휴일 제외 + 09~18 클리핑 → 회의실별 예약률")

uploaded = st.file_uploader("📂 예약현황 Raw 엑셀 업로드 (.xlsx)", type=["xlsx"])

with st.sidebar:
    st.header("설정")
    work_start = st.time_input("업무 시작", value=time(9, 0))
    work_end = st.time_input("업무 종료", value=time(18, 0))

    st.divider()
    st.subheader("가용일 옵션")
    exclude_weekends = st.checkbox("주말 제외", value=True)
    exclude_holidays = st.checkbox("공휴일 제외(대한민국)", value=True)

    st.divider()
    st.subheader("대시보드 연출")
    show_loading = st.checkbox("계산 로딩 연출", value=True)

def month_range(dt: pd.Timestamp):
    first = dt.replace(day=1)
    if first.month == 12:
        nxt = first.replace(year=first.year + 1, month=1)
    else:
        nxt = first.replace(month=first.month + 1)
    return first, nxt

def minutes_between(t1: time, t2: time) -> float:
    dummy = datetime(2000, 1, 1)
    return (datetime.combine(dummy.date(), t2) - datetime.combine(dummy.date(), t1)).total_seconds() / 60

def build_holiday_set(years):
    if not years:
        return set()
    if not exclude_holidays:
        return set()
    kr = holidays.KR(years=years)
    return set(kr.keys())

def is_workday(d: date, holiday_set: set[date]) -> bool:
    if exclude_weekends and d.weekday() >= 5:
        return False
    if exclude_holidays and (d in holiday_set):
        return False
    return True

def clip_interval_to_workhours(day: date, s: datetime, e: datetime):
    """요청사항 2) 09 이전 시작은 09로, 18 이후 종료는 18로 클리핑"""
    ws = datetime.combine(day, work_start)
    we = datetime.combine(day, work_end)
    cs = max(s, ws)
    ce = min(e, we)
    if ce <= cs:
        return None
    return cs, ce

def explode_reservation_to_daily_minutes(room: str, s: datetime, e: datetime, holiday_set: set[date]):
    """
    예약이 여러 날에 걸쳐도 날짜별로 분해해서,
    - 근무일만
    - 근무시간(09~18)만
    예약 분(min) 계산
    """
    out = []
    cur_day = s.date()
    last_day = e.date()

    while cur_day <= last_day:
        if not is_workday(cur_day, holiday_set):
            cur_day += timedelta(days=1)
            continue

        day_start = datetime.combine(cur_day, time(0, 0))
        day_end = datetime.combine(cur_day, time(23, 59, 59))

        # 그 날짜에서의 실제 예약 구간
        seg_s = max(s, day_start)
        seg_e = min(e, day_end)

        clipped = clip_interval_to_workhours(cur_day, seg_s, seg_e)
        if clipped:
            cs, ce = clipped
            mins = (ce - cs).total_seconds() / 60
            if mins > 0:
                out.append({"회의실": room, "일자": cur_day, "예약시간(분)": mins})

        cur_day += timedelta(days=1)

    return out

if not uploaded:
    st.info("엑셀을 업로드하면 대시보드가 생성돼요.")
    st.stop()

df = pd.read_excel(uploaded)
st.subheader("1) Raw 미리보기")
st.dataframe(df.head(30), use_container_width=True)

# 네 파일 기준 컬럼명 매핑
# (파일에 '중료'로 들어가 있어서 그대로 지원)
room_col = "회의실"
start_col = "시작"
end_col = "중료" if "중료" in df.columns else ("종료" if "종료" in df.columns else None)

missing = [c for c in [room_col, start_col, end_col] if (c is None or c not in df.columns)]
if missing:
    st.error(f"필수 컬럼을 못 찾았어: {missing}\n(현재는 회의실/시작/중료(or 종료) 필요)")
    st.stop()

# datetime 파싱
df = df.copy()
df[start_col] = pd.to_datetime(df[start_col], errors="coerce")
df[end_col] = pd.to_datetime(df[end_col], errors="coerce")

df = df[df[start_col].notna() & df[end_col].notna()].copy()
df = df[df[end_col] > df[start_col]].copy()

# 월 선택
df["month"] = df[start_col].dt.to_period("M").astype(str)
months = sorted(df["month"].unique())
selected_month = st.selectbox("2) 집계할 월(YYYY-MM)", months, index=len(months)-1)

ms = pd.to_datetime(selected_month + "-01")
month_start, month_end = month_range(ms)
years = sorted({month_start.year, (month_end - timedelta(days=1)).year})
holiday_set = build_holiday_set(years)

# 로딩 연출
if show_loading:
    ph = st.empty()
    prog = st.progress(0)
    for i in range(0, 101, 20):
        ph.markdown(f"🧮 계산 중… **{i}%**")
        prog.progress(i)
        import time as _t
        _t.sleep(0.03)
    ph.empty()
    prog.empty()

# 선택 월에 “겹치는” 예약만 남기기 (예약이 월을 가로질러도 일부만 반영)
month_s = month_start.to_pydatetime()
month_e = month_end.to_pydatetime()
df_m = df[(df[end_col] > month_s) & (df[start_col] < month_e)].copy()

# 예약 분해(날짜별 예약시간)
rows = []
for _, r in df_m.iterrows():
    room = r[room_col]
    s = max(r[start_col].to_pydatetime(), month_s)
    e = min(r[end_col].to_pydatetime(), month_e)
    rows.extend(explode_reservation_to_daily_minutes(room, s, e, holiday_set))

daily = pd.DataFrame(rows)
if daily.empty:
    st.warning("선택한 월에 계산 가능한 예약이 없어요(근무일/근무시간 기준).")
    st.stop()

# 월간 회의실별 예약시간
usage = daily.groupby("회의실", as_index=False)["예약시간(분)"].sum()
usage["예약시간(시간)"] = usage["예약시간(분)"] / 60

# 월간 가용시간(요청사항 3)
# 가용시간 = 근무일수(주말/공휴일 제외) * 9시간(=work_start~work_end)
days = pd.date_range(month_start, month_end - pd.Timedelta(days=1), freq="D")
workdays = []
for ts in days:
    d = ts.date()
    if is_workday(d, holiday_set):
        workdays.append(d)
workday_count = len(workdays)
daily_minutes = minutes_between(work_start, work_end)
available_minutes_per_room = workday_count * daily_minutes

usage["가용시간(시간)"] = available_minutes_per_room / 60
usage["예약률(%)"] = (usage["예약시간(분)"] / available_minutes_per_room) * 100
usage = usage.sort_values("예약률(%)", ascending=False).reset_index(drop=True)

# KPI
st.subheader("3) 대시보드")
c1, c2, c3, c4 = st.columns(4)
c1.metric("대상 월", selected_month)
c2.metric("근무일 수", f"{workday_count}일")
c3.metric("회의실 수", f"{usage['회의실'].nunique()}개")
c4.metric("최고 예약률", f"{usage.iloc[0]['예약률(%)']:.1f}%", help=f"{usage.iloc[0]['회의실']}")

# 차트(인터랙티브)
chart_df = usage.copy()
chart_df["예약률(%)"] = chart_df["예약률(%)"].round(1)

bar = alt.Chart(chart_df).mark_bar().encode(
    x=alt.X("회의실:N", sort="-y", title="회의실"),
    y=alt.Y("예약률(%):Q", title="예약률(%)"),
    tooltip=[
        alt.Tooltip("회의실:N"),
        alt.Tooltip("예약률(%):Q", format=".1f"),
        alt.Tooltip("예약시간(시간):Q", format=".1f"),
        alt.Tooltip("가용시간(시간):Q", format=".1f"),
    ],
).transform_calculate(
    **{"예약률(%)": "datum['예약률(%)']"}
)

text = alt.Chart(chart_df).mark_text(dy=-8).encode(
    x=alt.X("회의실:N", sort="-y"),
    y=alt.Y("예약률(%) :Q"),
    text=alt.Text("예약률(%) :Q", format=".1f")
)

st.altair_chart((bar + text).properties(height=420).interactive(), use_container_width=True)

# 테이블
st.markdown("### 📋 회의실별 예약률 테이블")
st.dataframe(
    usage[["회의실", "예약시간(시간)", "가용시간(시간)", "예약률(%)"]]
      .round({"예약시간(시간)": 1, "가용시간(시간)": 1, "예약률(%)": 1}),
    use_container_width=True
)

with st.expander("🔎 일자별 예약시간(분) 보기"):
    st.dataframe(daily.head(500), use_container_width=True)
