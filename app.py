import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
from datetime import datetime, time, timedelta, date
import re

import holidays  # requirements.txt에 포함되어 있어야 함

st.set_page_config(page_title="회의실 예약률 대시보드", layout="wide")
st.title("🏢 회의실별 월간 예약률 대시보드")

uploaded = st.file_uploader("📂 예약현황 Raw 엑셀 업로드 (.xlsx)", type=["xlsx"])

# -----------------------------
# Sidebar
# -----------------------------
with st.sidebar:
    st.header("설정")

    time_mode = st.radio(
        "시간대 기준",
        options=["전체업무시간 (09:00~18:00)", "주요업무시간 (09~11, 14~17)"],
        index=0
    )

    exclude_weekends = st.checkbox("주말 제외", value=True)
    exclude_holidays = st.checkbox("공휴일 제외(대한민국)", value=True)

    st.divider()
    st.subheader("대시보드 연출")
    show_loading = st.checkbox("계산 로딩 연출", value=True)

# -----------------------------
# Helpers
# -----------------------------
PASTEL = [
    "#AEC6CF", "#FFB347", "#B39EB5", "#77DD77", "#FF6961",
    "#FDFD96", "#CFCFC4", "#F49AC2", "#CB99C9", "#BDB2FF",
    "#A0E7E5", "#FFDAC1", "#E2F0CB", "#C7CEEA", "#FFD1DC"
]

def month_range(dt: pd.Timestamp):
    first = dt.replace(day=1)
    if first.month == 12:
        nxt = first.replace(year=first.year + 1, month=1)
    else:
        nxt = first.replace(month=first.month + 1)
    return first, nxt

def build_holiday_set(years):
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

def windows_for_mode():
    # 요청: 주요업무시간 = 09~11, 14~17
    if time_mode.startswith("주요업무시간"):
        return [(time(9, 0), time(11, 0)), (time(14, 0), time(17, 0))]
    # 기본: 전체업무시간 = 09~18
    return [(time(9, 0), time(18, 0))]

def minutes_in_windows(day: date, s: datetime, e: datetime, windows):
    """
    (day, s~e) 구간을 지정된 windows(예: 09-11, 14-17)로 클리핑해서 총 분 반환
    """
    total = 0.0
    for ws_t, we_t in windows:
        ws = datetime.combine(day, ws_t)
        we = datetime.combine(day, we_t)
        cs = max(s, ws)
        ce = min(e, we)
        if ce > cs:
            total += (ce - cs).total_seconds() / 60
    return total

def explode_reservation_to_daily_minutes(room: str, s: datetime, e: datetime, holiday_set: set[date], windows):
    """
    예약이 여러 날에 걸쳐도 날짜별로 분해:
    - 근무일만
    - 지정 windows에 해당하는 시간만
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

        seg_s = max(s, day_start)
        seg_e = min(e, day_end)

        mins = minutes_in_windows(cur_day, seg_s, seg_e, windows)
        if mins > 0:
            out.append({"회의실": room, "일자": cur_day, "예약시간(분)": mins})

        cur_day += timedelta(days=1)

    return out

def clean_room_name(x: str) -> str:
    # 회의실명 괄호 제거: "회의실A(본관)" -> "회의실A"
    x = str(x)
    x = re.sub(r"\s*\(.*\)\s*$", "", x)  # 끝 괄호 부분 제거
    return x.strip()

def make_auto_commentary(usage: pd.DataFrame, selected_month: str) -> str:
    if usage.empty:
        return "표시할 데이터가 없어요."

    avg = usage["예약률(%)"].mean()
    top3 = usage.head(3)[["회의실", "예약률(%)"]].values.tolist()
    bottom3 = usage.tail(3)[["회의실", "예약률(%)"]].values.tolist()

    low = usage[usage["예약률(%)"] < 10][["회의실", "예약률(%)"]]
    low_line = ""
    if len(low) > 0:
        low_list = ", ".join([f"{r}({p:.1f}%)" for r, p in low.values[:8]])
        if len(low) > 8:
            low_list += f" 외 {len(low)-8}개"
        low_line = f"- **저활용(10% 미만)**: {low_list}"

    lines = []
    lines.append(f"**{selected_month} 자동 요약**")
    lines.append(f"- **평균 예약률**: {avg:.1f}%")
    lines.append("- **상위 3개 회의실**: " + ", ".join([f"{r}({p:.1f}%)" for r, p in top3]))
    lines.append("- **하위 3개 회의실**: " + ", ".join([f"{r}({p:.1f}%)" for r, p in bottom3]))
    if low_line:
        lines.append(low_line)

    # 운영팀용 한 줄 코멘트
    busiest_room, busiest_rate = top3[0]
    lines.append(f"- **한 줄 코멘트**: 이번 달은 **{busiest_room}** 예약률이 가장 높아요(**{busiest_rate:.1f}%**).")

    return "\n".join(lines)

# -----------------------------
# Main
# -----------------------------
if not uploaded:
    st.info("엑셀을 업로드하면 대시보드가 생성돼요.")
    st.stop()

df = pd.read_excel(uploaded)
st.subheader("1) Raw 미리보기")
st.dataframe(df.head(30), use_container_width=True)

# 네 파일 기준 컬럼명(오타 중료 지원)
room_col = "회의실"
start_col = "시작"
end_col = "중료" if "중료" in df.columns else ("종료" if "종료" in df.columns else None)

missing = [c for c in [room_col, start_col, end_col] if (c is None or c not in df.columns)]
if missing:
    st.error(f"필수 컬럼을 못 찾았어: {missing}\n(현재는 회의실/시작/중료(or 종료) 필요)")
    st.stop()

df = df.copy()
df[start_col] = pd.to_datetime(df[start_col], errors="coerce")
df[end_col] = pd.to_datetime(df[end_col], errors="coerce")
df = df[df[start_col].notna() & df[end_col].notna()].copy()
df = df[df[end_col] > df[start_col]].copy()

# 회의실명 정리(괄호 제거)
df["회의실_clean"] = df[room_col].apply(clean_room_name)

# 월 선택
df["month"] = df[start_col].dt.to_period("M").astype(str)
months = sorted(df["month"].unique())
st.subheader("2) 월 선택")
selected_month = st.selectbox("집계할 월(YYYY-MM)", months, index=len(months)-1)

ms = pd.to_datetime(selected_month + "-01")
month_start, month_end = month_range(ms)
years = sorted({month_start.year, (month_end - timedelta(days=1)).year})
holiday_set = build_holiday_set(years)

windows = windows_for_mode()

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

# 선택 월과 겹치는 예약만
month_s = month_start.to_pydatetime()
month_e = month_end.to_pydatetime()
df_m = df[(df[end_col] > month_s) & (df[start_col] < month_e)].copy()

# 날짜별 예약 분해(=windows로 클리핑 포함)
rows = []
for _, r in df_m.iterrows():
    room = r["회의실_clean"]
    s = max(r[start_col].to_pydatetime(), month_s)
    e = min(r[end_col].to_pydatetime(), month_e)
    rows.extend(explode_reservation_to_daily_minutes(room, s, e, holiday_set, windows))

daily = pd.DataFrame(rows)
if daily.empty:
    st.warning("선택한 월에 계산 가능한 예약이 없어요(근무일/시간대 기준).")
    st.stop()

# 월간 회의실별 예약시간
usage = daily.groupby("회의실", as_index=False)["예약시간(분)"].sum()
usage["예약시간(시간)"] = usage["예약시간(분)"] / 60

# 가용시간 계산: 근무일 수 * (windows 합산 시간)
days = pd.date_range(month_start, month_end - pd.Timedelta(days=1), freq="D")
workdays = [ts.date() for ts in days if is_workday(ts.date(), holiday_set)]
workday_count = len(workdays)

window_minutes_per_day = 0.0
for ws_t, we_t in windows:
    dummy = datetime(2000, 1, 1)
    window_minutes_per_day += (datetime.combine(dummy.date(), we_t) - datetime.combine(dummy.date(), ws_t)).total_seconds() / 60

available_minutes_per_room = workday_count * window_minutes_per_day

usage["가용시간(시간)"] = available_minutes_per_room / 60
usage["예약률(%)"] = (usage["예약시간(분)"] / available_minutes_per_room) * 100

# 0% 미만 방지(부동소수점/예외 케이스)
usage["예약률(%)"] = usage["예약률(%)"].clip(lower=0)

usage = usage.sort_values("예약률(%)", ascending=False).reset_index(drop=True)

# -----------------------------
# Dashboard
# -----------------------------
st.subheader("3) 대시보드")

# KPI
top_room = usage.iloc[0]["회의실"]
top_rate = usage.iloc[0]["예약률(%)"]
avg_rate = usage["예약률(%)"].mean()
total_rooms = usage["회의실"].nunique()

c1, c2, c3, c4 = st.columns(4)
c1.metric("대상 월", selected_month)
c2.metric("근무일 수", f"{workday_count}일")
c3.metric("회의실 수", f"{total_rooms}개")
c4.metric("최고 예약률", f"{top_rate:.1f}%", help=f"{top_room}")

# 자동 코멘트
with st.container(border=True):
    st.markdown(make_auto_commentary(usage, selected_month))

# 차트: 가로 막대 + 파스텔 + 회의실명 가독성
chart_df = usage.copy()
chart_df["예약률(%)"] = chart_df["예약률(%)"].round(1)

# 파스텔 색상을 회의실 수만큼 순환 배정
rooms = chart_df["회의실"].tolist()
color_map = {}
for i, r in enumerate(rooms):
    color_map[r] = PASTEL[i % len(PASTEL)]

color_scale = alt.Scale(domain=list(color_map.keys()), range=list(color_map.values()))

bar = alt.Chart(chart_df).mark_bar(cornerRadius=6).encode(
    y=alt.Y("회의실:N", sort="-x", title=None, axis=alt.Axis(labelLimit=300, labelFontSize=12)),
    x=alt.X("예약률_plot:Q", title="예약률(%)", scale=alt.Scale(domain=[0, max(5, float(chart_df["예약률(%)"].max()) * 1.1)])),
    color=alt.Color("회의실:N", scale=color_scale, legend=None),
    tooltip=[
        alt.Tooltip("회의실:N"),
        alt.Tooltip("예약률(%):Q", format=".1f"),
        alt.Tooltip("예약시간(시간):Q", format=".1f"),
        alt.Tooltip("가용시간(시간):Q", format=".1f"),
    ],
).transform_calculate(
    **{
        # altair 필드명 이슈 회피용
        "예약률_plot": "datum['예약률(%)']",
        "예약률(%)": "datum['예약률(%)']"
    }
)

text = alt.Chart(chart_df).mark_text(align="left", dx=5, fontSize=12).encode(
    y=alt.Y("회의실:N", sort="-x"),
    x=alt.X("예약률_plot:Q"),
    text=alt.Text("예약률_plot:Q", format=".1f")
)

st.altair_chart((bar + text).properties(height=min(700, 30 * len(chart_df) + 80)).interactive(), use_container_width=True)

# 테이블
st.markdown("### 📋 회의실별 예약률 테이블")
st.dataframe(
    usage[["회의실", "예약시간(시간)", "가용시간(시간)", "예약률(%)"]]
    .round({"예약시간(시간)": 1, "가용시간(시간)": 1, "예약률(%)": 1}),
    use_container_width=True
)

with st.expander("🔎 일자별 예약시간(분) 보기"):
    st.dataframe(daily.head(1000), use_container_width=True)

