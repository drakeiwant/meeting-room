import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
from datetime import datetime, time, timedelta, date
import re

import holidays  # requirements.txt에 포함

st.set_page_config(page_title="회의실 예약률 대시보드", layout="wide")
st.title("🏢 회의실별 월간 예약률 대시보드")
st.caption("엑셀 업로드 → 월 선택 → (주말/공휴일 제외) + 시간대 클리핑 → 회의실별 예약률")

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

PASTEL = [
    "#AEC6CF", "#FFB347", "#B39EB5", "#77DD77", "#FF6961",
    "#FDFD96", "#CFCFC4", "#F49AC2", "#CB99C9", "#BDB2FF",
    "#A0E7E5", "#FFDAC1", "#E2F0CB", "#C7CEEA", "#FFD1DC"
]

# -----------------------------
# Helpers
# -----------------------------
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
    # 주요업무시간 = 09~11, 14~17
    if time_mode.startswith("주요업무시간"):
        return [(time(9, 0), time(11, 0)), (time(14, 0), time(17, 0))]
    # 전체업무시간 = 09~18
    return [(time(9, 0), time(18, 0))]

def minutes_in_windows(day: date, s: datetime, e: datetime, windows):
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
            out.append({"room": room, "day": cur_day, "reserved_min": mins})

        cur_day += timedelta(days=1)

    return out

def clean_room_name(x: str) -> str:
    x = str(x)
    x = re.sub(r"\s*\(.*\)\s*$", "", x)  # 끝 괄호 제거
    return x.strip()

def make_auto_commentary(usage: pd.DataFrame, selected_month: str) -> str:
    if usage.empty:
        return "표시할 데이터가 없어요."

    avg = float(usage["rate"].mean())
    top3 = usage.head(3)[["room", "rate"]].values.tolist()
    bottom3 = usage.tail(3)[["room", "rate"]].values.tolist()

    low = usage[usage["rate"] < 10][["room", "rate"]]
    low_line = ""
    if len(low) > 0:
        low_list = ", ".join([f"{r}({p:.1f}%)" for r, p in low.values[:8]])
        if len(low) > 8:
            low_list += f" 외 {len(low)-8}개"
        low_line = f"- **저활용(10% 미만)**: {low_list}"

    busiest_room, busiest_rate = top3[0]
    lines = [
        f"**{selected_month} 자동 요약**",
        f"- **평균 예약률**: {avg:.1f}%",
        "- **상위 3개 회의실**: " + ", ".join([f"{r}({p:.1f}%)" for r, p in top3]),
        "- **하위 3개 회의실**: " + ", ".join([f"{r}({p:.1f}%)" for r, p in bottom3]),
    ]
    if low_line:
        lines.append(low_line)
    lines.append(f"- **한 줄 코멘트**: 이번 달은 **{busiest_room}** 예약률이 가장 높아요(**{busiest_rate:.1f}%**).")
    return "\n".join(lines)

def window_minutes_per_day(windows):
    dummy = datetime(2000, 1, 1)
    mins = 0.0
    for ws_t, we_t in windows:
        mins += (datetime.combine(dummy.date(), we_t) - datetime.combine(dummy.date(), ws_t)).total_seconds() / 60
    return mins

# -----------------------------
# Main
# -----------------------------
if not uploaded:
    st.info("엑셀을 업로드하면 대시보드가 생성돼요.")
    st.stop()

df = pd.read_excel(uploaded)

# 네 파일 기준 컬럼명(오타 '중료' 지원)
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

df["room"] = df[room_col].apply(clean_room_name)
df["month"] = df[start_col].dt.to_period("M").astype(str)

months = sorted(df["month"].unique())
if not months:
    st.error("월 정보를 만들 수 없어요. '시작' 컬럼이 날짜/시간 형식인지 확인해줘.")
    st.stop()

st.subheader("1) 월 선택")
selected_month = st.selectbox("집계할 월(YYYY-MM)", months, index=len(months) - 1)

ms = pd.to_datetime(selected_month + "-01")
month_start, month_end = month_range(ms)
years = sorted({month_start.year, (month_end - timedelta(days=1)).year})
holiday_set = build_holiday_set(years)
windows = windows_for_mode()

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

# 일자별 분해
rows = []
for _, r in df_m.iterrows():
    s = max(r[start_col].to_pydatetime(), month_s)
    e = min(r[end_col].to_pydatetime(), month_e)
    rows.extend(explode_reservation_to_daily_minutes(r["room"], s, e, holiday_set, windows))

daily = pd.DataFrame(rows)
if daily.empty:
    st.warning("선택한 월에 계산 가능한 예약이 없어요(근무일/시간대 기준).")
    st.stop()

# 월간 집계
usage = daily.groupby("room", as_index=False)["reserved_min"].sum()
usage["reserved_h"] = usage["reserved_min"] / 60

days = pd.date_range(month_start, month_end - pd.Timedelta(days=1), freq="D")
workdays = [ts.date() for ts in days if is_workday(ts.date(), holiday_set)]
workday_count = len(workdays)

avail_min_per_room = workday_count * window_minutes_per_day(windows)
usage["avail_h"] = avail_min_per_room / 60
usage["rate"] = (usage["reserved_min"] / avail_min_per_room) * 100
usage["rate"] = usage["rate"].clip(lower=0)

usage = usage.sort_values("rate", ascending=False).reset_index(drop=True)

# -----------------------------
# Dashboard
# -----------------------------
st.subheader("2) 대시보드")

top_room = usage.iloc[0]["room"]
top_rate = float(usage.iloc[0]["rate"])
total_rooms = int(usage["room"].nunique())

c1, c2, c3, c4 = st.columns(4)
c1.metric("대상 월", selected_month)
c2.metric("근무일 수", f"{workday_count}일")
c3.metric("회의실 수", f"{total_rooms}개")
c4.metric("최고 예약률", f"{top_rate:.1f}%", help=f"{top_room}")

with st.container(border=True):
    st.markdown(make_auto_commentary(usage, selected_month))

# ---- 차트(여기서 막대 안 나오던 문제 해결: 안전한 컬럼명 사용) ----
chart_df = usage.copy()
chart_df["rate"] = pd.to_numeric(chart_df["rate"], errors="coerce").fillna(0.0)
chart_df["reserved_h"] = pd.to_numeric(chart_df["reserved_h"], errors="coerce").fillna(0.0)
chart_df["avail_h"] = pd.to_numeric(chart_df["avail_h"], errors="coerce").fillna(0.0)

rooms = chart_df["room"].tolist()
color_map = {r: PASTEL[i % len(PASTEL)] for i, r in enumerate(rooms)}
color_scale = alt.Scale(domain=list(color_map.keys()), range=list(color_map.values()))

y_max = float(chart_df["rate"].max()) if len(chart_df) else 0.0
y_domain_max = max(5.0, y_max * 1.15)

bar = alt.Chart(chart_df).mark_bar(cornerRadius=6).encode(
    x=alt.X(
        "room:N",
        sort="-y",
        title="회의실",
        axis=alt.Axis(labelAngle=-30, labelFontSize=12, labelLimit=220)
    ),
    y=alt.Y(
        "rate:Q",
        title="예약률(%)",
        scale=alt.Scale(domain=[0, y_domain_max])
    ),
    color=alt.Color("room:N", scale=color_scale, legend=None),
    tooltip=[
        alt.Tooltip("room:N", title="회의실"),
        alt.Tooltip("rate:Q", title="예약률(%)", format=".1f"),
        alt.Tooltip("reserved_h:Q", title="예약시간(시간)", format=".1f"),
        alt.Tooltip("avail_h:Q", title="가용시간(시간)", format=".1f"),
    ],
)

text = alt.Chart(chart_df).mark_text(dy=-8, fontSize=11).encode(
    x=alt.X("room:N", sort="-y"),
    y=alt.Y("rate:Q"),
    text=alt.Text("rate:Q", format=".1f")
)

st.altair_chart((bar + text).properties(height=420).interactive(), use_container_width=True)

# 테이블
st.markdown("### 📋 회의실별 예약률 테이블")
table_df = usage.rename(columns={
    "room": "회의실",
    "reserved_h": "예약시간(시간)",
    "avail_h": "가용시간(시간)",
    "rate": "예약률(%)"
})
st.dataframe(
    table_df[["회의실", "예약시간(시간)", "가용시간(시간)", "예약률(%)"]]
    .round({"예약시간(시간)": 1, "가용시간(시간)": 1, "예약률(%)": 1}),
    use_container_width=True
)

with st.expander("🔎 일자별 예약시간(분) 보기"):
    st.dataframe(daily.head(2000), use_container_width=True)
