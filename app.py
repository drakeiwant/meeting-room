import streamlit as st
import pandas as pd
from datetime import datetime, time, timedelta, date
import re
from io import BytesIO

import holidays
import plotly.express as px

st.set_page_config(page_title="회의실 예약률 대시보드", layout="wide")
st.title("🏢 회의실별 월간 예약률 대시보드")
st.caption("엑셀 업로드 → 월 선택 → 주말/공휴일 제외 + 시간대(전체/주요업무) 기준 예약률 + 최근 3개월 평균/추이 + 인사이트 + 엑셀 다운로드")

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
    st.subheader("인사이트 기준(최근 3개월 평균)")
    th_reduce = st.slider("축소/통합 검토(%) 미만", min_value=0, max_value=50, value=20, step=1)
    th_improve = st.slider("활용 개선 필요(%) 미만", min_value=10, max_value=70, value=40, step=1)

    st.divider()
    st.subheader("표시 옵션")
    cap_at_100 = st.checkbox("그래프는 0~100% 스케일로 보기(100% 초과는 100으로 표시)", value=True)
    top_n = st.slider("막대 그래프 상위 N개 보기(0이면 전체)", min_value=0, max_value=100, value=0, step=5)

    st.divider()
    show_debug = st.checkbox("디버그 정보(선택)", value=False)

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

def add_months(ts: pd.Timestamp, n: int) -> pd.Timestamp:
    y = ts.year
    m = ts.month + n
    while m <= 0:
        m += 12
        y -= 1
    while m > 12:
        m -= 12
        y += 1
    return pd.Timestamp(year=y, month=m, day=1)

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
    if time_mode.startswith("주요업무시간"):
        return [(time(9, 0), time(11, 0)), (time(14, 0), time(17, 0))]
    return [(time(9, 0), time(18, 0))]

def window_minutes_per_day(windows):
    dummy = datetime(2000, 1, 1)
    mins = 0.0
    for ws_t, we_t in windows:
        mins += (datetime.combine(dummy.date(), we_t) - datetime.combine(dummy.date(), ws_t)).total_seconds() / 60
    return mins

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

def classify(avg_rate: float) -> str:
    if avg_rate < th_reduce:
        return "🟥 축소/통합 검토"
    if avg_rate < th_improve:
        return "🟨 활용 개선 필요"
    return "🟩 유지 권장"

def df_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "result") -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

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

# 최근 3개월(선택월 포함)
sel_ts = pd.to_datetime(selected_month + "-01")
m2 = add_months(sel_ts, -2).strftime("%Y-%m")
m1 = add_months(sel_ts, -1).strftime("%Y-%m")
m0 = sel_ts.strftime("%Y-%m")
last3_months = [m2, m1, m0]

windows = windows_for_mode()

# 3개월치 계산을 위해 기간 설정
range_start = add_months(sel_ts, -2)
range_end = add_months(sel_ts, 1)  # 선택월 다음달 1일
period_start, _ = month_range(range_start)
_, period_end = month_range(sel_ts)  # 선택월의 다음달 1일

# 공휴일 세트(연도 범위 넉넉히)
years = sorted({period_start.year, (period_end - timedelta(days=1)).year})
holiday_set = build_holiday_set(years)

# 선택된 3개월과 겹치는 예약만 남기기
df_3 = df[(df[end_col] > period_start) & (df[start_col] < period_end)].copy()

# 일자별 분해(시간대 클리핑/주말공휴일 제외 포함)
rows = []
for _, r in df_3.iterrows():
    s = max(r[start_col].to_pydatetime(), period_start.to_pydatetime())
    e = min(r[end_col].to_pydatetime(), period_end.to_pydatetime())
    rows.extend(explode_reservation_to_daily_minutes(r["room"], s, e, holiday_set, windows))

daily = pd.DataFrame(rows)
if daily.empty:
    st.warning("최근 3개월 범위에 계산 가능한 예약이 없어요(근무일/시간대 기준).")
    st.stop()

daily["month"] = pd.to_datetime(daily["day"]).dt.to_period("M").astype(str)
daily = daily[daily["month"].isin(last3_months)].copy()

# 월별 가용시간(회의실당): 근무일수 * windows합산
def available_minutes_for_month(month_str: str) -> float:
    ms = pd.to_datetime(month_str + "-01")
    ms_start, ms_end = month_range(ms)
    days = pd.date_range(ms_start, ms_end - pd.Timedelta(days=1), freq="D")
    workdays = [ts.date() for ts in days if is_workday(ts.date(), holiday_set)]
    return len(workdays) * window_minutes_per_day(windows)

avail_map = {m: available_minutes_for_month(m) for m in last3_months}

# 월별/회의실별 예약률
monthly = daily.groupby(["month", "room"], as_index=False)["reserved_min"].sum()
monthly["avail_min"] = monthly["month"].map(avail_map)
monthly["rate"] = (monthly["reserved_min"] / monthly["avail_min"]) * 100
monthly["rate"] = monthly["rate"].clip(lower=0)

# 선택월 현재(월간) 테이블
cur = monthly[monthly["month"] == selected_month].copy()
if cur.empty:
    st.warning("선택한 월에 계산 가능한 예약이 없어요(근무일/시간대 기준).")
    st.stop()

cur["reserved_h"] = cur["reserved_min"] / 60
cur["avail_h"] = cur["avail_min"] / 60

cur = cur.sort_values("rate", ascending=False).reset_index(drop=True)

# 최근 3개월 평균 계산(회의실별)
avg3 = (
    monthly.groupby("room", as_index=False)["rate"]
    .mean()
    .rename(columns={"rate": "avg_3m_rate"})
)
avg3["판단"] = avg3["avg_3m_rate"].apply(classify)

# 현재월 테이블에 3개월 평균/판단 붙이기
cur2 = cur.merge(avg3, on="room", how="left")

# KPI
st.subheader("2) 대시보드")
workday_count_cur = int(avail_map[selected_month] / window_minutes_per_day(windows))
top_room = cur2.iloc[0]["room"]
top_rate = float(cur2.iloc[0]["rate"])

c1, c2, c3, c4 = st.columns(4)
c1.metric("대상 월", selected_month)
c2.metric("근무일 수", f"{workday_count_cur}일")
c3.metric("회의실 수", f"{cur2['room'].nunique()}개")
c4.metric("최고 예약률", f"{top_rate:.1f}%", help=f"{top_room}")

# 자동 요약 + 인사이트
avg_rate_cur = float(cur2["rate"].mean())
avg_rate_3m_overall = float(avg3["avg_3m_rate"].mean())

reduce_list = avg3[avg3["판단"].str.contains("축소")].sort_values("avg_3m_rate").head(10)
improve_list = avg3[avg3["판단"].str.contains("개선")].sort_values("avg_3m_rate").head(10)

with st.container(border=True):
    st.markdown(f"**{selected_month} 자동 요약**")
    st.markdown(f"- 이번 달 평균 예약률: **{avg_rate_cur:.1f}%**")
    st.markdown(f"- 최근 3개월 전체 평균 예약률: **{avg_rate_3m_overall:.1f}%**")
    st.markdown(f"- 이번 달 TOP: **{top_room} ({top_rate:.1f}%)**")

    if len(reduce_list) > 0:
        st.markdown("**🧠 인사이트: 축소/통합 검토 후보(최근 3개월 평균)**")
        st.markdown("- " + ", ".join([f"{r}({p:.1f}%)" for r, p in reduce_list[["room", "avg_3m_rate"]].values]))
    else:
        st.markdown("**🧠 인사이트**: 현재 기준으로는 축소/통합 검토 후보가 뚜렷하지 않아요.")

    if len(improve_list) > 0:
        st.markdown("**📌 활용 개선 필요 후보(최근 3개월 평균)**")
        st.markdown("- " + ", ".join([f"{r}({p:.1f}%)" for r, p in improve_list[["room", "avg_3m_rate"]].values]))

# -----------------------------
# (A) 이번 달 막대그래프 (회의실명 세로축)
# -----------------------------
chart_df = cur2.copy()
if top_n and top_n > 0:
    chart_df = chart_df.head(top_n).copy()

chart_df["rate_show"] = chart_df["rate"].clip(upper=100) if cap_at_100 else chart_df["rate"]

palette = (PASTEL * ((len(chart_df) // len(PASTEL)) + 1))[:len(chart_df)]
color_map = dict(zip(chart_df["room"], palette))

fig = px.bar(
    chart_df,
    x="rate_show",
    y="room",
    orientation="h",
    text=chart_df["rate"].round(1),   # 실제 값 표시(100 초과 포함)
    color="room",
    color_discrete_map=color_map,
    labels={"rate_show": "예약률(%)", "room": "회의실"},
)

fig.update_traces(textposition="outside", cliponaxis=False)
x_max = 105 if cap_at_100 else max(105, float(chart_df["rate"].max()) * 1.1)
fig.update_layout(
    showlegend=False,
    xaxis=dict(range=[0, x_max]),
    yaxis=dict(autorange="reversed"),
    margin=dict(l=10, r=30, t=10, b=10),
    height=min(900, 30 * len(chart_df) + 140),
)
st.plotly_chart(fig, use_container_width=True)

# -----------------------------
# (B) 최근 3개월 추이(라인 차트)
# -----------------------------
st.markdown("### 📈 최근 3개월 예약률 추이")

# 추이용: 없는 월은 0으로 채우기(회의실×월 그리드)
rooms_all = sorted(set(monthly["room"]))
grid = pd.MultiIndex.from_product([last3_months, rooms_all], names=["month", "room"]).to_frame(index=False)
trend = grid.merge(monthly[["month", "room", "rate"]], on=["month", "room"], how="left")
trend["rate"] = trend["rate"].fillna(0.0)

default_rooms = list(cur2.head(min(5, len(cur2)))["room"])
sel_rooms = st.multiselect("추이를 볼 회의실 선택", options=rooms_all, default=default_rooms)

trend_view = trend[trend["room"].isin(sel_rooms)].copy()
trend_view["month_dt"] = pd.to_datetime(trend_view["month"] + "-01")

fig2 = px.line(
    trend_view.sort_values(["room", "month_dt"]),
    x="month_dt",
    y="rate",
    color="room",
    markers=True,
    labels={"month_dt": "월", "rate": "예약률(%)", "room": "회의실"},
)
fig2.update_layout(
    xaxis=dict(tickformat="%Y-%m"),
    yaxis=dict(range=[0, min(110, max(10, float(trend_view["rate"].max()) * 1.2))]),
    margin=dict(l=10, r=10, t=10, b=10),
    height=420
)
st.plotly_chart(fig2, use_container_width=True)

# -----------------------------
# (C) 테이블 + 엑셀 다운로드
# -----------------------------
st.markdown("### 📋 회의실별 예약률 테이블 (이번 달 + 최근3개월 평균/판단)")

table_df = cur2.rename(columns={
    "room": "회의실",
    "reserved_h": "예약시간(시간)",
    "avail_h": "가용시간(시간)",
    "rate": "예약률(%)",
    "avg_3m_rate": "최근3개월평균예약률(%)"
})

table_df["예약시간(시간)"] = table_df["예약시간(시간)"].round(1)
table_df["가용시간(시간)"] = table_df["가용시간(시간)"].round(1)
table_df["예약률(%)"] = table_df["예약률(%)"].round(1)
table_df["최근3개월평균예약률(%)"] = table_df["최근3개월평균예약률(%)"].round(1)

table_out = table_df[["회의실", "예약시간(시간)", "가용시간(시간)", "예약률(%)", "최근3개월평균예약률(%)", "판단"]]
st.dataframe(table_out, use_container_width=True)

# 다운로드: 결과(테이블) + 추이(trend)도 같이 2시트로 저장
excel_bytes = None
try:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        table_out.to_excel(writer, index=False, sheet_name=f"예약률_{selected_month}")
        # 추이도 저장(선택된 회의실만)
        trend_save = trend_view.copy()
        trend_save["월"] = trend_save["month"]
        trend_save = trend_save[["월", "room", "rate"]].rename(columns={"room": "회의실", "rate": "예약률(%)"})
        trend_save.to_excel(writer, index=False, sheet_name="최근3개월_추이")
    excel_bytes = output.getvalue()
except Exception as e:
    st.warning(f"엑셀 생성 중 오류: {e}")

if excel_bytes:
    st.download_button(
        label="📥 결과 엑셀 다운로드",
        data=excel_bytes,
        file_name=f"회의실_예약률_{selected_month}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# -----------------------------
# Debug
# -----------------------------
if show_debug:
    st.markdown("### 🧪 디버그")
    st.write("최근3개월:", last3_months)
    st.write("가용분(회의실당):", avail_map)
    st.dataframe(monthly.head(50), use_container_width=True)
