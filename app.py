import streamlit as st
import pandas as pd
from datetime import datetime, time, timedelta, date
import re
from io import BytesIO

import holidays
import plotly.express as px

st.set_page_config(page_title="회의실 예약률 대시보드", layout="wide")
st.title("🏢 회의실 예약률 대시보드")
st.caption("엑셀 업로드 → 월/전체 선택 → (주말·공휴일 제외) + 시간대 기준 → 예약률/추이/인사이트/엑셀 다운로드")

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
    st.subheader("인사이트 기준(평균 예약률)")
    th_reduce = st.slider("축소/통합 검토(%) 미만", 0, 50, 20, 1)
    th_improve = st.slider("활용 개선 필요(%) 미만", 10, 70, 40, 1)
    th_busy = st.slider("과밀(%) 이상", 50, 100, 80, 1)

    st.divider()
    st.subheader("표시 옵션")
    cap_at_100 = st.checkbox("그래프는 0~100% 스케일(100% 초과는 100으로 표시)", value=True)
    top_n = st.slider("막대 그래프 상위 N개(0이면 전체)", 0, 300, 0, 10)

    st.divider()
    show_debug = st.checkbox("디버그(선택)", value=False)

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
    x = re.sub(r"\s*\(.*\)\s*$", "", x)
    return x.strip()

def classify(avg_rate: float) -> str:
    if avg_rate < th_reduce:
        return "🟥 축소/통합 검토"
    if avg_rate < th_improve:
        return "🟨 활용 개선 필요"
    return "🟩 유지 권장"

def df_to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df_ in sheets.items():
            df_.to_excel(writer, index=False, sheet_name=name[:31])
    return output.getvalue()

def available_minutes_for_month(month_str: str, holiday_set: set[date], windows) -> float:
    ms = pd.to_datetime(month_str + "-01")
    ms_start, ms_end = month_range(ms)
    days = pd.date_range(ms_start, ms_end - pd.Timedelta(days=1), freq="D")
    workdays = [ts.date() for ts in days if is_workday(ts.date(), holiday_set)]
    return len(workdays) * window_minutes_per_day(windows)

# -----------------------------
# Main
# -----------------------------
if not uploaded:
    st.info("엑셀을 업로드하면 대시보드가 생성돼요.")
    st.stop()

df = pd.read_excel(uploaded)

room_col = "회의실"
start_col = "시작"
end_col = "중료" if "중료" in df.columns else ("종료" if "종료" in df.columns else None)
missing = [c for c in [room_col, start_col, end_col] if (c is None or c not in df.columns)]
if missing:
    st.error(f"필수 컬럼을 못 찾았어: {missing}")
    st.stop()

df = df.copy()
df[start_col] = pd.to_datetime(df[start_col], errors="coerce")
df[end_col] = pd.to_datetime(df[end_col], errors="coerce")
df = df[df[start_col].notna() & df[end_col].notna()].copy()
df = df[df[end_col] > df[start_col]].copy()

df["room"] = df[room_col].apply(clean_room_name)
df["month"] = df[start_col].dt.to_period("M").astype(str)

months_raw = sorted(df["month"].unique())
if not months_raw:
    st.error("월 정보를 만들 수 없어요. '시작' 컬럼 확인해줘.")
    st.stop()

month_options = ["전체(업로드 기간)"] + months_raw
st.subheader("1) 집계 범위 선택")
selected = st.selectbox("월 또는 전체", month_options, index=0)

windows = windows_for_mode()

# holiday years
min_day = df[start_col].min().date()
max_day = df[end_col].max().date()
years = tuple(range(min_day.year, max_day.year + 1))
holiday_set = build_holiday_set(list(years))

# explode → daily → monthly
rows = []
for _, r in df.iterrows():
    s = r[start_col].to_pydatetime()
    e = r[end_col].to_pydatetime()
    rows.extend(explode_reservation_to_daily_minutes(r["room"], s, e, holiday_set, windows))

daily = pd.DataFrame(rows)
if daily.empty:
    st.warning("계산 가능한 예약이 없어요(근무일/시간대 기준).")
    st.stop()

daily["month"] = pd.to_datetime(daily["day"]).dt.to_period("M").astype(str)

all_months = sorted(daily["month"].unique())
rooms_all = sorted(daily["room"].unique())

avail_map = {m: available_minutes_for_month(m, holiday_set, windows) for m in all_months}

monthly = daily.groupby(["month", "room"], as_index=False)["reserved_min"].sum()
monthly["avail_min"] = monthly["month"].map(avail_map)
monthly["rate"] = (monthly["reserved_min"] / monthly["avail_min"]) * 100
monthly["rate"] = monthly["rate"].clip(lower=0)

# -------- 전체 평균을 위해: 0인 달 제외 평균 --------
# 즉: monthly에 존재하는 month-room만 대상으로 평균 (rate>0만)
monthly_nonzero = monthly[monthly["rate"] > 0].copy()

if selected == "전체(업로드 기간)":
    scope_title = f"전체(업로드 기간) · {all_months[0]} ~ {all_months[-1]}"

    # ✅ 핵심: 0인 달 제외 평균
    agg = (monthly_nonzero.groupby("room", as_index=False)["rate"].mean()
           .rename(columns={"rate": "avg_rate"}))

    # 어떤 회의실은 전 기간 0이라 monthly_nonzero에 아예 없을 수 있음 → 표시용으로 0으로 붙임
    base_rooms = pd.DataFrame({"room": rooms_all})
    agg = base_rooms.merge(agg, on="room", how="left")
    agg["avg_rate"] = agg["avg_rate"].fillna(0.0)

    agg["rate"] = agg["avg_rate"]
    agg["판단"] = agg["avg_rate"].apply(classify)

    cur_table = agg.sort_values("rate", ascending=False).reset_index(drop=True)

    trend_months = all_months

    # 전월 대비: 마지막월 vs 전월 (없는 room은 0)
    if len(all_months) >= 2:
        last_m, prev_m = all_months[-1], all_months[-2]
        m_last = monthly[monthly["month"] == last_m][["room", "rate"]].rename(columns={"rate": "cur_rate"})
        m_prev = monthly[monthly["month"] == prev_m][["room", "rate"]].rename(columns={"rate": "prev_rate"})
        mom = pd.DataFrame({"room": rooms_all}).merge(m_last, on="room", how="left").merge(m_prev, on="room", how="left")
        mom["cur_rate"] = mom["cur_rate"].fillna(0.0)
        mom["prev_rate"] = mom["prev_rate"].fillna(0.0)
        mom["delta_pp"] = mom["cur_rate"] - mom["prev_rate"]
        mom_month_label = f"{prev_m} → {last_m}"
    else:
        mom = pd.DataFrame(columns=["room", "cur_rate", "prev_rate", "delta_pp"])
        mom_month_label = ""
else:
    scope_title = selected

    # 선택월 테이블(없는 room은 0)
    cur = pd.DataFrame({"room": rooms_all}).merge(
        monthly[monthly["month"] == selected][["room", "rate"]],
        on="room", how="left"
    )
    cur["rate"] = cur["rate"].fillna(0.0)

    # 최근 3개월 평균도 "0 제외"로 (원하면 여긴 0 포함도 가능)
    sel_ts = pd.to_datetime(selected + "-01")
    last3 = [
        (sel_ts - pd.DateOffset(months=2)).strftime("%Y-%m"),
        (sel_ts - pd.DateOffset(months=1)).strftime("%Y-%m"),
        sel_ts.strftime("%Y-%m")
    ]
    last3_exist = [m for m in last3 if m in all_months]

    avg3 = (monthly[(monthly["month"].isin(last3_exist)) & (monthly["rate"] > 0)]
            .groupby("room", as_index=False)["rate"]
            .mean()
            .rename(columns={"rate": "avg_rate"}))
    cur2 = cur.merge(avg3, on="room", how="left")
    cur2["avg_rate"] = cur2["avg_rate"].fillna(0.0)

    cur2["판단"] = cur2["avg_rate"].apply(classify)
    cur_table = cur2.sort_values("rate", ascending=False).reset_index(drop=True)

    # 전월 대비(없는 room 0)
    prev_m = (sel_ts - pd.DateOffset(months=1)).strftime("%Y-%m")
    prev_df = monthly[monthly["month"] == prev_m][["room", "rate"]].rename(columns={"rate": "prev_rate"}) if prev_m in all_months else None
    cur_df = monthly[monthly["month"] == selected][["room", "rate"]].rename(columns={"rate": "cur_rate"})

    mom = pd.DataFrame({"room": rooms_all}).merge(cur_df, on="room", how="left")
    mom["cur_rate"] = mom["cur_rate"].fillna(0.0)
    if prev_df is None:
        mom["prev_rate"] = 0.0
        mom_month_label = f"(전월 데이터 없음) → {selected}"
    else:
        mom = mom.merge(prev_df, on="room", how="left")
        mom["prev_rate"] = mom["prev_rate"].fillna(0.0)
        mom_month_label = f"{prev_m} → {selected}"
    mom["delta_pp"] = mom["cur_rate"] - mom["prev_rate"]

    trend_months = last3_exist if len(last3_exist) >= 2 else all_months

# -----------------------------
# Dashboard
# -----------------------------
st.subheader("2) 요약")

avg_rate_all = float(cur_table["rate"].mean()) if len(cur_table) else 0.0
top_room = cur_table.iloc[0]["room"] if len(cur_table) else "-"
top_rate = float(cur_table.iloc[0]["rate"]) if len(cur_table) else 0.0

c1, c2, c3, c4 = st.columns(4)
c1.metric("집계 범위", scope_title)
c2.metric("평균 예약률", f"{avg_rate_all:.1f}%")
c3.metric("회의실 수", f"{cur_table['room'].nunique()}개")
c4.metric("TOP 회의실", f"{top_room} ({top_rate:.1f}%)")

reduce_list = cur_table[cur_table["판단"].str.contains("축소")].sort_values("rate").head(10)
busy_list = cur_table[cur_table["rate"] >= th_busy].sort_values("rate", ascending=False).head(10)

up5 = mom.sort_values("delta_pp", ascending=False).head(5) if len(mom) else pd.DataFrame()
down5 = mom.sort_values("delta_pp", ascending=True).head(5) if len(mom) else pd.DataFrame()

with st.container(border=True):
    st.markdown("**STEP2 — 월별 비교(전월 대비)**")
    if len(mom):
        st.markdown(f"- 기준: **{mom_month_label}**")
        st.markdown("- 급등 TOP5: " + ", ".join([f"{r}({d:+.1f}p)" for r, d in up5[["room", "delta_pp"]].values]))
        st.markdown("- 급락 TOP5: " + ", ".join([f"{r}({d:+.1f}p)" for r, d in down5[["room", "delta_pp"]].values]))
    else:
        st.markdown("- 전월 비교를 계산할 수 있는 월 데이터가 부족해요.")

    st.markdown("---")
    st.markdown("**STEP3 — 인사이트(액션 후보)**")
    if len(reduce_list) > 0:
        st.markdown("🟥 축소/통합 검토: " + ", ".join([f"{r}({p:.1f}%)" for r, p in reduce_list[["room", "rate"]].values]))
    else:
        st.markdown("🟥 축소/통합 검토 후보: 뚜렷하지 않음")

    if len(busy_list) > 0:
        st.markdown("🔥 과밀 후보: " + ", ".join([f"{r}({p:.1f}%)" for r, p in busy_list[["room", "rate"]].values]))

# -----------------------------
# Bar chart
# -----------------------------
st.markdown("### 📊 회의실별 예약률")

chart_df = cur_table.copy()
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
    text=chart_df["rate"].round(1),
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
    height=min(950, 30 * len(chart_df) + 160),
)
st.plotly_chart(fig, use_container_width=True)

# -----------------------------
# Trend chart (추이는 0도 표시하는 게 직관적이라 그대로)
# -----------------------------
st.markdown("### 📈 예약률 추이")

trend = pd.DataFrame({"month": all_months}).merge(pd.DataFrame({"room": rooms_all}), how="cross")
trend = trend.merge(monthly[["month", "room", "rate"]], on=["month", "room"], how="left")
trend["rate"] = trend["rate"].fillna(0.0)
trend = trend[trend["month"].isin(trend_months)].copy()
trend["month_dt"] = pd.to_datetime(trend["month"] + "-01")

default_rooms = list(cur_table.head(min(5, len(cur_table)))["room"])
sel_rooms = st.multiselect("추이를 볼 회의실 선택", options=rooms_all, default=default_rooms)

trend_view = trend[trend["room"].isin(sel_rooms)].copy()
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
    yaxis=dict(range=[0, min(110, max(10, float(trend_view["rate"].max()) * 1.2) if len(trend_view) else 10))]),
    margin=dict(l=10, r=10, t=10, b=10),
    height=420
)
st.plotly_chart(fig2, use_container_width=True)

# -----------------------------
# Table + Excel
# -----------------------------
st.markdown("### 📋 결과 테이블")

out = cur_table.rename(columns={
    "room": "회의실",
    "rate": "예약률(%)",
    "avg_rate": "평균예약률(%)"
}).copy()

for col in ["예약률(%)", "평균예약률(%)"]:
    if col in out.columns:
        out[col] = out[col].astype(float).round(1)

cols = ["회의실", "예약률(%)"]
if selected != "전체(업로드 기간)" and "평균예약률(%)" in out.columns:
    cols.append("평균예약률(%)")
cols.append("판단")

out_view = out[cols].sort_values("예약률(%)", ascending=False)
st.dataframe(out_view, use_container_width=True)

trend_sheet = trend.rename(columns={"month": "월", "room": "회의실", "rate": "예약률(%)"})[["월", "회의실", "예약률(%)"]].copy()
trend_sheet["예약률(%)"] = trend_sheet["예약률(%)"].round(1)

mom_sheet = mom.copy()
if len(mom_sheet):
    mom_sheet = mom_sheet.rename(columns={"room": "회의실", "cur_rate": "이번달(%)", "prev_rate": "전월(%)", "delta_pp": "전월대비(Δ%p)"})
    for c in ["이번달(%)", "전월(%)", "전월대비(Δ%p)"]:
        mom_sheet[c] = mom_sheet[c].astype(float).round(1)
    mom_sheet = mom_sheet.sort_values("전월대비(Δ%p)", ascending=False)

insight_sheet = pd.DataFrame([
    {"구분": "축소/통합 검토", "기준": f"평균 < {th_reduce}%", "대상": ", ".join(reduce_list["room"].tolist()) if len(reduce_list) else "-"},
    {"구분": "활용 개선 필요", "기준": f"{th_reduce}% ~ {th_improve}%", "대상": ", ".join(out_view[out_view["판단"].str.contains("개선")]["회의실"].head(30).tolist()) if len(out_view) else "-"},
    {"구분": "과밀", "기준": f">= {th_busy}%", "대상": ", ".join(busy_list["room"].tolist()) if len(busy_list) else "-"},
])

sheets = {
    "집계결과": out_view,
    "월별_추이(전체)": trend_sheet,
    "인사이트": insight_sheet
}
if len(mom_sheet):
    sheets["전월대비"] = mom_sheet

excel_bytes = df_to_excel_bytes(sheets)
fname = f"회의실_예약률_{'전체' if selected=='전체(업로드 기간)' else selected}.xlsx"
st.download_button(
    "📥 결과 엑셀 다운로드",
    data=excel_bytes,
    file_name=fname,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

if show_debug:
    st.markdown("### 🧪 디버그")
    st.write("all_months:", all_months)
    st.dataframe(monthly.head(50), use_container_width=True)
    st.dataframe(monthly_nonzero.head(50), use_container_width=True)
