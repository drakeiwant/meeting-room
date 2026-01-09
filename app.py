import streamlit as st
import pandas as pd
from datetime import datetime, time, timedelta, date
import re
from io import BytesIO

import holidays
import plotly.express as px

st.set_page_config(page_title="회의실 예약률 대시보드", layout="wide")
st.title("🏢 회의실별 예약률 대시보드")
st.caption("엑셀 업로드 → 월/전체 선택 → (주말/공휴일 제외) + 시간대 기준 → 예약률/인사이트/추이 + 엑셀 다운로드")

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
    th_reduce = st.slider("축소/통합 검토(%) 미만", min_value=0, max_value=50, value=20, step=1)
    th_improve = st.slider("활용 개선 필요(%) 미만", min_value=10, max_value=70, value=40, step=1)
    th_busy = st.slider("과밀(%) 이상", min_value=50, max_value=100, value=80, step=1)

    st.divider()
    st.subheader("표시 옵션")
    cap_at_100 = st.checkbox("그래프는 0~100% 스케일로 보기(100% 초과는 100으로 표시)", value=True)
    top_n = st.slider("막대 그래프 상위 N개 보기(0이면 전체)", min_value=0, max_value=200, value=0, step=10)
    trend_mode = st.radio(
        "추이 범위",
        options=["최근 3개월", "전체 기간"],
        index=0
    )

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

def available_minutes_for_month(month_str: str, holiday_set: set[date], windows) -> float:
    ms = pd.to_datetime(month_str + "-01")
    ms_start, ms_end = month_range(ms)
    days = pd.date_range(ms_start, ms_end - pd.Timedelta(days=1), freq="D")
    workdays = [ts.date() for ts in days if is_workday(ts.date(), holiday_set)]
    return len(workdays) * window_minutes_per_day(windows)

def df_to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df_ in sheets.items():
            safe = name[:31]
            df_.to_excel(writer, index=False, sheet_name=safe)
    return output.getvalue()

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

month_options = ["전체(기간 평균)"] + months

st.subheader("1) 집계 기간 선택")
selected = st.selectbox("집계할 월(또는 전체)", month_options, index=0)

windows = windows_for_mode()

data_start = df[start_col].min().normalize()
data_end = df[end_col].max().normalize() + pd.Timedelta(days=1)

years = list(range(int(data_start.year), int((data_end - pd.Timedelta(days=1)).year) + 1))
holiday_set = build_holiday_set(years)

is_total = selected.startswith("전체")
if is_total:
    period_start = data_start
    period_end = data_end
    period_label = f"{data_start.strftime('%Y-%m-%d')} ~ {(data_end - pd.Timedelta(days=1)).strftime('%Y-%m-%d')}"
else:
    sel_ts = pd.to_datetime(selected + "-01")
    period_start, period_end = month_range(sel_ts)
    period_label = selected

df_p = df[(df[end_col] > period_start) & (df[start_col] < period_end)].copy()

rows = []
ps = period_start.to_pydatetime()
pe = period_end.to_pydatetime()
for _, r in df_p.iterrows():
    s = max(r[start_col].to_pydatetime(), ps)
    e = min(r[end_col].to_pydatetime(), pe)
    rows.extend(explode_reservation_to_daily_minutes(r["room"], s, e, holiday_set, windows))

daily = pd.DataFrame(rows)
if daily.empty:
    st.warning("선택한 기간에 계산 가능한 예약이 없어요(근무일/시간대 기준).")
    st.stop()

daily["month"] = pd.to_datetime(daily["day"]).dt.to_period("M").astype(str)

# -----------------------------
# 월별 가용시간 계산(데이터에 존재하는 월들)
# -----------------------------
all_months_in_daily = sorted(daily["month"].unique())
avail_map = {m: available_minutes_for_month(m, holiday_set, windows) for m in all_months_in_daily}

# 월별/회의실별 예약률(추이/비교용)
monthly = daily.groupby(["month", "room"], as_index=False)["reserved_min"].sum()
monthly["avail_min"] = monthly["month"].map(avail_map)
monthly["rate"] = (monthly["reserved_min"] / monthly["avail_min"]) * 100
monthly["rate"] = monthly["rate"].clip(lower=0)

# -----------------------------
# 선택 기간에 대한 집계
# -----------------------------
if is_total:
    # ====== (1) 전체(기간) 가중평균: 기간 전체 reserved / 기간 전체 avail ======
    days = pd.date_range(period_start, period_end - pd.Timedelta(days=1), freq="D")
    workdays = [ts.date() for ts in days if is_workday(ts.date(), holiday_set)]
    avail_min_period_per_room = len(workdays) * window_minutes_per_day(windows)

    cur = daily.groupby("room", as_index=False)["reserved_min"].sum()
    cur["avail_min"] = avail_min_period_per_room
    cur["weighted_rate"] = (cur["reserved_min"] / cur["avail_min"]) * 100
    cur["weighted_rate"] = cur["weighted_rate"].clip(lower=0)

    cur["reserved_h"] = cur["reserved_min"] / 60
    cur["avail_h"] = cur["avail_min"] / 60

    # ====== (2) 월 평균 예약률의 평균(단순평균): 월별 rate를 평균(월을 동일가중치) ======
    # 전체 기간에 포함되는 "월 리스트"를 완전하게 만든다(예약이 없던 달도 포함 -> rate 0)
    months_in_period = pd.period_range(period_start, period_end - pd.Timedelta(days=1), freq="M").astype(str).tolist()

    # 월별 가용시간(기간에 포함되는 월 모두)
    avail_map_period = {m: available_minutes_for_month(m, holiday_set, windows) for m in months_in_period}

    # 기간 내 월별/회의실별 reserved_min (없는 조합은 0)
    rooms_all = sorted(cur["room"].unique())
    grid = pd.MultiIndex.from_product([months_in_period, rooms_all], names=["month", "room"]).to_frame(index=False)

    monthly_in_period = monthly[monthly["month"].isin(months_in_period)].copy()
    monthly_full = grid.merge(monthly_in_period[["month", "room", "reserved_min"]], on=["month", "room"], how="left")
    monthly_full["reserved_min"] = monthly_full["reserved_min"].fillna(0.0)
    monthly_full["avail_min"] = monthly_full["month"].map(avail_map_period)
    monthly_full["month_rate"] = (monthly_full["reserved_min"] / monthly_full["avail_min"]) * 100
    monthly_full["month_rate"] = monthly_full["month_rate"].clip(lower=0)

    # 회의실별 단순평균(월 평균의 평균)
    simple_avg = monthly_full.groupby("room", as_index=False)["month_rate"].mean().rename(columns={"month_rate": "simple_avg_rate"})

    # 전체(전체 회의실 합) 월별 예약률 -> 그 월별 값의 평균(월 단순평균 KPI용)
    # (각 월을 동일 가중치로 보는 "월 평균 예약률의 평균")
    overall_month = monthly_full.groupby("month", as_index=False).agg(
        total_reserved=("reserved_min", "sum")
    )
    overall_month["total_avail"] = overall_month["month"].map(avail_map_period) * len(rooms_all)
    overall_month["overall_month_rate"] = (overall_month["total_reserved"] / overall_month["total_avail"]) * 100
    overall_month["overall_month_rate"] = overall_month["overall_month_rate"].clip(lower=0)
    overall_simple_kpi = float(overall_month["overall_month_rate"].mean()) if len(overall_month) else 0.0

    # 가중 KPI(전체 기간): (회의실별 가중평균을 단순평균 내도 동일 — 모든 회의실 avail 동일)
    overall_weighted_kpi = float(cur["weighted_rate"].mean()) if len(cur) else 0.0

    # merge
    cur2 = cur.merge(simple_avg, on="room", how="left")
    cur2["simple_avg_rate"] = cur2["simple_avg_rate"].fillna(0.0)

    # 인사이트/변동성(월별 들쭉날쭉)
    vol = monthly_full.pivot_table(index="room", columns="month", values="month_rate", aggfunc="mean").std(axis=1, skipna=True).fillna(0.0)
    vol = vol.reset_index().rename(columns={0: "volatility"})
    cur2 = cur2.merge(vol, on="room", how="left")
    cur2["volatility"] = pd.to_numeric(cur2["volatility"], errors="coerce").fillna(0.0)

    cur2["판단(가중)"] = cur2["weighted_rate"].apply(classify)
    cur2["판단(단순월평균)"] = cur2["simple_avg_rate"].apply(classify)

    cur2 = cur2.sort_values("weighted_rate", ascending=False).reset_index(drop=True)

    # KPI
    st.subheader("2) 대시보드")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("대상", "전체(기간 평균)")
    c2.metric("기간", period_label)
    c3.metric("회의실 수", f"{cur2['room'].nunique()}개")
    c4.metric("전체 예약률(가중평균)", f"{overall_weighted_kpi:.1f}%")
    c5.metric("월 평균 예약률의 평균(단순)", f"{overall_simple_kpi:.1f}%")

    # 인사이트
    reduce_list = cur2[cur2["판단(가중)"].str.contains("축소")].sort_values("weighted_rate").head(10)
    busy_list = cur2[cur2["weighted_rate"] >= th_busy].sort_values("weighted_rate", ascending=False).head(10)
    volatile = cur2.sort_values("volatility", ascending=False).head(10)

    with st.container(border=True):
        st.markdown("**인사이트(전체 기간 기준)**")
        st.markdown("- 예약률은 두 방식으로 같이 제공합니다: **가중평균(정확)** / **월 단순평균(직관)**")
        if len(reduce_list) > 0:
            st.markdown("🟥 **축소/통합 검토 후보(가중)**: " + ", ".join([f"{r}({p:.1f}%)" for r, p in reduce_list[["room","weighted_rate"]].values]))
        else:
            st.markdown("🟥 축소/통합 검토 후보: 현재 기준으로 뚜렷하지 않음")
        if len(busy_list) > 0:
            st.markdown("🔥 **과밀 후보(가중)**: " + ", ".join([f"{r}({p:.1f}%)" for r, p in busy_list[["room","weighted_rate"]].values]))
        st.markdown("📈 **변동성 TOP(월별 들쭉날쭉)**: " + ", ".join([f"{r}(σ={s:.1f})" for r, s in volatile[["room","volatility"]].values]))

    # 막대그래프(가중평균 기준)
    chart_df = cur2.copy()
    if top_n and top_n > 0:
        chart_df = chart_df.head(top_n).copy()

    chart_df["rate_show"] = chart_df["weighted_rate"].clip(upper=100) if cap_at_100 else chart_df["weighted_rate"]

    palette = (PASTEL * ((len(chart_df) // len(PASTEL)) + 1))[:len(chart_df)]
    color_map = dict(zip(chart_df["room"], palette))

    fig = px.bar(
        chart_df,
        x="rate_show",
        y="room",
        orientation="h",
        text=chart_df["weighted_rate"].round(1),
        color="room",
        color_discrete_map=color_map,
        labels={"rate_show": "예약률(가중, %)", "room": "회의실"},
        hover_data={
            "reserved_h":":.1f",
            "avail_h":":.1f",
            "weighted_rate":":.1f",
            "simple_avg_rate":":.1f",
            "volatility":":.1f"
        }
    )
    fig.update_traces(textposition="outside", cliponaxis=False)
    x_max = 105 if cap_at_100 else max(105, float(chart_df["weighted_rate"].max()) * 1.1)
    fig.update_layout(
        showlegend=False,
        xaxis=dict(range=[0, x_max]),
        yaxis=dict(autorange="reversed"),
        margin=dict(l=10, r=30, t=10, b=10),
        height=min(900, 30 * len(chart_df) + 140),
    )
    st.plotly_chart(fig, use_container_width=True)

    # 추이(전체 기간 or 최근3개월) — 전체 선택일 때도 제공
    st.markdown("### 📈 예약률 추이")
    if trend_mode == "최근 3개월":
        last_month = pd.to_datetime(months[-1] + "-01")
        m2 = add_months(last_month, -2).strftime("%Y-%m")
        m1 = add_months(last_month, -1).strftime("%Y-%m")
        m0 = last_month.strftime("%Y-%m")
        trend_months = [m2, m1, m0]
    else:
        trend_months = months_in_period

    rooms_all = sorted(cur2["room"].unique())
    default_rooms = list(cur2.head(min(5, len(cur2)))["room"])
    sel_rooms = st.multiselect("추이를 볼 회의실 선택", options=rooms_all, default=default_rooms)

    trend = monthly_full[monthly_full["month"].isin(trend_months) & monthly_full["room"].isin(sel_rooms)].copy()
    trend["month_dt"] = pd.to_datetime(trend["month"] + "-01")

    fig2 = px.line(
        trend.sort_values(["room", "month_dt"]),
        x="month_dt",
        y="month_rate",
        color="room",
        markers=True,
        labels={"month_dt":"월", "month_rate":"예약률(월, %)", "room":"회의실"}
    )
    fig2.update_layout(
        xaxis=dict(tickformat="%Y-%m"),
        yaxis=dict(range=[0, min(110, max(10, float(trend["month_rate"].max()) * 1.2))]),
        margin=dict(l=10, r=10, t=10, b=10),
        height=420
    )
    st.plotly_chart(fig2, use_container_width=True)

    # 테이블
    st.markdown("### 📋 회의실별 예약률 테이블 (전체 기간)")
    table_out = cur2.rename(columns={
        "room":"회의실",
        "reserved_h":"예약시간(시간)",
        "avail_h":"가용시간(시간)",
        "weighted_rate":"전체예약률(가중,%)",
        "simple_avg_rate":"월평균예약률의평균(단순,%)",
        "volatility":"변동성(σ)"
    }).copy()

    for col in ["예약시간(시간)","가용시간(시간)","전체예약률(가중,%)","월평균예약률의평균(단순,%)","변동성(σ)"]:
        table_out[col] = pd.to_numeric(table_out[col], errors="coerce").fillna(0.0).round(1)

    table_show = table_out[[
        "회의실","예약시간(시간)","가용시간(시간)",
        "전체예약률(가중,%)","월평균예약률의평균(단순,%)",
        "판단(가중)","판단(단순월평균)","변동성(σ)"
    ]].sort_values("전체예약률(가중,%)", ascending=False)

    st.dataframe(table_show, use_container_width=True)

    # 엑셀 다운로드(전체)
    kpi_sheet = pd.DataFrame([{
        "기간": period_label,
        "전체 예약률(가중평균,%)": round(overall_weighted_kpi, 1),
        "월 평균 예약률의 평균(단순,%)": round(overall_simple_kpi, 1),
        "회의실 수": int(cur2["room"].nunique()),
        "시간대": time_mode
    }])

    trend_all_sheet = monthly_full.rename(columns={
        "month":"월",
        "room":"회의실",
        "month_rate":"예약률(월,%)"
    })[["월","회의실","예약률(월,%)"]].copy()
    trend_all_sheet["예약률(월,%)"] = trend_all_sheet["예약률(월,%)"].round(1)

    sheets = {
        "KPI_전체": kpi_sheet,
        "전체_회의실별(가중vs단순)": table_show,
        "월별_추이(전체)": trend_all_sheet,
    }
    excel_bytes = df_to_excel_bytes(sheets)

    st.download_button(
        label="📥 결과 엑셀 다운로드(전체: KPI+테이블+추이)",
        data=excel_bytes,
        file_name="회의실_예약률_전체기간.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    # ====== 월 선택(기존처럼 월 예약률만 보여줌) ======
    cur = monthly[monthly["month"] == selected].copy()
    if cur.empty:
        st.warning("선택한 월에 계산 가능한 예약이 없어요(근무일/시간대 기준).")
        st.stop()

    cur["reserved_h"] = cur["reserved_min"] / 60
    cur["avail_h"] = (cur["month"].map(avail_map) / 60).astype(float)
    cur = cur.sort_values("rate", ascending=False).reset_index(drop=True)

    # KPI
    st.subheader("2) 대시보드")
    avail_min_sel = avail_map[selected]
    workday_count_sel = int(avail_min_sel / window_minutes_per_day(windows))
    avg_rate_cur = float(cur["rate"].mean()) if len(cur) else 0.0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("대상 월", selected)
    c2.metric("근무일 수", f"{workday_count_sel}일")
    c3.metric("회의실 수", f"{cur['room'].nunique()}개")
    c4.metric("이번 달 평균 예약률", f"{avg_rate_cur:.1f}%")

    # 막대그래프
    chart_df = cur.copy()
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
        labels={"rate_show":"예약률(%)", "room":"회의실"},
        hover_data={"reserved_h":":.1f", "avail_h":":.1f", "rate":":.1f"}
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

    st.markdown("### 📋 회의실별 예약률 테이블 (선택 월)")
    table_out = cur.rename(columns={
        "room":"회의실",
        "reserved_h":"예약시간(시간)",
        "avail_h":"가용시간(시간)",
        "rate":"예약률(%)"
    }).copy()
    table_out["예약시간(시간)"] = table_out["예약시간(시간)"].round(1)
    table_out["가용시간(시간)"] = table_out["가용시간(시간)"].round(1)
    table_out["예약률(%)"] = table_out["예약률(%)"].round(1)

    st.dataframe(table_out[["회의실","예약시간(시간)","가용시간(시간)","예약률(%)"]], use_container_width=True)

    sheets = {f"예약률_{selected}": table_out}
    excel_bytes = df_to_excel_bytes(sheets)
    st.download_button(
        label="📥 결과 엑셀 다운로드(선택 월)",
        data=excel_bytes,
        file_name=f"회의실_예약률_{selected}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# -----------------------------
# Debug
# -----------------------------
if show_debug:
    st.markdown("### 🧪 디버그")
    st.write("선택:", selected)
    st.write("기간:", period_start, "~", period_end)
    st.write("daily rows:", len(daily))
    st.dataframe(monthly.head(50), use_container_width=True)
