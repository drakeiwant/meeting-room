import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
from datetime import datetime, time, timedelta

try:
    import holidays  # pip install holidays
except Exception:
    holidays = None

st.set_page_config(page_title="회의실 예약률 대시보드", layout="wide")

# -----------------------------
# UI Header
# -----------------------------
st.title("🏢 회의실별 월간 예약률 대시보드")
st.caption("Raw 엑셀 업로드 → 월 선택 → (주말/공휴일 제외) 가용시간 대비 예약률 자동 계산")

uploaded = st.file_uploader("📂 예약현황 Raw 엑셀 업로드 (.xlsx)", type=["xlsx"])

# -----------------------------
# Sidebar Settings
# -----------------------------
with st.sidebar:
    st.header("설정")
    work_start = st.time_input("업무 시작", value=time(9, 0))
    work_end = st.time_input("업무 종료", value=time(18, 0))

    st.divider()
    st.subheader("가용일 계산 옵션")
    exclude_weekends = st.checkbox("주말 제외", value=True)
    exclude_holidays = st.checkbox("공휴일 제외(대한민국)", value=True)
    country = st.selectbox("공휴일 국가", options=["KR"], index=0, disabled=True)

    st.divider()
    st.subheader("대시보드 연출")
    show_animation = st.checkbox("계산 애니메이션(로딩 느낌)", value=True)
    highlight_top = st.checkbox("TOP 회의실 강조", value=True)

# -----------------------------
# Helpers
# -----------------------------
def parse_datetime(date_series, time_series):
    """
    Robust-ish parsing:
    - date: excel date / 'YYYY-MM-DD' / 'YYYY.MM.DD' 등
    - time: 'HH:MM' / excel time(float 0~1) 등
    """
    d = pd.to_datetime(date_series, errors="coerce")

    t_raw = time_series.copy()

    # Excel time floats (0~1)
    t_num = pd.to_numeric(t_raw, errors="coerce")
    is_excel_time = t_num.notna() & (t_num.between(0, 1))
    t_excel_td = pd.to_timedelta(np.where(is_excel_time, t_num * 24 * 3600, np.nan), unit="s")

    # String time parsing
    # (pandas가 "09:00" 등을 datetime으로 잘 파싱하는 편)
    t_str_time = pd.to_datetime(t_raw.astype(str), errors="coerce").dt.time
    t_str_td = pd.to_timedelta(pd.Series(t_str_time).astype(str), errors="coerce")

    t = t_excel_td.copy()
    t = t.fillna(t_str_td)

    return d + t

def month_range(dt):
    first = dt.replace(day=1)
    if first.month == 12:
        nxt = first.replace(year=first.year + 1, month=1)
    else:
        nxt = first.replace(month=first.month + 1)
    return first, nxt

def get_kr_holidays(years):
    """
    대한민국 공휴일 set(date) 반환.
    holidays 패키지가 없으면 빈 set 반환(=공휴일 제외 기능 비활성)
    """
    if holidays is None:
        return set()
    kr = holidays.KR(years=years)
    return set(kr.keys())

def is_workday(d, holiday_set, exclude_weekends=True, exclude_holidays=True):
    # d: datetime.date
    if exclude_weekends and d.weekday() >= 5:  # 5=Sat, 6=Sun
        return False
    if exclude_holidays and (d in holiday_set):
        return False
    return True

def business_days_in_month(month_start, month_end, holiday_set, exclude_weekends=True, exclude_holidays=True):
    days = pd.date_range(month_start, month_end - pd.Timedelta(days=1), freq="D")
    out = []
    for ts in days:
        d = ts.date()
        if is_workday(d, holiday_set, exclude_weekends, exclude_holidays):
            out.append(ts)
    return pd.DatetimeIndex(out)

def minutes_between(t1: time, t2: time) -> float:
    dummy = datetime(2000, 1, 1)
    return (datetime.combine(dummy.date(), t2) - datetime.combine(dummy.date(), t1)).total_seconds() / 60

# -----------------------------
# Main
# -----------------------------
if not uploaded:
    st.info("왼쪽에서 Raw 엑셀 파일(.xlsx)을 업로드하면 대시보드가 생성돼요.")
    st.stop()

df = pd.read_excel(uploaded)

st.subheader("1) Raw 미리보기")
st.dataframe(df.head(30), use_container_width=True)

required = ["회의실명", "날짜", "예약시작 시간", "예약끝나는 시간"]
missing = [c for c in required if c not in df.columns]
if missing:
    st.error(f"필수 컬럼이 없어요: {missing}\n\n내일 실제 파일 기준으로 컬럼명만 맞춰주면 바로 동작해요.")
    st.stop()

# Parse datetimes
df = df.copy()
df["start_dt"] = parse_datetime(df["날짜"], df["예약시작 시간"])
df["end_dt"] = parse_datetime(df["날짜"], df["예약끝나는 시간"])

bad = df["start_dt"].isna() | df["end_dt"].isna()
if bad.any():
    st.warning("일부 행에서 날짜/시간 파싱이 실패했어요. 아래 샘플 확인:")
    st.dataframe(df.loc[bad, ["회의실명", "날짜", "예약시작 시간", "예약끝나는 시간"]].head(30), use_container_width=True)

df = df.loc[~bad].copy()

# 종료<=시작 제거
invalid_order = df["end_dt"] <= df["start_dt"]
if invalid_order.any():
    st.warning("종료시간이 시작시간보다 빠르거나 같은 행이 있어 제외했어요.")
    df = df.loc[~invalid_order].copy()

# 월 리스트
df["month"] = df["start_dt"].dt.to_period("M").astype(str)
months = sorted(df["month"].unique())
if not months:
    st.error("유효한 데이터가 없어요. 날짜/시간 형식을 확인해줘.")
    st.stop()

st.subheader("2) 월 선택")
selected_month = st.selectbox("집계할 월(YYYY-MM)", months, index=len(months) - 1)

# 선택 월 데이터
df_m = df[df["month"] == selected_month].copy()

# --------------- OPTION 1: 주말/공휴일 제외 ---------------
ms = pd.to_datetime(selected_month + "-01")
month_start, month_end = month_range(ms)
years = sorted({month_start.year, (month_end - timedelta(days=1)).year})
holiday_set = get_kr_holidays(years) if exclude_holidays else set()

# 예약일 자체가 비근무일이면 제외
df_m["date_only"] = df_m["start_dt"].dt.date
df_m["is_workday"] = df_m["date_only"].apply(
    lambda d: is_workday(d, holiday_set, exclude_weekends=exclude_weekends, exclude_holidays=exclude_holidays)
)
before_days = len(df_m)
df_m = df_m[df_m["is_workday"]].copy()
dropped_days = before_days - len(df_m)

# --------------- OPTION 2: 09~18로 클리핑 ---------------
def clip_to_work_hours(row):
    day = row["start_dt"].date()
    ws = datetime.combine(day, work_start)
    we = datetime.combine(day, work_end)

    # 요청 2: 09 이전 시작 -> 09로, 18 이후 종료 -> 18로
    s = max(row["start_dt"], ws)
    e = min(row["end_dt"], we)

    if e <= s:
        return None
    return (s, e)

df_m["clipped"] = df_m.apply(clip_to_work_hours, axis=1)
before_clip = len(df_m)
df_m = df_m[df_m["clipped"].notna()].copy()
dropped_clip = before_clip - len(df_m)

df_m["s"] = df_m["clipped"].apply(lambda x: x[0])
df_m["e"] = df_m["clipped"].apply(lambda x: x[1])
df_m["reserved_minutes"] = (df_m["e"] - df_m["s"]).dt.total_seconds() / 60

# --------------- OPTION 3: 예약률 계산 (월간 총 예약시간 / 월간 총 가용시간) ---------------
biz_days = business_days_in_month(
    month_start, month_end, holiday_set,
    exclude_weekends=exclude_weekends,
    exclude_holidays=exclude_holidays
)
daily_minutes = minutes_between(work_start, work_end)
available_minutes_per_room = len(biz_days) * daily_minutes

# 집계
usage = (
    df_m.groupby("회의실명", as_index=False)["reserved_minutes"]
    .sum()
    .rename(columns={"reserved_minutes": "예약시간(분)"})
)

usage["예약시간(시간)"] = usage["예약시간(분)"] / 60
usage["가용시간(시간)"] = available_minutes_per_room / 60
usage["예약률(%)"] = (usage["예약시간(분)"] / available_minutes_per_room) * 100
usage = usage.sort_values("예약률(%)", ascending=False).reset_index(drop=True)

# -----------------------------
# A bit of "living" dashboard
# -----------------------------
st.subheader("3) 대시보드")

# 애니메이션(로딩 느낌)
if show_animation:
    ph = st.empty()
    prog = st.progress(0)
    for i in range(1, 101, 10):
        ph.markdown(f"🧮 계산 중… **{i}%**")
        prog.progress(i)
        # 너무 오래 걸리면 귀찮으니 짧게
        import time as _t
        _t.sleep(0.03)
    ph.empty()
    prog.empty()

# KPI
total_rooms = usage["회의실명"].nunique()
total_reserved_h = usage["예약시간(시간)"].sum()
avg_rate = usage["예약률(%)"].mean() if len(usage) else 0.0
top_room = usage.iloc[0]["회의실명"] if len(usage) else "-"
top_rate = usage.iloc[0]["예약률(%)"] if len(usage) else 0.0

k1, k2, k3, k4 = st.columns(4)
k1.metric("대상 월", selected_month)
k2.metric("회의실 수", f"{total_rooms}개")
k3.metric("총 예약시간", f"{total_reserved_h:.1f}시간")
k4.metric("최고 예약률", f"{top_rate:.1f}%", help=f"회의실: {top_room}")

# 계산/제외 요약
with st.expander("⚙️ 계산/제외 요약 보기"):
    st.write(f"- 주말 제외: {exclude_weekends}")
    st.write(f"- 공휴일 제외(KR): {exclude_holidays} (holidays 패키지 {'사용' if holidays else '미사용'})")
    st.write(f"- 선택 월 근무일 수: **{len(biz_days)}일**")
    st.write(f"- 1일 가용시간: **{daily_minutes/60:.1f}시간** ({work_start.strftime('%H:%M')}~{work_end.strftime('%H:%M')})")
    st.write(f"- 회의실당 월 가용시간: **{available_minutes_per_room/60:.1f}시간**")
    st.write(f"- 비근무일(주말/공휴일)로 제외된 예약 건수: **{dropped_days}건**")
    st.write(f"- 업무시간 클리핑 후 0분이 되어 제외된 예약 건수: **{dropped_clip}건**")

# 레이아웃: 탭으로 생동감 + 탐색
tab1, tab2, tab3 = st.tabs(["📊 차트", "📋 테이블", "🔎 원천(클리핑 후)"])

with tab1:
    # 인터랙티브 바차트(hover + 정렬)
    chart_df = usage.copy()
    chart_df["예약률(%)"] = chart_df["예약률(%)"].round(1)

    base = alt.Chart(chart_df).mark_bar().encode(
        x=alt.X("회의실명:N", sort="-y", title="회의실"),
        y=alt.Y("예약률(%):Q", title="예약률(%)"),
        tooltip=[
            alt.Tooltip("회의실명:N", title="회의실"),
            alt.Tooltip("예약률(%):Q", title="예약률(%)"),
            alt.Tooltip("예약시간(시간):Q", title="예약시간(시간)", format=".1f"),
            alt.Tooltip("가용시간(시간):Q", title="가용시간(시간)", format=".1f"),
        ],
    ).transform_calculate(
        **{"예약률(%)": "datum['예약률(%)']"}  # altair 필드명 이슈 방지
    )

    # TOP 강조(색은 altair 기본, 대신 라벨/선으로 “움직임” 느낌)
    text = alt.Chart(chart_df).mark_text(
        dy=-8
    ).encode(
        x=alt.X("회의실명:N", sort="-y"),
        y=alt.Y("예약률(%) :Q"),
        text=alt.Text("예약률(%) :Q", format=".1f")
    )

    if highlight_top and len(chart_df) > 0:
        top_name = chart_df.iloc[0]["회의실명"]
        rule = alt.Chart(pd.DataFrame({"회의실명": [top_name]})).mark_rule(strokeDash=[6, 4]).encode(
            x="회의실명:N"
        )
        st.altair_chart((base + text + rule).properties(height=420).interactive(), use_container_width=True)
    else:
        st.altair_chart((base + text).properties(height=420).interactive(), use_container_width=True)

    # “살아있는 느낌” 추가: TOP 예약률 진행바
    if len(usage) > 0:
        st.markdown("#### 🏆 TOP 회의실 예약률")
        st.write(f"**{top_room}**  —  {top_rate:.1f}%")
        st.progress(min(int(round(top_rate)), 100))

with tab2:
    st.dataframe(
        usage[["회의실명", "예약시간(시간)", "가용시간(시간)", "예약률(%)"]]
        .round({"예약시간(시간)": 1, "가용시간(시간)": 1, "예약률(%)": 1}),
        use_container_width=True
    )

with tab3:
    show_cols = ["회의실명", "start_dt", "end_dt", "s", "e", "reserved_minutes"]
    st.dataframe(
        df_m[show_cols].rename(columns={"reserved_minutes": "예약시간(분)"}).head(500),
        use_container_width=True
    )
