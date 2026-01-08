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

    ex

