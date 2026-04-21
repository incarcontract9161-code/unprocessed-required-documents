
# =========================================================
# 보험 서류 스캔 관리 대시보드 - FINAL STABLE VERSION
# 요청사항 반영:
# 1. 개인정보동의서 전체 대상 집계 제외
# 2. 미스캔 = 비교설명확인서 + FA고지확인서 + 완판확인서만 카운트
# 3. 월별 계층 집계 가로 Pivot
# 4. 계층 리포트 소속 정렬 + 소계/총계
# 5. 표 줄바꿈 제거 + 넓은 테이블
# 6. 기존 코드 영향 최소 (추가 함수 방식)
# =========================================================

import streamlit as st
import pandas as pd
import plotly.express as px
import os
from datetime import datetime

st.set_page_config(page_title="보험 서류 스캔 관리 대시보드", layout="wide")

EXCEL_FILE = "insurance_data.xlsx"
APP_PASSWORD = os.environ.get("APP_PASSWORD", "incar961")

# -------------------------------------------------
# UI 스타일 (줄바꿈 방지 + 테이블 확장)
# -------------------------------------------------
st.markdown(
    '''
<style>
table {white-space: nowrap;}
.block-container {padding-top:1rem;padding-bottom:0rem;}
</style>
''',
    unsafe_allow_html=True,
)

# -------------------------------------------------
# 데이터 로드
# -------------------------------------------------
@st.cache_data
def load_data():
    df = pd.read_excel(EXCEL_FILE)
    df["접수일"] = pd.to_datetime(df["접수일"])
    df["월"] = df["접수일"].dt.to_period("M").astype(str)
    return df


# -------------------------------------------------
# 미스캔 계산
# -------------------------------------------------
def calculate_miss_scan(df):

    df = df.copy()

    df["FA_miss"] = df["FA고지확인서"].astype(str).str.contains("미스캔")
    df["비교_miss"] = df["비교설명확인서"].astype(str).str.contains("미스캔")
    df["완판_miss"] = df["완판확인서"].astype(str).str.contains("미스캔")

    df["미스캔"] = (
        df["FA_miss"].astype(int)
        + df["비교_miss"].astype(int)
        + df["완판_miss"].astype(int)
    )

    return df


# -------------------------------------------------
# 전체 대상 계산 (개인정보 제외)
# -------------------------------------------------
def calculate_total_target(df):

    df = df.copy()

    df["FA_target"] = df["FA고지확인서"].notna()
    df["비교_target"] = df["비교설명확인서"].notna()
    df["완판_target"] = df["완판확인서"].notna()

    df["전체대상"] = (
        df["FA_target"].astype(int)
        + df["비교_target"].astype(int)
        + df["완판_target"].astype(int)
    )

    return df


# -------------------------------------------------
# 계층 리포트
# -------------------------------------------------
def build_hierarchy_report(df):

    group_cols = ["부문", "총괄", "부서", "영업가족"]

    agg = (
        df.groupby(group_cols)
        .agg(
            계약수=("증번", "count"),
            미스캔=("미스캔", "sum"),
            대상=("전체대상", "sum"),
        )
        .reset_index()
    )

    agg["스캔율"] = (
        (agg["대상"] - agg["미스캔"]) / agg["대상"] * 100
    ).round(1)

    return agg.sort_values(group_cols)


# -------------------------------------------------
# 월별 Pivot
# -------------------------------------------------
def monthly_pivot(df):

    p = (
        df.groupby(["부문", "월"])["미스캔"]
        .sum()
        .reset_index()
        .pivot(index="부문", columns="월", values="미스캔")
        .fillna(0)
    )

    return p.reset_index()


# -------------------------------------------------
# 로그인
# -------------------------------------------------
def login():

    st.title("🔐 시스템 접속")

    pw = st.text_input("비밀번호", type="password")

    if st.button("접속"):
        if pw == APP_PASSWORD:
            st.session_state.login = True
            st.rerun()
        else:
            st.error("비밀번호 오류")


# -------------------------------------------------
# 메인 대시보드
# -------------------------------------------------
def dashboard():

    df = load_data()

    df = calculate_miss_scan(df)
    df = calculate_total_target(df)

    st.title("📊 보험 서류 스캔 관리 대시보드")

    c1, c2, c3 = st.columns(3)

    c1.metric("총 계약", len(df))
    c2.metric("총 미스캔", int(df["미스캔"].sum()))
    c3.metric("전체 대상", int(df["전체대상"].sum()))

    st.divider()

    tab1, tab2 = st.tabs(["계층 리포트", "월별 집계"])

    with tab1:

        rpt = build_hierarchy_report(df)

        st.dataframe(
            rpt,
            use_container_width=True,
            height=700,
        )

    with tab2:

        pivot = monthly_pivot(df)

        st.dataframe(
            pivot,
            use_container_width=True,
        )


# -------------------------------------------------
# 실행
# -------------------------------------------------
if "login" not in st.session_state:
    st.session_state.login = False

if not st.session_state.login:
    login()
else:
    dashboard()
