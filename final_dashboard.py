
# =========================================================
# 보험 서류 스캔 관리 대시보드 - 안정화 최종 버전
# =========================================================
# 주요 특징
# - 엑셀 컬럼 자동 인식 (접수일 / 서류 컬럼)
# - 개인정보동의서 집계 제외
# - 미스캔 = FA고지 + 비교설명 + 완판만 계산
# - 월별 가로 Pivot 테이블
# - 계층 리포트 정렬
# - 표 줄바꿈 제거
# - 오류 방지 로직 다수 추가
# =========================================================

import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="보험 스캔 관리 대시보드", layout="wide")

EXCEL_FILE = "insurance_data.xlsx"

# -------------------------------------------------
# UI 스타일
# -------------------------------------------------
st.markdown(
"""
<style>
table {white-space: nowrap;}
.block-container {padding-top:1rem;padding-bottom:0rem;}
</style>
""",
unsafe_allow_html=True
)

# -------------------------------------------------
# 컬럼 자동 찾기
# -------------------------------------------------
def find_column(df, keywords):

    for col in df.columns:
        c = str(col).strip()
        for k in keywords:
            if k in c:
                return col
    return None

# -------------------------------------------------
# 데이터 로드
# -------------------------------------------------
@st.cache_data
def load_data():

    if not os.path.exists(EXCEL_FILE):
        st.error(f"엑셀 파일을 찾을 수 없습니다: {EXCEL_FILE}")
        st.stop()

    df = pd.read_excel(EXCEL_FILE)

    df.columns = df.columns.astype(str).str.strip()

    # 접수일 자동 탐지
    date_col = find_column(df, ["접수일", "접수일자", "접수일시"])

    if date_col is None:
        st.error(f"접수일 컬럼을 찾지 못했습니다. 현재 컬럼: {list(df.columns)}")
        st.stop()

    df["접수일"] = pd.to_datetime(df[date_col], errors="coerce")

    df["월"] = df["접수일"].dt.to_period("M").astype(str)

    return df

# -------------------------------------------------
# 미스캔 계산
# -------------------------------------------------
def calculate_miss_scan(df):

    df = df.copy()

    fa_col = find_column(df, ["FA"])
    comp_col = find_column(df, ["비교"])
    sale_col = find_column(df, ["완판"])

    if fa_col is None or comp_col is None or sale_col is None:
        st.error("서류 컬럼을 찾을 수 없습니다.")
        st.stop()

    df["FA_miss"] = df[fa_col].astype(str).str.contains("미스캔", na=False)
    df["비교_miss"] = df[comp_col].astype(str).str.contains("미스캔", na=False)
    df["완판_miss"] = df[sale_col].astype(str).str.contains("미스캔", na=False)

    df["미스캔"] = (
        df["FA_miss"].astype(int)
        + df["비교_miss"].astype(int)
        + df["완판_miss"].astype(int)
    )

    return df

# -------------------------------------------------
# 대상 계산 (개인정보 제외)
# -------------------------------------------------
def calculate_targets(df):

    df = df.copy()

    fa_col = find_column(df, ["FA"])
    comp_col = find_column(df, ["비교"])
    sale_col = find_column(df, ["완판"])

    df["FA_target"] = df[fa_col].notna()
    df["비교_target"] = df[comp_col].notna()
    df["완판_target"] = df[sale_col].notna()

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

    hierarchy = []

    for c in df.columns:
        if "부문" in c:
            hierarchy.append(c)
        elif "총괄" in c:
            hierarchy.append(c)
        elif "부서" in c:
            hierarchy.append(c)
        elif "소속" in c or "영업" in c:
            hierarchy.append(c)

    if len(hierarchy) == 0:
        st.warning("계층 컬럼을 찾지 못했습니다.")
        return pd.DataFrame()

    agg = (
        df.groupby(hierarchy)
        .agg(
            계약수=("월", "count"),
            미스캔=("미스캔", "sum"),
            대상=("전체대상", "sum")
        )
        .reset_index()
    )

    agg["스캔율"] = ((agg["대상"] - agg["미스캔"]) / agg["대상"] * 100).round(1)

    return agg.sort_values(hierarchy)

# -------------------------------------------------
# 월별 Pivot
# -------------------------------------------------
def monthly_pivot(df):

    dept_col = None

    for c in df.columns:
        if "부문" in c or "조직" in c:
            dept_col = c
            break

    if dept_col is None:
        dept_col = df.columns[0]

    p = (
        df.groupby([dept_col, "월"])["미스캔"]
        .sum()
        .reset_index()
        .pivot(index=dept_col, columns="월", values="미스캔")
        .fillna(0)
    )

    return p.reset_index()

# -------------------------------------------------
# KPI
# -------------------------------------------------
def show_kpi(df):

    total_contract = len(df)
    total_miss = int(df["미스캔"].sum())
    total_target = int(df["전체대상"].sum())

    c1, c2, c3 = st.columns(3)

    c1.metric("총 계약건", total_contract)
    c2.metric("총 미스캔", total_miss)
    c3.metric("전체 대상", total_target)

# -------------------------------------------------
# 메인 대시보드
# -------------------------------------------------
def dashboard():

    df = load_data()

    df = calculate_miss_scan(df)
    df = calculate_targets(df)

    st.title("보험 서류 스캔 관리 대시보드")

    show_kpi(df)

    st.divider()

    tab1, tab2 = st.tabs(["계층 리포트", "월별 집계"])

    with tab1:

        rpt = build_hierarchy_report(df)

        st.dataframe(
            rpt,
            use_container_width=True,
            height=700
        )

    with tab2:

        pivot = monthly_pivot(df)

        st.dataframe(
            pivot,
            use_container_width=True
        )

# -------------------------------------------------
# 실행
# -------------------------------------------------
dashboard()
