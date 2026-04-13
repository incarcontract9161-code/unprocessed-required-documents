import streamlit as st
import pandas as pd
import os
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

# ==========================================
# 페이지 설정
# ==========================================
st.set_page_config(page_title="보험 서류 스캔 관리 대시보드", layout="wide", page_icon="📊")

# ==========================================
# 설정
# ==========================================
EXCEL_FILE = "insurance_data.xlsx"  # GitHub에 업로드할 엑셀 파일명

# ==========================================
# 데이터 로딩
# ==========================================
@st.cache_data(ttl=300)  # 5분마다 자동 갱신
def load_data():
    """GitHub의 엑셀 파일을 읽어서 DataFrame 반환"""
    if not os.path.exists(EXCEL_FILE):
        st.error(f"⚠️ '{EXCEL_FILE}' 파일이 없습니다. GitHub에 엑셀 파일을 업로드해주세요.")
        st.stop()
    
    try:
        df = pd.read_excel(EXCEL_FILE)
        
        # 날짜 및 필드 전처리
        df["보험시작일_dt"] = pd.to_datetime(df["보험시작일"], errors="coerce")
        df["월_피리어드"] = df["보험시작일_dt"].dt.to_period("M").astype(str)
        df["FA고지_c"] = df["FA고지"].fillna("").astype(str).str.strip()
        df["비교설명_c"] = df["비교설명"].fillna("").astype(str).str.strip()
        df["완전판매_c"] = df["완전판매"].fillna("").astype(str).str.strip()
        
        return df
    except Exception as e:
        st.error(f"❌ 엑셀 파일 읽기 오류: {e}")
        st.stop()

def get_file_update_time():
    """엑셀 파일의 마지막 수정 시간 반환"""
    if os.path.exists(EXCEL_FILE):
        timestamp = os.path.getmtime(EXCEL_FILE)
        return datetime.fromtimestamp(timestamp).strftime("%Y-%m-%d %H:%M:%S")
    return "알 수 없음"

# ==========================================
# 집계 함수
# ==========================================
def _miss(s):
    return (s.fillna("").astype(str).str.strip() == "미스캔").sum()

def _miss_cs(s):
    s2 = s.fillna("").astype(str).str.strip()
    return ((s2 != "해당없음") & (s2 == "미스캔")).sum()

# ==========================================
# 메인 앱
# ==========================================
def main():
    # 헤더
    st.title("📊 보험 서류 스캔 관리 대시보드")
    
    # 데이터 로드
    df = load_data()
    
    # 상태바
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        st.success(f"✅ 총 **{len(df):,}건**의 데이터 로드 완료")
    with col2:
        st.info(f"📅 기준: **{get_file_update_time()}**")
    with col3:
        if st.button("🔄 새로고침"):
            st.cache_data.clear()
            st.rerun()
    
    st.divider()
    
    # 사이드바 필터
    st.sidebar.header("🔍 필터 옵션")
    
    # 월 선택
    all_months = sorted(df["월_피리어드"].dropna().unique().tolist(), reverse=True)
    if not all_months:
        st.warning("⚠️ 날짜 데이터가 없습니다.")
        return
    
    sel_months = st.sidebar.multiselect(
        "📅 조회 월 선택",
        all_months,
        default=all_months[:1]
    )
    
    if not sel_months:
        st.warning("⚠️ 최소 1개 월을 선택하세요.")
        return
    
    df_sel = df[df["월_피리어드"].isin(sel_months)].copy()
    
    # 부문 필터
    if "부문" in df_sel.columns:
        all_depts = ["전체"] + sorted(df_sel["부문"].dropna().unique().tolist())
        sel_dept = st.sidebar.selectbox("🏢 부문", all_depts)
        if sel_dept != "전체":
            df_sel = df_sel[df_sel["부문"] == sel_dept]
    
    # 총괄 필터
    if "총괄" in df_sel.columns:
        all_teams = ["전체"] + sorted(df_sel["총괄"].dropna().unique().tolist())
        sel_team = st.sidebar.selectbox("👥 총괄", all_teams)
        if sel_team != "전체":
            df_sel = df_sel[df_sel["총괄"] == sel_team]
    
    st.sidebar.divider()
    st.sidebar.caption(f"선택된 데이터: **{len(df_sel):,}건**")
    
    # ==========================================
    # 탭 구성
    # ==========================================
    tab1, tab2, tab3, tab4 = st.tabs(["📊 종합 현황", "📈 조직별 분석", "📋 상세 데이터", "ℹ️ 안내"])
    
    # ── TAB 1: 종합 현황 ──
    with tab1:
        st.header("📊 종합 현황")
        
        # 전체 통계
        total = len(df_sel)
        fa_miss = _miss(df_sel["FA고지_c"])
        bi_miss = _miss(df_sel["비교설명_c"])
        cs_miss = _miss_cs(df_sel["완전판매_c"])
        total_miss = fa_miss + bi_miss + cs_miss
        miss_rate = (total_miss / total * 100) if total > 0 else 0
        
        # KPI 카드
        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("총 계약 건수", f"{total:,}")
        col2.metric("FA고지 미스캔", f"{fa_miss:,}")
        col3.metric("비교설명 미스캔", f"{bi_miss:,}")
        col4.metric("완전판매 미스캔", f"{cs_miss:,}")
        col5.metric("총 미스캔", f"{total_miss:,}", delta=f"{miss_rate:.1f}%")
        
        st.divider()
        
        # 차트
        col1, col2 = st.columns(2)
        
        with col1:
            # 서류별 미스캔 비율
            pie_data = pd.DataFrame({
                "서류": ["FA고지", "비교설명", "완전판매"],
                "건수": [fa_miss, bi_miss, cs_miss]
            })
            fig = px.pie(pie_data, values="건수", names="서류", 
                        title="서류별 미스캔 비율",
                        color_discrete_sequence=px.colors.qualitative.Set3,
                        hole=0.3)
            fig.update_traces(textposition='inside', textinfo='percent+label')
            st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            # 완료 vs 미스캔 비교
            status_data = pd.DataFrame({
                "상태": ["완료", "미스캔"],
                "건수": [total - total_miss, total_miss]
            })
            fig = px.bar(status_data, x="상태", y="건수",
                        title="전체 처리 현황",
                        color="상태",
                        color_discrete_map={"완료": "#2ecc71", "미스캔": "#e74c3c"})
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig, use_container_width=True)
        
        # 월별 트렌드
        if len(sel_months) > 1:
            st.subheader("📅 월별 미스캔 추이")
            monthly_data = []
            for month in sorted(sel_months):
                df_month = df_sel[df_sel["월_피리어드"] == month]
                monthly_data.append({
                    "월": month,
                    "FA고지": _miss(df_month["FA고지_c"]),
                    "비교설명": _miss(df_month["비교설명_c"]),
                    "완전판매": _miss_cs(df_month["완전판매_c"])
                })
            
            monthly_df = pd.DataFrame(monthly_data)
            monthly_melted = monthly_df.melt(id_vars="월", var_name="서류", value_name="건수")
            
            fig = px.line(monthly_melted, x="월", y="건수", color="서류",
                         markers=True, title="월별 서류 미스캔 추이")
            st.plotly_chart(fig, use_container_width=True)
    
    # ── TAB 2: 조직별 분석 ──
    with tab2:
        st.header("📈 조직별 미스캔 분석")
        
        # 조직 레벨 선택
        org_level = st.radio(
            "조직 단위 선택",
            ["부문", "총괄", "부서", "영업가족"],
            horizontal=True
        )
        
        if org_level in df_sel.columns:
            # 조직별 집계
            org_data = []
            for org_name in df_sel[org_level].dropna().unique():
                df_org = df_sel[df_sel[org_level] == org_name]
                fa = _miss(df_org["FA고지_c"])
                bi = _miss(df_org["비교설명_c"])
                cs = _miss_cs(df_org["완전판매_c"])
                total_org = len(df_org)
                total_miss_org = fa + bi + cs
                
                org_data.append({
                    org_level: org_name,
                    "총건수": total_org,
                    "FA고지": fa,
                    "비교설명": bi,
                    "완전판매": cs,
                    "총미스캔": total_miss_org,
                    "미처리율": round(total_miss_org / total_org * 100, 1) if total_org > 0 else 0
                })
            
            org_df = pd.DataFrame(org_data).sort_values("총미스캔", ascending=False)
            
            # 테이블 표시
            st.dataframe(
                org_df,
                use_container_width=True,
                hide_index=True
            )
            
            # 차트
            col1, col2 = st.columns(2)
            
            with col1:
                # Top 10 미스캔
                top_10 = org_df.head(10)
                fig = px.bar(top_10, x=org_level, y="총미스캔",
                           title=f"미스캔 TOP 10 ({org_level})",
                           color="총미스캔",
                           color_continuous_scale="Reds")
                fig.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)
            
            with col2:
                # 서류별 비교
                org_melted = org_df.head(10).melt(
                    id_vars=[org_level],
                    value_vars=["FA고지", "비교설명", "완전판매"],
                    var_name="서류",
                    value_name="건수"
                )
                fig = px.bar(org_melted, x=org_level, y="건수", color="서류",
                           title=f"서류별 미스캔 비교 (TOP 10)",
                           barmode="group")
                fig.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)
    
    # ── TAB 3: 상세 데이터 ──
    with tab3:
        st.header("📋 상세 데이터 조회")
        
        # 미스캔 필터
        show_type = st.radio(
            "표시 옵션",
            ["전체", "미스캔만", "완료만"],
            horizontal=True
        )
        
        df_display = df_sel.copy()
        
        if show_type == "미스캔만":
            df_display = df_display[
                (df_display["FA고지_c"] == "미스캔") |
                (df_display["비교설명_c"] == "미스캔") |
                ((df_display["완전판매_c"] == "미스캔") & (df_display["완전판매_c"] != "해당없음"))
            ]
        elif show_type == "완료만":
            df_display = df_display[
                (df_display["FA고지_c"] != "미스캔") &
                (df_display["비교설명_c"] != "미스캔") &
                ((df_display["완전판매_c"] != "미스캔") | (df_display["완전판매_c"] == "해당없음"))
            ]
        
        st.info(f"📊 표시 데이터: **{len(df_display):,}건**")
        
        # 데이터 테이블
        display_cols = [
            "보험시작일", "부문", "총괄", "부서", "영업가족", "담당자",
            "계약자", "증권번호", "보험료", "FA고지", "비교설명", "완전판매"
        ]
        available_cols = [col for col in display_cols if col in df_display.columns]
        
        st.dataframe(
            df_display[available_cols],
            use_container_width=True,
            height=500,
            hide_index=True
        )
        
        # 다운로드
        col1, col2 = st.columns(2)
        with col1:
            # Excel 다운로드
            from io import BytesIO
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_display.to_excel(writer, index=False, sheet_name='데이터')
            excel_data = output.getvalue()
            
            st.download_button(
                "📥 Excel 다운로드",
                excel_data,
                f"필터링데이터_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
        with col2:
            # CSV 다운로드
            csv_data = df_display.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                "📥 CSV 다운로드",
                csv_data,
                f"필터링데이터_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                "text/csv",
                use_container_width=True
            )
    
    # ── TAB 4: 안내 ──
    with tab4:
        st.header("ℹ️ 시스템 안내")
        
        st.markdown("""
        ### 📊 대시보드 사용 방법
        
        #### 1. 데이터 업데이트
        - GitHub 저장소에 `insurance_data.xlsx` 파일을 업로드/수정
        - Streamlit Cloud가 자동으로 재배포 (약 1-2분 소요)
        - 페이지가 자동으로 새로고침되며 최신 데이터 반영
        
        #### 2. 필터 사용
        - 왼쪽 사이드바에서 월/부문/총괄 선택
        - 여러 월을 선택하여 비교 가능
        - 조직별 상세 분석 가능
        
        #### 3. 데이터 다운로드
        - "상세 데이터" 탭에서 Excel/CSV 다운로드
        - 필터링된 데이터만 다운로드 가능
        
        #### 4. 자동 갱신
        - 5분마다 자동으로 데이터 갱신
        - 수동 새로고침: 🔄 버튼 클릭
        
        ### 📋 필수 서류 안내
        
        **책임판매 필수서류:**
        - 개인정보동의서
        - 비교설명확인서 
        - 고지의무확인서 (FA고지)
        - 완전판매확인서 (대상계약 限)
        
        ⚠️ 금융소비자보호법 및 보험업 감독규정에 따라 신계약 체결 전 반드시 완비되어야 하며,
        미비 시 리스크 계약으로 간주됩니다.
        
        ### 🔧 기술 정보
        - **플랫폼**: Streamlit Cloud
        - **데이터 소스**: GitHub (insurance_data.xlsx)
        - **업데이트**: 실시간 (GitHub 커밋 시 자동 반영)
        - **버전**: v5.0 (GitHub 기반 단순 버전)
        """)
        
        st.divider()
        st.caption(f"마지막 데이터 업데이트: {get_file_update_time()}")
        st.caption("© 2026 보험 서류 관리 시스템")

if __name__ == "__main__":
    main()
