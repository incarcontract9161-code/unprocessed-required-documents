import streamlit as st
import pandas as pd
import os
import plotly.express as px
import plotly.graph_objects as go
from openpyxl import Workbook
from datetime import datetime
import io

# ==========================================
# 0. 페이지 설정
# ==========================================
st.set_page_config(page_title="M스캔 전용 서류 처리 대시보드", layout="wide", page_icon="📊")

# ==========================================
# 1. 전역 설정 & 가이드 내용
# ==========================================
EXCEL_FILE = "insurance_data.xlsx"
APP_PASSWORD = os.environ.get("APP_PASSWORD", "incar961")
MANUAL_FILES = ["모바일동의_독려_안내.pdf", "모바일_보험가입확인서_장기_계피동일건발송절차_v2.pdf"]

GUIDANCE_DOCS = [
    ["No.", "서류명", "법적근거", "목적 및 주요내용"],
    ["1", "개인정보동의서", "개인정보보호법 15조\n대리점 표준 내부통제기준 27조", "개인정보 처리 적법 근거, 보유계약 전산 관리 과정에 따른 개인정보 처리로 신계약시 필수 징구"],
    ["2", "비교설명확인서", "보험업감독규정 별표 5-6", "유사 상품 3개 이상 비교·설명 이행 사실 고객 확인 서명"],
    ["3", "고지의무확인서", "금융소비자보호법 26조와 동법시행령 24조", "판매자 권한·책임·보상 관련 핵심 사항 고지, 소비자 오인 예방"],
    ["4", "완전판매확인서\n(대상: 종신, CI, CEO정기, 고액)", "금융소비자보호법 제17·19조\n영업지원기준안", "약관,청약서 부본 제공, 중요 상품 이해 및 자발적 가입 확인, 설명 의무 이행 증빙력 확보"]
]

MOBILE_GUIDE = {
    "title": "📱 모바일동의(M스캔) 집중 관리 안내",
    "reasons": [
        "**자동 매칭** : 실적 입력 시 서류 자동 연결, 수작업 업로드 불필요, 오류 감소",
        "**타임스탬프** : 전자서명 시점 실시간 기록, 계약 전 작성 객관적 증빙, 서명 위·변조 불가",
        "**누락 방지** : 필수 항목 입력 후 다음 단계 진행, 불완전판매 리스크 최소화",
        "**비용 절감** : 비교확인서 스캔 시 5년 원본 보관 의무 해소"
    ],
    "faq": [
        ("스캔 업로드도 가능한가요?", "네, 가능합니다. 다만 자동매칭·타임스탬프·누락방지 기능으로 업무 효율과 법적 보호가 강력한 모바일동의를 권장합니다."),
        ("완전판매확인서는 모든 계약에 필수인가요?", "영업지원기준안대로 종신, CI, CEO정기보험 및 고액 계약(저축성 300만원, 비저축성 100만원 이상)은 필수이며, 이외에는 금소법 취지에 따라 모든 계약에서 활용을 권장합니다."),
        ("모바일동의 절차는 어떻게 진행되나요?", "고객 동의 본인인증 → 설계사 동의 본인인증 → 전자서명 → 타임스탬프 기록 → 실적 입력 매칭\n*(법인, 미성년 계약 외 모바일 동의 전건 가능)*")
    ],
    "do_list": [
        "✅ 계약 체결 전 필수 서류 4종 100% 완비 (미완비 시 불완전판매 및 리스크 계약 간주)",
        "✅ 모바일동의 적극 활용 (자동매칭·타임스탬프 확보)",
        "✅ 사후징구 금지 (2026년 5월부터 서류 미비 시 내부 통제 미충족 조직으로 관리)",
        "✅ 고객 정보 확인 기록 및 적합성 판단 근거 남기기",
        "✅ 대리·중개 고지사항(권한·책임·보상) 사전 안내"
    ]
}

PROCESS_FLOW = [
    {"step": "1", "title": "보험구분 및 계/피동일 설정", "desc": "손보/생보 선택 후 계약자와 피보험자가 동일할 경우 '계/피동일' 체크박스 선택. 체크 시 피보험자 입력칸 자동 생략."},
    {"step": "2", "title": "완전판매확인서 발송 기준 설정", "desc": "필수발송대상: 생보보장성(변액/종신/100만↑), 생보저축성(300만↑), 손보보장/저축성(100만↑). 조건 외 선택적 발송 가능."},
    {"step": "3", "title": "계약자 결재", "desc": "페이퍼 화면 변경 → 동의함/부동의함 선택 → 서명 입력 → 다음 클릭 → 마지막 페이지 상단 저장 클릭."},
    {"step": "4", "title": "피보험자 결재 (계피상이 시)", "desc": "계약자 결재 완료 시 자동 문자 발송 → 화면 확대 가능 → 계약자 결재와 동일한 순서로 진행."},
    {"step": "5", "title": "설계사 결재", "desc": "계약자/피보험자 결재 완료 시 자동 문자 발송 → 글씨 진하게 선택(필요시) → 다음 클릭 → 저장 완료."}
]

# ==========================================
# 2. 데이터 로딩 (NAType 오류 해결 적용)
# ==========================================
@st.cache_data(ttl=300)
def load_data():
    if not os.path.exists(EXCEL_FILE):
        st.error(f"⚠️ '{EXCEL_FILE}' 파일이 없습니다.")
        return pd.DataFrame()
    try:
        df = pd.read_excel(EXCEL_FILE)
        if df.empty: return pd.DataFrame()

        df["보험시작일_dt"] = pd.to_datetime(df["보험시작일"], errors="coerce")
        df["월_피리어드"] = df["보험시작일_dt"].dt.to_period("M").astype(str)
        
        for col in ["FA고지", "비교설명", "완전판매"]:
            df[f"{col}_c"] = df[col].fillna(" ").astype(str).str.strip()

        def is_total_scanned(val):
            if pd.isna(val) or val == " ": return False
            return str(val).strip() in ["스캔", "M스캔", "보험사스캔"]

        def is_m_scanned(val):
            if pd.isna(val) or val == " ": return False
            return str(val).strip() == "M스캔"

        def is_cs_target(val):
            if pd.isna(val) or val == " ": return False
            return str(val).strip() in ["스캔", "M스캔", "미스캔"]

        df["FA_전체스캔"] = df["FA고지_c"].apply(is_total_scanned).astype(int)
        df["비교_전체스캔"] = df["비교설명_c"].apply(is_total_scanned).astype(int)
        df["완판_전체스캔"] = df["완전판매_c"].apply(is_total_scanned).astype(int)

        df["FA_M스캔"] = df["FA고지_c"].apply(is_m_scanned).astype(int)
        df["비교_M스캔"] = df["비교설명_c"].apply(is_m_scanned).astype(int)
        df["완판_M스캔"] = df["완전판매_c"].apply(is_m_scanned).astype(int)

        df["완판_대상"] = df["완전판매_c"].apply(is_cs_target).astype(int)
        df["FA_target"] = 1
        df["비교_target"] = 1
        df["완판_target"] = df["완판_대상"]

        df["대상건"] = df[["FA_target", "비교_target", "완판_target"]].sum(axis=1).astype(int)
        df["전체스캔건"] = df[["FA_전체스캔", "비교_전체스캔", "완판_전체스캔"]].sum(axis=1).astype(int)
        df["M스캔건"] = df[["FA_M스캔", "비교_M스캔", "완판_M스캔"]].sum(axis=1).astype(int)

        # ✅ NAType round 오류 해결: pd.NA → float('nan') 변경
        # float('nan')은 round() 및 fillna()와 100% 호환됩니다.
        df["전체스캔율"] = (df["전체스캔건"] / df["대상건"].replace(0, float('nan')) * 100).round(1).fillna(0.0)
        df["M스캔율"] = (df["M스캔건"] / df["전체스캔건"].replace(0, float('nan')) * 100).round(1).fillna(0.0)

        return df
    except Exception as e:
        st.error(f"❌ 엑셀 파일 읽기 오류: {e}")
        return pd.DataFrame()

def get_file_update_time():
    if os.path.exists(EXCEL_FILE):
        return datetime.fromtimestamp(os.path.getmtime(EXCEL_FILE)).strftime("%Y-%m-%d %H:%M:%S")
    return "알 수 없음"

# ==========================================
# 3. 집계 헬퍼
# ==========================================
@st.cache_data(ttl=300)
def build_org_stats(df, months=None, group_col="영업가족"):
    src = df[df["월_피리어드"].isin(months)].copy() if months else df.copy()
    if src.empty: return pd.DataFrame()

    agg = src.groupby(group_col).agg(
        대상건=("대상건", "sum"),
        전체스캔건=("전체스캔건", "sum"),
        M스캔건=("M스캔건", "sum")
    ).reset_index()
    agg = agg.rename(columns={group_col: "조직"})
    
    # ✅ 집계 함수 내 동일 적용
    agg["전체스캔율"] = (agg["전체스캔건"] / agg["대상건"].replace(0, float('nan')) * 100).round(1).fillna(0.0)
    agg["M스캔율"] = (agg["M스캔건"] / agg["전체스캔건"].replace(0, float('nan')) * 100).round(1).fillna(0.0)
    agg["순위"] = range(1, len(agg) + 1)
    
    return agg.sort_values("M스캔건", ascending=False).reset_index(drop=True)

# ==========================================
# 4. UI – 로그인 & 대시보드
# ==========================================
def login_page():
    st.title("🔐 시스템 접속")
    pwd = st.text_input("접속 비밀번호를 입력하세요", type="password")
    if st.button("접속하기", use_container_width=True, type="primary"):
        if pwd == APP_PASSWORD:
            st.session_state.logged_in = True
            st.rerun()
        else:
            st.error("비밀번호가 올바르지 않습니다.")

def dashboard_page():
    st.title("📱 M스캔 전용 서류 처리 현황 대시보드")
    
    with st.expander("📌 책임판매 필수서류 4종 & 모바일동의 집중 관리 안내 (2026.05 시행)", expanded=True):
        st.markdown("✅ **체결 전 완비 원칙** : 4종 중 단 1개라도 미완비 시 불완전판매 및 리스크 계약 간주\n"
                    "📱 **모바일동의 표준화** : 자동매칭·타임스탬프·누락방지 기능으로 업무 효율과 법적 증빙력 확보\n"
                    "⚠️ **사후징구 금지** : 2026년 5월부터 서류 미비 계약은 내부 통제 미충족 조직으로 관리")

    df = load_data()
    if df.empty:
        st.warning("데이터가 없습니다. insurance_data.xlsx 파일을 확인해주세요.")
        return

    col1, col2 = st.columns([2, 1])
    with col1: st.success(f"총 {len(df):,}건의 데이터 로드 완료")
    with col2: st.info(f"기준: {get_file_update_time()}")

    month_col = "월_피리어드"
    all_months = sorted(df[month_col].dropna().unique())
    st.subheader("분석 기간 선택")
    sel_months = st.multiselect("월 선택", all_months, default=[all_months[-1]] if all_months else [])
    if not sel_months: st.warning("최소 1개 이상의 월을 선택해주세요."); return

    df_sel = df[df[month_col].isin(sel_months)].copy()
    if df_sel.empty: st.info("선택한 기간에 데이터가 없습니다."); return

    # 🟦 KPI 카드
    total_docs = int(df_sel["대상건"].sum())
    total_scanned = int(df_sel["전체스캔건"].sum())
    m_scanned = int(df_sel["M스캔건"].sum())
    total_rate = round(total_scanned / total_docs * 100, 1) if total_docs else 0.0
    m_rate = round(m_scanned / total_scanned * 100, 1) if total_scanned else 0.0

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("총 대상건", f"{total_docs:,}건", help="필수 징구 대상 서류 총합")
    m2.metric("전체 스캔건", f"{total_scanned:,}건", delta=f"{total_rate:.1f}% (완료율)", delta_color="normal")
    m3.metric("M스캔건", f"{m_scanned:,}건", delta=f"{m_rate:.1f}% (M스캔 비중)", delta_color="inverse" if m_rate < 50 else "normal")
    m4.metric("미처리 건", f"{total_docs - total_scanned:,}건")
    
    st.divider()

    tab_dash, tab_map, tab_guide, tab_manual = st.tabs(["📊 현황 대시보드", "🗺️ M스캔 활용 현황", "📱 가이드 & 프로세스", "📚 매뉴얼 다운로드"])

    # ==========================================
    # 탭 1: 현황 대시보드
    # ==========================================
    with tab_dash:
        cs1, cs2 = st.columns([2, 1])
        with cs1: search_text = st.text_input("조직 검색", placeholder="조직명 입력")
        with cs2: agg_group = st.selectbox("집계 기준", ["부문", "총괄", "부서", "영업가족"], key="agg_group")

        agg = build_org_stats(df_sel, sel_months, agg_group)
        if search_text: agg = agg[agg["조직"].astype(str).str.contains(search_text, case=False, na=False)]
        if agg.empty: st.info("조건에 맞는 데이터가 없습니다."); return

        rate_option = st.radio(
            "📊 표시할 지표 선택",
            ["전체스캔율 (대상건 대비)", "M스캔율 (전체스캔건 대비)"],
            horizontal=True,
            key="rate_option"
        )
        
        is_total_rate = "전체스캔율" in rate_option
        rate_col = "전체스캔율" if is_total_rate else "M스캔율"
        rate_label = "전체스캔율" if is_total_rate else "M스캔율"
        
        st.dataframe(
            agg[["순위", "조직", "대상건", "전체스캔건", "M스캔건", "전체스캔율", "M스캔율"]]
            .style.format({"전체스캔율":"{:.1f}%", "M스캔율":"{:.1f}%"})
            .highlight_max(subset=[rate_col], color="#e6f0fa"),
            use_container_width=True, hide_index=True
        )
        
        top_n = st.slider("차트 표시 개수", 5, 30, 15, key="dash_top_n")
        top = agg.head(top_n)

        c1, c2 = st.columns(2)
        with c1:
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=top["조직"], y=top[rate_col],
                mode="lines+markers",
                line=dict(width=3, color="#2ECC71" if is_total_rate else "#3498DB"),
                marker=dict(size=8),
                name=rate_label,
                text=[f"{v:.1f}%" for v in top[rate_col]],
                textposition="top center"
            ))
            fig.update_layout(
                xaxis_tickangle=-45, height=400, 
                title=f"조직별 {rate_label} (M스캔건 상위 {top_n}개)",
                yaxis_range=[0, 100], yaxis_title="비율(%)",
                legend=dict(orientation="h", y=1.12)
            )
            st.plotly_chart(fig, use_container_width=True)

        with c2:
            fig2 = go.Figure()
            fig2.add_trace(go.Bar(
                x=top["조직"], y=top["M스캔건"],
                marker_color="#FF6B6B",
                name="M스캔건",
                texttemplate="%{y:,.0f}", textposition="outside"
            ))
            fig2.update_layout(
                xaxis_tickangle=-45, height=400,
                title="조직별 M스캔 건수", yaxis_title="건수",
                showlegend=False
            )
            st.plotly_chart(fig2, use_container_width=True)

    # ==========================================
    # 탭 2: M스캔 활용 현황 (단일 막대그래프)
    # ==========================================
    with tab_map:
        st.subheader("조직별 M스캔 활용도 분포")
        mc1, mc2 = st.columns([1, 2])
        with mc1: map_level = st.selectbox("집계 단위", ["부문", "총괄", "부서", "영업가족"], key="map_level")
        with mc2: st.caption("📊 M스캔 건수 기준 내림차순 정렬")

        map_agg = build_org_stats(df_sel, sel_months, map_level)
        map_agg = map_agg[map_agg["M스캔건"] > 0].reset_index(drop=True)
        
        if map_agg.empty:
            st.info("M스캔 활용 데이터가 없습니다.")
        else:
            fig_bar = px.bar(
                map_agg, y="조직", x="M스캔율", orientation="h",
                color="M스캔건", text_auto=".1f%",
                color_continuous_scale="YlOrRd",
                title="전체스캔건 중 M스캔 비중 및 건수 분포"
            )
            fig_bar.update_layout(
                height=600, 
                xaxis_title="M스캔율 (%)",
                yaxis=dict(autorange="reversed")
            )
            st.plotly_chart(fig_bar, use_container_width=True)

    # ==========================================
    # 탭 3: 가이드 & 프로세스
    # ==========================================
    with tab_guide:
        g1, g2 = st.columns(2)
        with g1:
            st.subheader(MOBILE_GUIDE["title"])
            for reason in MOBILE_GUIDE["reasons"]: st.markdown(f"🔹 {reason}")
            st.divider()
            st.subheader("📝 책임판매 필수 서류 4종")
            st.dataframe(pd.DataFrame(GUIDANCE_DOCS[1:], columns=GUIDANCE_DOCS[0]).set_index("No."), 
                         use_container_width=True, hide_index=True)
            st.divider()
            st.subheader("✅ 반드시 해야 할 일(Do)")
            for do_item in MOBILE_GUIDE["do_list"]: st.markdown(do_item)

        with g2:
            st.subheader("🔄 모바일가입확인서 발송 및 결재 프로세스")
            for step in PROCESS_FLOW:
                with st.expander(f"🔹 Step {step['step']}: {step['title']}"):
                    st.markdown(step["desc"])
            st.divider()
            st.subheader("❓ 자주 묻는 질문(FAQ)")
            for q, a in MOBILE_GUIDE["faq"]: st.markdown(f"**Q. {q}**\n\nA. {a}")

    # ==========================================
    # 탭 4: 매뉴얼 다운로드
    # ==========================================
    with tab_manual:
        st.subheader("📚 모바일동의 매뉴얼 다운로드")
        st.markdown("첨부된 가이드 문서를 PDF로 다운로드할 수 있습니다.")
        st.divider()
        
        found = False
        for mf in MANUAL_FILES:
            if os.path.exists(mf):
                found = True
                try:
                    with open(mf, "rb") as f:
                        st.download_button(
                            label=f"📥 {mf}",
                            data=f.read(),
                            file_name=mf,
                            mime="application/pdf",
                            key=f"dl_{mf}",
                            use_container_width=True
                        )
                except Exception as e:
                    st.error(f"❌ {mf} 읽기 오류: {e}")
        if not found:
            st.warning("⚠️ 매뉴얼 파일이 현재 폴더에 없습니다. 실행 디렉토리에 PDF 파일을 복사해주세요.")

    # ==========================================
    # 다운로드 버튼
    # ==========================================
    st.divider()
    st.subheader("📥 리포트 내보내기")
    if st.button("현황 데이터 Excel 다운로드", use_container_width=True):
        wb = Workbook(); ws = wb.active; ws.title = "M스캔 현황"
        ws.append(["순위", "조직", "대상건", "전체스캔건", "M스캔건", "전체스캔율", "M스캔율"])
        for _, row in agg.iterrows():
            ws.append([row["순위"], row["조직"], int(row["대상건"]), int(row["전체스캔건"]), 
                       int(row["M스캔건"]), f"{row['전체스캔율']:.1f}%", f"{row['M스캔율']:.1f}%"])
        buf = io.BytesIO(); wb.save(buf); buf.seek(0)
        st.download_button("Excel 저장", buf, f"M스캔_현황_{datetime.now().strftime('%Y%m%d')}.xlsx", 
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ==========================================
# 5. Main
# ==========================================
def main():
    if not st.session_state.get("logged_in"):
        login_page()
    else:
        with st.sidebar:
            st.success("👋 접속 완료")
            if st.button("🚪 로그아웃", use_container_width=True):
                st.session_state.logged_in = False; st.rerun()
            st.divider()
            st.caption("v9.1 | M스캔 전용 집계 | NAType round 오류 해결 완료")
        dashboard_page()

if __name__ == "__main__":
    main()