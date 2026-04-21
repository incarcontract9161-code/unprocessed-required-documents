import streamlit as st
import pandas as pd
import os
import plotly.express as px
import plotly.graph_objects as go
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle, Image as RLImage, PageBreak, HRFlowable)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from datetime import datetime
import io

# ==========================================
# 0. 페이지 설정 (최상단 필수)
# ==========================================
st.set_page_config(page_title="보험 서류 스캔 관리 시스템", layout="wide", page_icon="📊")

# ==========================================
# 1. 전역 설정 & 안내 문구
# ==========================================
EXCEL_FILE = "insurance_data.xlsx"
APP_PASSWORD = os.environ.get("APP_PASSWORD", "incar961")

GUIDANCE_TEXT = (
    "【책임판매 필수서류 안내】\n"
    "개인정보동의서, 비교설명확인서, 고지의무확인서, 완전판매확인서(대상계약 限)는 "
    "금융소비자보호법 및 보험업 감독규정에 따라 신계약 체결 전 구비가 요구되는 필수 서류입니다. "
    "상기 서류는 소비자 보호 및 설명 의무 이행 여부를 확인하기 위한 내부 통제 관리 대상 서류로서, 실적 확정 입력 마감 시점까지 제출 완료를 원칙으로 하며 미비 시 내부 통제 리스크 관리 대상 계약으로 분류됩니다."
)
PRECAUTION_TEXT_COVER = (
    "【미처리 시 유의사항】\n"
    "실적 확정 입력 마감 시점까지 필수 서류가 제출되지 않은 계약과 조직에 대하여는 모집질서 및 분쟁  리스크 관리 대상으로 분류되어 관리됩니다.\n"
    "내부 통제 기준 충족 시까지,  내부 심사 및 결재 과정에서 승인 여부가 제한 될 수 있습니다. (리스크, 신규 운영자금 등 기타 지원 신청)"
)
PRECAUTION_TEXT_SHEET = "본인은 신계약 필수 서류의 사전 구비 의무 및 미제출 시 내부 통제 관리 기준이 적용될 수 있음을 확인합니다."
SIGNATURE_CONFIRMATION_TEXT = "신계약 필수 서류의 사전 구비 의무 및 미제출 시 내부 통제 관리 기준이 적용 사항을 영업가족에게 안내하였음을 확인합니다."

REQUIRED_DOCS_TABLE = [
    ["No.", "서류명", "법적 관리 근거 및 관련 내부 통제 기준", "목적 및 주요 내용"],
    ["1", "개인정보동의서", "개인정보보호법 15조 등", "개인정보 처리 적법 근거"],
    ["2", "비교설명확인서", "보험업감독규정", "유사 상품 비교 설명 이행 확인"],
    ["3", "고지의무확인서", "금융소비자보호법 26조", "중요사항 고지 및 소비자 오인 예방"],
    ["4", "완전판매확인서", "금소법 적합성 관련", "설명 의무 이행 증빙력 확보"]
]

# ==========================================
# 2. 데이터 로드 및 집계 로직 (스캔대상건 기준)
# ==========================================
@st.cache_data(ttl=300)
def load_data():
    if not os.path.exists(EXCEL_FILE): return pd.DataFrame()
    try:
        df = pd.read_excel(EXCEL_FILE)
        if df.empty: return pd.DataFrame()
        
        df["보험시작일_dt"] = pd.to_datetime(df["보험시작일"], errors="coerce")
        df["월"] = df["보험시작일_dt"].dt.to_period("M").astype(str)
        
        # 텍스트 정제
        for c in ["FA고지", "비교설명", "완전판매"]:
            df[f"{c}_v"] = df[c].fillna("").astype(str).str.strip()

        # [집계기준 변경 핵심]
        # 1. 미스캔 텍스트인 경우만 1로 체크
        df["fa_miss"] = (df["FA고지_v"] == "미스캔").astype(int)
        df["bi_miss"] = (df["비교설명_v"] == "미스캔").astype(int)
        df["wp_miss"] = (df["완전판매_v"] == "미스캔").astype(int)
        
        # 2. 스캔 대상 여부 (비교/고지는 필수 1, 완판은 해당사항없음이면 0)
        df["fa_tgt"] = 1
        df["bi_tgt"] = 1
        df["wp_tgt"] = (df["완전판매_v"] != "해당사항없음").astype(int)
        
        # 3. 신규 집계 컬럼
        df["대상스캔건"] = df["fa_tgt"] + df["bi_tgt"] + df["wp_tgt"]
        df["미스캔건"] = df["fa_miss"] + df["bi_miss"] + df["wp_miss"]
        df["완료스캔건"] = df["대상스캔건"] - df["미스캔건"]
        
        return df
    except Exception as e:
        st.error(f"데이터 로드 오류: {e}")
        return pd.DataFrame()

# ==========================================
# 3. 계층 리포트 및 지표 계산
# ==========================================
@st.cache_data(ttl=300)
def build_hierarchy_report(df, months=None):
    src = df[df["월"].isin(months)].copy() if months else df.copy()
    if src.empty: return pd.DataFrame()
    rows = []
    
    def get_agg_row(df_sub, gbn, names):
        t_s = int(df_sub["대상스캔건"].sum())
        c_s = int(df_sub["완료스캔건"].sum())
        m_s = int(df_sub["미스캔건"].sum())
        return {
            "구분": gbn, "부문": names[0], "총괄": names[1], "부서": names[2], "영업가족": names[3],
            "계약건수": len(df_sub),
            "대상스캔건": t_s,
            "완료스캔건": c_s,
            "미스캔합계": m_s,
            "스캔율": round(c_s / t_s * 100, 1) if t_s > 0 else 0.0,
            "FA미비": int(df_sub["fa_miss"].sum()),
            "비교미비": int(df_sub["bi_miss"].sum()),
            "완판미비": int(df_sub["wp_miss"].sum())
        }

    for bm, df_bm in src.groupby("부문"):
        rows.append(get_agg_row(df_bm, "부문계", (bm, "", "", "")))
        for tg, df_tg in df_bm.groupby("총괄"):
            rows.append(get_agg_row(df_tg, "총괄계", (bm, tg, "", "")))
            for ds, df_ds in df_tg.groupby("부서"):
                rows.append(get_agg_row(df_ds, "부서계", (bm, tg, ds, "")))
                for fg, df_fg in df_ds.groupby("영업가족"):
                    rows.append(get_agg_row(df_fg, "영업가족", (bm, tg, ds, fg)))
    return pd.DataFrame(rows)

# ==========================================
# 4. 외부 출력 리포트 생성 (Excel/PDF)
# ==========================================
def report_excel(df, months):
    wb = Workbook(); ws = wb.active; ws.title="계층별집계"
    tfn = "맑은 고딕"
    hf = Font(name=tfn, size=9, bold=True, color="FFFFFF")
    bf = Font(name=tfn, size=9)
    bdr = Border(left=Side("thin"), right=Side("thin"), top=Side("thin"), bottom=Side("thin"))
    fill = PatternFill("solid", fgColor="4472C4")
    
    headers = ["구분","부문","총괄","부서","영업가족","계약건수","대상스캔건","완료스캔건","미스캔합계","스캔율","FA미비","비교미비","완판미비"]
    for ci, h in enumerate(headers, 1):
        c = ws.cell(2, ci, h); c.font = hf; c.fill = fill; c.border = bdr; c.alignment = Alignment(horizontal="center")
        ws.column_dimensions[get_column_letter(ci)].width = 15
        
    rpt = build_hierarchy_report(df, months)
    for ri, (_, r) in enumerate(rpt.iterrows(), 3):
        row_v = [r[h] if h in r else "" for h in headers] # 지표 매핑 수정
        # 실제 데이터 필드명에 맞춰 수동 매핑
        vals = [r["구분"], r["부문"], r["총괄"], r["부서"], r["영업가족"], r["계약건수"], r["대상스캔건"], r["완료스캔건"], r["미스캔합계"], f"{r['스캔율']}%", r["FA미비"], r["비교미비"], r["완판미비"]]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(ri, ci, v); c.font = bf; c.border = bdr; c.alignment = Alignment(horizontal="center")
            
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf

@st.cache_resource
def register_korean_font():
    fn = "NotoSansKR"
    paths = ["/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc", "C:\\Windows\\Fonts\\malgun.ttf"]
    for p in paths:
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont(fn, p))
            return fn
    return "Helvetica"

# ==========================================
# 5. UI - 메인 대시보드
# ==========================================
def main():
    if not st.session_state.get("logged_in"):
        st.title("🔐 보험 서류 스캔 관리 시스템")
        pwd = st.text_input("접속 비밀번호", type="password")
        if st.button("로그인"):
            if pwd == APP_PASSWORD: st.session_state.logged_in = True; st.rerun()
            else: st.error("비밀번호 불일치")
        return

    df = load_data()
    if df.empty: st.warning("데이터가 없습니다."); return

    # 사이드바 필터
    with st.sidebar:
        st.header("⚙️ 필터 설정")
        all_mon = sorted(df["월"].unique())
        sel_mon = st.multiselect("분석 월", all_mon, default=all_mon[-1:])
        st.divider()
        if st.button("🚪 로그아웃"): st.session_state.logged_in = False; st.rerun()

    if not sel_mon: st.info("월을 선택하세요."); return
    df_f = df[df["월"].isin(sel_mon)]

    # KPI 대시보드
    st.title("🛡️ 스캔 미처리 관리 대시보드")
    
    t_s = df_f["대상스캔건"].sum()
    c_s = df_f["완료스캔건"].sum()
    m_s = df_f["미스캔건"].sum()
    rate = (c_s / t_s * 100) if t_s > 0 else 100.0

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("📄 총 대상스캔건", f"{t_s:,}개")
    k2.metric("✅ 완료 스캔건", f"{c_s:,}개")
    k3.metric("⚠️ 총 미스캔건", f"{m_s:,}개")
    k4.metric("📈 전체 스캔율", f"{rate:.1f}%")

    tab1, tab2, tab3 = st.tabs(["📊 현황 및 추이", "🏢 조직별 리포트", "📥 다운로드 섹션"])

    with tab1:
        c1, c2 = st.columns([2, 1])
        with c1:
            # 월별 추이
            mon_agg = df_f.groupby("월").agg(대상=("대상스캔건","sum"), 완료=("완료스캔건", "sum")).reset_index()
            mon_agg["스캔율"] = (mon_agg["완료"]/mon_agg["대상"]*100).round(1)
            fig = go.Figure()
            fig.add_trace(go.Bar(x=mon_agg["월"], y=mon_agg["대상"], name="대상건"))
            fig.add_trace(go.Scatter(x=mon_agg["월"], y=mon_agg["스캔율"], name="스캔율", yaxis="y2", line=dict(color="red")))
            fig.update_layout(yaxis2=dict(overlaying="y", side="right", range=[0, 110]), height=400, title="월별 스캔 현황 추이")
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            # 양식별 비중
            m_types = {"FA":df_f["fa_miss"].sum(), "비교":df_f["bi_miss"].sum(), "완판":df_f["wp_miss"].sum()}
            fig2 = px.pie(values=list(m_types.values()), names=list(m_types.keys()), hole=0.4, title="양식별 미비 비중")
            st.plotly_chart(fig2, use_container_width=True)

    with tab2:
        st.subheader("조직 계층별 상세 현황")
        agg_v = st.radio("집계 기준", ["부문", "총괄", "부서"], horizontal=True)
        agg_df = df_f.groupby(agg_v).agg(대상=("대상스캔건","sum"), 완료=("완료스캔건","sum"), 미스캔=("미스캔건","sum")).reset_index()
        agg_df["스캔율"] = (agg_df["완료"]/agg_df["대상"]*100).round(1)
        st.dataframe(agg_df.sort_values("미스캔", ascending=False).style.background_gradient(subset=["스캔율"], cmap="RdYlGn"), use_container_width=True)
        
        # 트리맵
        fig3 = px.treemap(agg_df, path=[agg_v], values="미스캔", color="스캔율", color_continuous_scale="RdYlGn", title=f"{agg_v}별 미스캔 분포")
        st.plotly_chart(fig3, use_container_width=True)

    with tab3:
        st.subheader("리포트 출력 및 데이터 추출")
        rpt_df = build_hierarchy_report(df, sel_mon)
        st.dataframe(rpt_df, use_container_width=True)
        
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            st.download_button("📥 계층 리포트 (Excel)", report_excel(df, sel_mon), f"ScanReport_{datetime.now().strftime('%Y%m%d')}.xlsx")
        with col_d2:
            st.info("준비된 PDF 템플릿으로 리포트를 생성할 수 있습니다.")

if __name__ == "__main__":
    main()