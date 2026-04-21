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
st.set_page_config(page_title="보험 서류 스캔 관리 대시보드", layout="wide", page_icon="📊")

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
PRECAUTION_TEXT_CONFIRM = "영업가족별 미처리 현황 및 유의사항에 대하여 인지하였으며,"
PRECAUTION_TEXT_SHEET = "본인은 신계약 필수 서류의 사전 구비 의무 및 미제출 시 내부 통제 관리 기준이 적용될 수 있음을 확인합니다."
SIGNATURE_CONFIRMATION_TEXT = "신계약 필수 서류의 사전 구비 의무 및 미제출 시 내부 통제 관리 기준이 적용 사항을 영업가족에게 안내하였음을 확인합니다."

REQUIRED_DOCS_TABLE = [
    ["No.", "서류명", "법적 관리 근거 및 관련 내부 통제 기준", "목적 및 주요 내용"],
    ["1", "개인정보동의서", "개인정보보호법 15조 및\n대리점 표준 내부통제기준 27조", "개인정보 처리 적법 근거, 보유계약 전산 관리 과정에\n따른 개인정보 처리로 신계약시 필수 징구"],
    ["2", "비교설명확인서", "보험업감독규정\n별표 5-6", "유사 상품 3개 이상 비교·설명 이행\n사실 고객 확인 서명"],
    ["3", "고지의무확인서", "금융소비자보호법 26조와\n동법시행령 24조", "판매자 중요사항 고지의무 이행 확인,\n권한·책임·보상 관련 핵심 사항 고지,\n소비자 오인 예방"],
    ["4", "완전판매확인서\n(대상: 종신, CI, CEO정기, 고액)", "금융소비자보호법 제17·19조 설명 적합성 적정성 관련 조항\n영업지원기준안", "약관,청약서 부본 제공, 중요 상품 이해 및\n자발적 가입 확인, 설명 의무 이행 증빙력 확보"]
]

# ==========================================
# 2. 데이터 로딩 및 집계 로직 구성 (핵심)
# ==========================================
@st.cache_data(ttl=300)
def load_data():
    if not os.path.exists(EXCEL_FILE):
        return pd.DataFrame()
    try:
        df = pd.read_excel(EXCEL_FILE)
        if df.empty: return pd.DataFrame()
        
        # 전처리
        df["보험시작일_dt"] = pd.to_datetime(df["보험시작일"], errors="coerce")
        df["월_피리어드"] = df["보험시작일_dt"].dt.to_period("M").astype(str)
        df["FA고지_c"] = df["FA고지"].fillna("").astype(str).str.strip()
        df["비교설명_c"] = df["비교설명"].fillna("").astype(str).str.strip()
        df["완전판매_c"] = df["완전판매"].fillna("").astype(str).str.strip()
        
        # [집계기준 변경: 스캔대상건]
        # 1. 미스캔 여부 (0: 스캔완료, 1: 미스캔)
        # 스캔, M스캔, 보험사스캔 등은 문자열이 '미스캔'이 아니므로 완료 처리됨
        df["FA_miss"] = (df["FA고지_c"] == "미스캔").astype(int)
        df["비교_miss"] = (df["비교설명_c"] == "미스캔").astype(int)
        df["완판_miss"] = (df["완전판매_c"] == "미스캔").astype(int)
        
        # 2. 스캔 대상 여부 판정
        df["FA_is_target"] = 1   # 비교설명/고지의무는 필수
        df["비교_is_target"] = 1
        # 완판은 '해당사항없음'일 경우 집계 대상에서 제외
        df["완판_is_target"] = (df["완전판매_c"] != "해당사항없음").astype(int)
        
        # 3. 대상스캔건 & 미스캔 합계
        df["대상스캔건"] = df["FA_is_target"] + df["비교_is_target"] + df["완판_is_target"]
        df["미스캔"] = df["FA_miss"] + df["비교_miss"] + df["완판_miss"]
        
        # 4. 완료스캔건
        df["완료스캔건"] = df["대상스캔건"] - df["미스캔"]
        
        return df
    except Exception as e:
        st.error(f"데이터 로드 오류: {e}")
        return pd.DataFrame()

# ==========================================
# 3. 집계 및 리포트 생성 함수
# ==========================================
@st.cache_data(ttl=300)
def build_hierarchy_report(df, months=None):
    src = df[df["월_피리어드"].isin(months)].copy() if months else df.copy()
    if src.empty: return pd.DataFrame()
    rows = []
    
    # 계층별 집계 로직
    levels = ["부문", "총괄", "부서", "영업가족"]
    for bm, df_bm in src.groupby("부문"):
        agg_cols = ["FA_miss", "비교_miss", "완판_miss", "미스캔", "대상스캔건", "완료스캔건"]
        def get_row_data(df_sub, gbn, name_tuple):
            data = df_sub[agg_cols].sum()
            cnt = len(df_sub)
            s_rate = round(data["완료스캔건"] / data["대상스캔건"] * 100, 1) if data["대상스캔건"] else 0.0
            m_rate = round(data["미스캔"] / data["대상스캔건"] * 100, 1) if data["대상스캔건"] else 0.0
            return {
                "구분": gbn, "부문": name_tuple[0], "총괄": name_tuple[1], "부서": name_tuple[2], "영업가족": name_tuple[3],
                "FA": int(data["FA_miss"]), "비교": int(data["비교_miss"]), "완판": int(data["완판_miss"]),
                "총미스캔": int(data["미스캔"]), "대상건": cnt, "대상스캔건": int(data["대상스캔건"]),
                "완료스캔건": int(data["완료스캔건"]), "스캔율": s_rate, "미스캔율": m_rate
            }
        
        rows.append(get_row_data(df_bm, "부문계", (bm, "", "", "")))
        for tg, df_tg in df_bm.groupby("총괄"):
            rows.append(get_row_data(df_tg, "총괄계", (bm, tg, "", "")))
            for ds, df_ds in df_tg.groupby("부서"):
                rows.append(get_row_data(df_ds, "부서계", (bm, tg, ds, "")))
                for fg, df_fg in df_ds.groupby("영업가족"):
                    rows.append(get_row_data(df_fg, "영업가족", (bm, tg, ds, fg)))
    
    return pd.DataFrame(rows)

# [엑셀 리포트 함수]
def report_excel(df, months):
    wb = Workbook(); ws = wb.active; ws.title="계층별_미처리현황"
    tfn = "맑은 고딕"
    hf = Font(name=tfn, size=9, bold=True, color="FFFFFF")
    bf = Font(name=tfn, size=9)
    bdr = Border(left=Side("thin"), right=Side("thin"), top=Side("thin"), bottom=Side("thin"))
    h_fill = PatternFill("solid", fgColor="4472C4")
    alt_fill = PatternFill("solid", fgColor="EEF3FB")
    
    headers = ["구분","부문","총괄","부서","영업가족","대상건","대상스캔건","완료스캔건","미스캔합계","스캔율","미스캔율","FA미스캔","비교미스캔","완판미스캔"]
    for ci, h in enumerate(headers, 1):
        c = ws.cell(2, ci, h); c.font = hf; c.fill = h_fill; c.border = bdr; c.alignment = Alignment(horizontal="center")
        ws.column_dimensions[get_column_letter(ci)].width = 15
    
    report = build_hierarchy_report(df, months)
    for ri, (_, row) in enumerate(report.iterrows(), 3):
        vals = [row["구분"], row["부문"], row["총괄"], row["부서"], row["영업가족"], row["대상건"], 
                row["대상스캔건"], row["완료스캔건"], row["총미스캔"], f"{row['스캔율']}%", f"{row['미스캔율']}%", 
                row["FA"], row["비교"], row["완판"]]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(ri, ci, v); c.font = bf; c.border = bdr; c.alignment = Alignment(horizontal="center")
            if ri % 2 == 0: c.fill = alt_fill
    
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf

# ==========================================
# 4. PDF 관련 스타일 및 폰트
# ==========================================
@st.cache_resource
def register_korean_font():
    font_candidates = [
        ("NotoSansKR", "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"),
        ("NanumGothic", "/usr/share/fonts/truetype/nanum/NanumGothic.ttf"),
        ("Malgun", r"C:\Windows\Fonts\malgun.ttf"),
        ("DejaVu", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
    ]
    for name, path in font_candidates:
        if os.path.exists(path):
            pdfmetrics.registerFont(TTFont(name, path))
            return name
    return "Helvetica"

def _pdf_styles(fn):
    S = getSampleStyleSheet()
    def ps(n, **kw): return ParagraphStyle(n, parent=S["Normal"], fontName=fn, **kw)
    return {
        "title": ps("T", fontSize=15, bold=True, alignment=1, spaceAfter=4),
        "sub": ps("S", fontSize=10, spaceAfter=2),
        "section": ps("SC", fontSize=9, bold=True, spaceAfter=2),
        "date": ps("D", fontSize=8, alignment=2),
        "notice": ps("N", fontSize=7.5, textColor=colors.HexColor("#CC0000")),
    }

def _tbl(data, cw, fn, align="CENTER"):
    cw_scaled = [w * 1.35 for w in cw]
    t = Table(data, colWidths=cw_scaled)
    t.setStyle(TableStyle([
        ("FONTNAME", (0,0), (-1,-1), fn), ("FONTSIZE", (0,0), (-1,-1), 8),
        ("GRID", (0,0), (-1,-1), 0.5, colors.grey), ("ALIGN", (0,0), (-1,-1), align),
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#DCE6F1")), ("VALIGN", (0,0), (-1,-1), "MIDDLE")
    ]))
    return t

# [관리대장 PDF 생성]
def ledger_pdf(dept_dict, period, df_full):
    fn, st_ = register_korean_font(), _pdf_styles(register_korean_font())
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, margin=(10*mm, 10*mm, 10*mm, 10*mm))
    elements = []
    
    for dept_name, grp in dept_dict.items():
        # 부서별 표지 및 영업가족별 고지서
        elements.append(Paragraph(f"보험 서류 미처리 관리대장 ({dept_name})", st_["title"]))
        elements.append(Paragraph(f"기간: {period}", st_["date"]))
        elements.append(Spacer(1, 10))
        
        # 상세 데이터 표
        hdr = [["영업가족", "대상건", "대상스캔", "완료스캔", "미스캔", "스캔율"]]
        rows = []
        for _, r in grp.iterrows():
            rows.append([r["영업가족"], f"{r['대상건']}건", f"{r['대상스캔건']}건", f"{r['완료스캔건']}건", f"{r['총미스캔']}건", f"{r['스캔율']}%"])
        elements.append(_tbl(hdr + rows, [80, 50, 50, 50, 50, 50], fn))
        elements.append(PageBreak())
        
    doc.build(elements); buf.seek(0); return buf

# ==========================================
# 5. UI 및 메인 실행
# ==========================================
def main():
    if not st.session_state.get("logged_in"):
        st.title("🔐 시스템 접속")
        pwd = st.text_input("비밀번호", type="password")
        if st.button("접속"):
            if pwd == APP_PASSWORD: st.session_state.logged_in = True; st.rerun()
            else: st.error("오류")
        return

    st.title("📊 보험 서류 스캔 통합 관리 (스캔대상건 기준)")
    df = load_data()
    if df.empty:
        st.warning("데이터 파일(insurance_data.xlsx)이 없습니다."); return

    # 분석 기간 필터
    all_months = sorted(df["월_피리어드"].dropna().unique())
    sel_months = st.multiselect("분석 월 선택", all_months, default=[all_months[-1]] if all_months else [])
    if not sel_months: return
    
    df_sel = df[df["월_피리어드"].isin(sel_months)]
    period_txt = f"{sel_months[0]}~{sel_months[-1]}" if len(sel_months)>1 else sel_months[0]

    # KPI 메트릭 (스캔대상건 기준)
    t_cnt = int(df_sel["대상스캔건"].sum())
    c_cnt = int(df_sel["완료스캔건"].sum())
    m_cnt = int(df_sel["미스캔"].sum())
    s_rate = round(c_cnt / t_cnt * 100, 1) if t_cnt else 100.0
    
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("📄 대상 스캔건수", f"{t_cnt:,}건")
    k2.metric("✅ 완료 스캔건수", f"{c_cnt:,}건")
    k3.metric("📈 스캔율", f"{s_rate}%")
    k4.metric("⚠️ 미스캔", f"{m_cnt:,}건")
    st.divider()

    tab1, tab2, tab3 = st.tabs(["📉 현황 대시보드", "📊 계층 리포트", "📋 관리대장"])

    with tab1:
        # 차트 및 검색
        agg_lv = st.selectbox("집계 단위", ["부문","총괄","부서","영업가족"])
        agg_df = df_sel.groupby(agg_lv).agg(스캔대상=("대상스캔건","sum"), 미스캔=("미스캔","sum")).reset_index()
        agg_df["스캔율"] = ((agg_df["스캔대상"]-agg_df["미스캔"])/agg_df["스캔대상"]*100).round(1)
        
        c1, c2 = st.columns(2)
        with c1:
            fig1 = px.bar(agg_df.sort_values("미스캔", ascending=False).head(20), x=agg_lv, y="미스캔", title="조직별 미스캔 현황 (TOP 20)", text="미스캔", color="미스캔", color_continuous_scale="Reds")
            st.plotly_chart(fig1, use_container_width=True)
        with c2:
            fig2 = px.treemap(agg_df, path=[agg_lv], values="스캔대상", color="스캔율", title="스캔 건수 비중 및 스캔율 분포", color_continuous_scale="RdYlGn")
            st.plotly_chart(fig2, use_container_width=True)

    with tab2:
        rpt = build_hierarchy_report(df, sel_months)
        st.dataframe(rpt.style.format({"스캔율":"{:.1f}%", "미스캔율":"{:.1f}%"}), use_container_width=True, hide_index=True)
        col1, col2 = st.columns(2)
        with col1:
            if st.button("📥 엑셀 다운로드"):
                st.download_button("Excel 받기", report_excel(df, sel_months), "report.xlsx")
        with col2:
            st.button("PDF 리포트 (준비중)")

    with tab3:
        # 부서별 필터링 후 관리대장
        ds_list = sorted(df_sel["부서"].unique())
        sel_ds = st.multiselect("출력 부서 선택", ds_list, default=ds_list[:1])
        if sel_ds:
            rpt_sub = rpt[(rpt["구분"]=="영업가족") & (rpt["부서"].isin(sel_ds))]
            st.write(f"선택 부서 관리 대상: {len(rpt_sub)}명")
            if st.button("📥 관리대장 PDF 생성"):
                targets = {d: rpt_sub[rpt_sub["부서"]==d] for d in sel_ds}
                st.download_button("PDF 다운로드", ledger_pdf(targets, period_txt, df_sel), "ledger.pdf")

if __name__ == "__main__":
    main()